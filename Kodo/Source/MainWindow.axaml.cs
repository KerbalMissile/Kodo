// Licensed under the Kodo Public License v1.0
// April 19th, 2026 - KerbalMissile - Changed "One file at a time" note to "No file open"
// April 19th, 2026 - KerbalMissile - Added proper comments
// April 19th, 2026 - KerbalMissile - Changed "No File Open" at top bar to be empty when no file is open, and to show the file name when a file is open, also shows "unsaved" if there are unsaved changes
// April 19th, 2026 - KerbalMissile - Changed open file icon to fit in better
// April 19th, 2026 - KerbalMissile - Re-added SS-YYC's changes to improve Kodo's UI, did some changes the New File buttons but they still do not work
// April 20th, 2026 - SS-YYC - Added collapsible file explorer panel with folder tree and expand/collapse
// April 20th, 2026 - KerbalMissile - Added extension support, re-added full screen by default, updated open / close folder icon
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Styling;
using Avalonia.Threading;
using AvaloniaEdit.Highlighting;


namespace Kodo;

// Represents a single row in the file explorer tree
public class FileTreeItem : INotifyPropertyChanged
{
    private bool _isExpanded;

    public string Name { get; init; } = string.Empty;
    public string FullPath { get; init; } = string.Empty;
    public bool IsDirectory { get; init; }
    public int Depth { get; init; }

    // Pixel indentation based on nesting depth
    public double IndentWidth => Depth * 14.0;

    // Chevron shown next to directories; blank for files
    public string ChevronText => IsDirectory ? (_isExpanded ? "▾" : "▸") : string.Empty;

    public bool IsExpanded
    {
        get => _isExpanded;
        set
        {
            if (_isExpanded == value) return;
            _isExpanded = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(ChevronText));
            OnPropertyChanged(nameof(Icon));
        }
    }

    // Icon varies between open/closed folder vs file
    public string Icon => IsDirectory ? (_isExpanded ? "❒" : "❑") : GetFileIcon(Name);

    // Muted colour for directories, normal text colour for files – resolved in AXAML via binding
    public string NameColor => IsDirectory ? "#A0A0A0" : "#F4F4F4";
    public string MutedTextBrush => "#A0A0A0";

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    // Returns a simple file-type icon based on extension
    private static string GetFileIcon(string fileName)
    {
        var ext = Path.GetExtension(fileName).ToLowerInvariant();
        return ext switch
        {
            ".cs" => "📄",
            ".axaml" or ".xaml" or ".xml" or ".html" or ".htm" => "📋",
            ".json" or ".yaml" or ".yml" or ".toml" => "📝",
            ".md" or ".txt" or ".rst" => "📃",
            ".png" or ".jpg" or ".jpeg" or ".gif" or ".svg" or ".ico" => "🖼",
            ".py" => "🐍",
            ".js" or ".ts" or ".jsx" or ".tsx" => "📜",
            ".css" or ".scss" or ".less" => "🎨",
            ".sh" or ".bat" or ".ps1" => "⚡",
            ".zip" or ".tar" or ".gz" or ".rar" => "📦",
            _ => "📄",
        };
    }
}

public class LoadedExtension : INotifyPropertyChanged
{
    public string Id { get; init; } = string.Empty;
    public string Version { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public string Type { get; init; } = string.Empty;
    public string Author { get; init; } = string.Empty;
    public string Description { get; init; } = string.Empty;
    public string[] Extensions { get; init; } = [];

    public string[] Keywords { get; set; } = [];
    public string[] Types { get; set; } = [];
    public string CommentLine { get; set; } = "//";
    public string CommentBlockStart { get; set; } = "/*";
    public string CommentBlockEnd { get; set; } = "*/";
    public Dictionary<string, string> ColorTokens { get; set; } = new();

    public IBrush AccentBrush { get; set; } = Brush.Parse("#0E639C");
    public IBrush CardBrush { get; set; } = Brush.Parse("#252526");
    public IBrush PrimaryTextBrush { get; set; } = Brush.Parse("#F4F4F4");
    public IBrush SurfaceBorderBrush { get; set; } = Brush.Parse("#2B2B2B");
    public IBrush MutedTextBrush { get; set; } = Brush.Parse("#A0A0A0");

    public event PropertyChangedEventHandler? PropertyChanged;

    public void NotifyAllBrushesChanged()
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(AccentBrush)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CardBrush)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(PrimaryTextBrush)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SurfaceBorderBrush)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(MutedTextBrush)));
    }
}

public partial class MainWindow : Window, INotifyPropertyChanged
{
    private string? _currentFilePath;
    private string? _currentFolderPath;
    private string _editorContent = string.Empty;
    private bool _isDirty;
    private bool _hasUntitledDocument;
    private bool _isSettingsPageVisible;
    private bool _isExtensionsPageVisible;
    private bool _isFileExplorerVisible;
    private bool _suppressDirtyTracking;
    private string _currentThemeName = "Dark";
    private string _editorStatsText = "0 lines";
    private event PropertyChangedEventHandler? ViewModelPropertyChanged;

    // Flat list that backs the ItemsControl – directories insert/remove their children in-place
    public ObservableCollection<FileTreeItem> FileTreeItems { get; } = new();
    public ObservableCollection<LoadedExtension> LoadedExtensions { get; } = new();

    // Collects human-readable errors from the last LoadExtensions() call
    public ObservableCollection<string> ExtensionLoadErrors { get; } = new();

    private string ExtensionsFolderPath => Path.Combine(Directory.GetCurrentDirectory(), "Extensions");
    private LoadedExtension? _currentLanguageExtension;
    public LoadedExtension? CurrentLanguageExtension
    {
        get => _currentLanguageExtension;
        private set
        {
            if (_currentLanguageExtension == value) return;
            _currentLanguageExtension = value;
            OnPropertyChanged();
        }
    }

    public MainWindow()
    {
        InitializeComponent();
        DataContext = this;
        // TextEditor uses EventHandler, not RoutedEventHandler, so hook up in code
        EditorTextBox.TextChanged += EditorTextBox_OnTextChanged;
        ApplyTheme("Dark");
        EnsureExtensionsFolder();
        LoadExtensions();
        RefreshState();
    }

    private void EnsureExtensionsFolder()
    {
        var folder = ExtensionsFolderPath;
        if (!Directory.Exists(folder))
        {
            Directory.CreateDirectory(folder);
        }
    }

    private void LoadExtensions()
    {
        LoadedExtensions.Clear();
        ExtensionLoadErrors.Clear();
        
        var searchPaths = new List<string>();
        
        var binExtensionsPath = Path.Combine(Directory.GetCurrentDirectory(), "Extensions");
        searchPaths.Add(binExtensionsPath);
        
        var projectRoot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..");
        var srcExtensionsPath = Path.GetFullPath(Path.Combine(projectRoot, "Extensions"));
        if (!string.Equals(srcExtensionsPath, binExtensionsPath, StringComparison.OrdinalIgnoreCase))
            searchPaths.Add(srcExtensionsPath);

        bool anyFolderFound = false;
        foreach (var searchPath in searchPaths)
        {
            if (!Directory.Exists(searchPath)) continue;
            anyFolderFound = true;
            
            foreach (var koxFile in Directory.GetFiles(searchPath, "*.kox"))
            {
                try
                {
                    var ext = LoadExtensionFromKox(koxFile);
                    if (ext is not null && !LoadedExtensions.Any(e => e.Id == ext.Id))
                        LoadedExtensions.Add(ext);
                }
                catch (Exception ex)
                {
                    ExtensionLoadErrors.Add($"Failed to load '{Path.GetFileName(koxFile)}': {ex.Message}");
                }
            }
            
            foreach (var dir in Directory.GetDirectories(searchPath))
            {
                try
                {
                    var manifestPath = Path.Combine(dir, "manifest.json");
                    if (File.Exists(manifestPath))
                    {
                        var ext = LoadExtensionFromFolder(dir);
                        if (ext is not null && !LoadedExtensions.Any(e => e.Id == ext.Id))
                            LoadedExtensions.Add(ext);
                    }
                }
                catch (Exception ex)
                {
                    ExtensionLoadErrors.Add($"Failed to load folder extension '{Path.GetFileName(dir)}': {ex.Message}");
                }
            }
        }

        if (!anyFolderFound)
            ExtensionLoadErrors.Add($"Extensions folder not found. Expected: {binExtensionsPath}");

        OnPropertyChanged(nameof(ExtensionLoadErrors));
        RefreshExtensionTheme();
    }

    private LoadedExtension? LoadExtensionFromFolder(string folderPath)
    {
        var manifestPath = Path.Combine(folderPath, "manifest.json");
        var languagePath = Path.Combine(folderPath, "language.json");
        
        if (!File.Exists(manifestPath)) return null;
        
        var manifestJson = File.ReadAllText(manifestPath);
        using var manifestDoc = JsonDocument.Parse(manifestJson);
        var manifest = manifestDoc.RootElement;

        var ext = new LoadedExtension
        {
            Id = manifest.TryGetProperty("id", out var id) ? id.GetString() ?? "" : "",
            Version = manifest.TryGetProperty("version", out var ver) ? ver.GetString() ?? "" : "",
            Name = manifest.TryGetProperty("name", out var name) ? name.GetString() ?? "" : "",
            Type = manifest.TryGetProperty("type", out var type) ? type.GetString() ?? "" : "",
            Author = manifest.TryGetProperty("author", out var author) ? author.GetString() ?? "" : "",
            Description = manifest.TryGetProperty("description", out var desc) ? desc.GetString() ?? "" : "",
            Extensions = manifest.TryGetProperty("extensions", out var exts) ? exts.EnumerateArray().Select(e => e.GetString() ?? "").ToArray() : []
        };

        if (File.Exists(languagePath))
        {
            var langJson = File.ReadAllText(languagePath);
            using var langDoc = JsonDocument.Parse(langJson);
            var lang = langDoc.RootElement;

            ext.Keywords = lang.TryGetProperty("keywords", out var kw) ? kw.EnumerateArray().Select(e => e.GetString() ?? "").ToArray() : [];
            ext.Types = lang.TryGetProperty("types", out var tp) ? tp.EnumerateArray().Select(e => e.GetString() ?? "").ToArray() : [];
            ext.CommentLine = lang.TryGetProperty("commentLine", out var cl) ? cl.GetString() ?? "//" : "//";
            ext.CommentBlockStart = lang.TryGetProperty("commentBlockStart", out var cbs) ? cbs.GetString() ?? "/*" : "/*";
            ext.CommentBlockEnd = lang.TryGetProperty("commentBlockEnd", out var cbe) ? cbe.GetString() ?? "*/" : "*/";

            if (lang.TryGetProperty("colorTokens", out var ct))
            {
                foreach (var prop in ct.EnumerateObject())
                    ext.ColorTokens[prop.Name] = prop.Value.GetString() ?? "#FFFFFF";
            }
        }

        return ext;
    }

    private void RefreshExtensionTheme()
    {
        foreach (var ext in LoadedExtensions)
        {
            ext.AccentBrush = AccentBrush;
            ext.CardBrush = CardBrush;
            ext.PrimaryTextBrush = PrimaryTextBrush;
            ext.SurfaceBorderBrush = SurfaceBorderBrush;
            ext.MutedTextBrush = MutedTextBrush;
            ext.NotifyAllBrushesChanged();
        }
    }

    private LoadedExtension? LoadExtensionFromKox(string koxPath)
    {
        using var archive = ZipFile.OpenRead(koxPath);
        var manifestEntry = archive.GetEntry("manifest.json");
        var languageEntry = archive.GetEntry("language.json");
        
        if (manifestEntry is null) return null;

        using var manifestStream = manifestEntry.Open();
        using var manifestReader = new StreamReader(manifestStream);
        var manifestJson = manifestReader.ReadToEnd();
        using var manifestDoc = JsonDocument.Parse(manifestJson);
        var manifest = manifestDoc.RootElement;

        var ext = new LoadedExtension
        {
            Id = manifest.TryGetProperty("id", out var id) ? id.GetString() ?? "" : "",
            Version = manifest.TryGetProperty("version", out var ver) ? ver.GetString() ?? "" : "",
            Name = manifest.TryGetProperty("name", out var name) ? name.GetString() ?? "" : "",
            Type = manifest.TryGetProperty("type", out var type) ? type.GetString() ?? "" : "",
            Author = manifest.TryGetProperty("author", out var author) ? author.GetString() ?? "" : "",
            Description = manifest.TryGetProperty("description", out var desc) ? desc.GetString() ?? "" : "",
            Extensions = manifest.TryGetProperty("extensions", out var exts) ? exts.EnumerateArray().Select(e => e.GetString() ?? "").ToArray() : []
        };

        if (languageEntry is not null)
        {
            using var langStream = languageEntry.Open();
            using var langReader = new StreamReader(langStream);
            var langJson = langReader.ReadToEnd();
            using var langDoc = JsonDocument.Parse(langJson);
            var lang = langDoc.RootElement;

            ext.Keywords = lang.TryGetProperty("keywords", out var kw) ? kw.EnumerateArray().Select(e => e.GetString() ?? "").ToArray() : [];
            ext.Types = lang.TryGetProperty("types", out var tp) ? tp.EnumerateArray().Select(e => e.GetString() ?? "").ToArray() : [];
            ext.CommentLine = lang.TryGetProperty("commentLine", out var cl) ? cl.GetString() ?? "//" : "//";
            ext.CommentBlockStart = lang.TryGetProperty("commentBlockStart", out var cbs) ? cbs.GetString() ?? "/*" : "/*";
            ext.CommentBlockEnd = lang.TryGetProperty("commentBlockEnd", out var cbe) ? cbe.GetString() ?? "*/" : "*/";

            if (lang.TryGetProperty("colorTokens", out var ct))
            {
                foreach (var prop in ct.EnumerateObject())
                    ext.ColorTokens[prop.Name] = prop.Value.GetString() ?? "#FFFFFF";
            }
        }

        return ext;
    }

    public string EditorContent
    {
        get => _editorContent;
        set
        {
            if (_editorContent == value) return;
            _editorContent = value;
            OnPropertyChanged();
        }
    }

    public bool IsSettingsPageVisible
    {
        get => _isSettingsPageVisible;
        set
        {
            if (_isSettingsPageVisible == value) return;
            _isSettingsPageVisible = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsEditorPageVisible));
        }
    }

    public bool IsFileExplorerVisible
    {
        get => _isFileExplorerVisible;
        private set
        {
            if (_isFileExplorerVisible == value) return;
            _isFileExplorerVisible = value;
            OnPropertyChanged();
        }
    }

    public bool IsExtensionsPageVisible
    {
        get => _isExtensionsPageVisible;
        set
        {
            if (_isExtensionsPageVisible == value) return;
            _isExtensionsPageVisible = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsEditorPageVisible));
        }
    }

    public bool IsEditorPageVisible => !IsSettingsPageVisible && !IsExtensionsPageVisible;

    public bool HasDocumentOpen => _currentFilePath is not null || _hasUntitledDocument;

    public bool HasFileOpen => _currentFilePath is not null;

    public bool IsFolderOpen => _currentFolderPath is not null;

    public bool IsEmptyStateVisible => !HasDocumentOpen;

    public string CurrentThemeName
    {
        get => _currentThemeName;
        private set
        {
            if (_currentThemeName == value) return;
            _currentThemeName = value;
            OnPropertyChanged();
        }
    }

    // Displays the file name and unsaved status in the top bar
    public string FileSummaryText => HasDocumentOpen
        ? $"{GetDocumentDisplayName()}{(_isDirty ? " • unsaved" : string.Empty)}"
        : "Open A File";

    public string FilePathText => HasFileOpen
        ? _currentFilePath!
        : HasDocumentOpen ? "Unsaved file"
        : IsFolderOpen ? $"📂 {_currentFolderPath}"
        : "No file open";

    // Header shown at the top of the file explorer panel
    public string ExplorerHeaderText => IsFolderOpen
        ? Path.GetFileName(_currentFolderPath!.TrimEnd(Path.DirectorySeparatorChar)).ToUpperInvariant()
        : "EXPLORER";

    public string ThemeStatusText => $"Current theme: {CurrentThemeName}";

    public string EditorStatsText
    {
        get => _editorStatsText;
        private set
        {
            if (_editorStatsText == value) return;
            _editorStatsText = value;
            OnPropertyChanged();
        }
    }

    public string EditorStatsTextFinal => HasDocumentOpen ? _editorStatsText : string.Empty;

    public IBrush WindowBackgroundBrush { get; private set; } = Brush.Parse("#1E1E1E");
    public IBrush TopBarBrush { get; private set; } = Brush.Parse("#181818");
    public IBrush SidebarBrush { get; private set; } = Brush.Parse("#181818");
    public IBrush ButtonBrush { get; private set; } = Brush.Parse("#252526");
    public IBrush ButtonHoverBrush { get; private set; } = Brush.Parse("#313437");
    public IBrush EditorBackgroundBrush { get; private set; } = Brush.Parse("#1E1E1E");
    public IBrush CardBrush { get; private set; } = Brush.Parse("#252526");
    public IBrush PrimaryTextBrush { get; private set; } = Brush.Parse("#F4F4F4");
    public IBrush MutedTextBrush { get; private set; } = Brush.Parse("#A0A0A0");
    public IBrush SurfaceBorderBrush { get; private set; } = Brush.Parse("#2B2B2B");
    public IBrush AccentBrush { get; private set; } = Brush.Parse("#0E639C");

    event PropertyChangedEventHandler? INotifyPropertyChanged.PropertyChanged
    {
        add => ViewModelPropertyChanged += value;
        remove => ViewModelPropertyChanged -= value;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        ViewModelPropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private void RefreshState()
    {
        var lines = EditorContent.Length == 0 ? 1 : EditorContent.Count(static c => c == '\n') + 1;
        EditorStatsText = $"{lines} lines  |  {EditorContent.Length} characters";
        Title = HasDocumentOpen ? $"{GetDocumentDisplayName()} - Kodo" : "Kodo";
        OnPropertyChanged(nameof(HasDocumentOpen));
        OnPropertyChanged(nameof(HasFileOpen));
        OnPropertyChanged(nameof(IsFolderOpen));
        OnPropertyChanged(nameof(IsEmptyStateVisible));
        OnPropertyChanged(nameof(FileSummaryText));
        OnPropertyChanged(nameof(FilePathText));
        OnPropertyChanged(nameof(ExplorerHeaderText));
        OnPropertyChanged(nameof(ThemeStatusText));
        OnPropertyChanged(nameof(EditorStatsTextFinal));
    }

    // Light Mode and Dark Mode color definitions
    private void ApplyTheme(string themeName)
    {
        CurrentThemeName = themeName == "Light" ? "Light" : "Dark";
        Application.Current!.RequestedThemeVariant = CurrentThemeName == "Light"
            ? ThemeVariant.Light
            : ThemeVariant.Dark;

        if (CurrentThemeName == "Light") // Light mode colours
        {
            WindowBackgroundBrush = Brush.Parse("#F3F3F3");
            TopBarBrush = Brush.Parse("#FFFFFF");
            SidebarBrush = Brush.Parse("#EFF2F7");
            ButtonBrush = Brush.Parse("#E3E8F1");
            ButtonHoverBrush = Brush.Parse("#D5DDE9");
            EditorBackgroundBrush = Brush.Parse("#FFFFFF");
            CardBrush = Brush.Parse("#F7F9FC");
            PrimaryTextBrush = Brush.Parse("#202124");
            MutedTextBrush = Brush.Parse("#5F6B7A");
            SurfaceBorderBrush = Brush.Parse("#D7DCE5");
            AccentBrush = Brush.Parse("#0067C0");
        }
        else // Dark mode colours
        {
            WindowBackgroundBrush = Brush.Parse("#1E1E1E");
            TopBarBrush = Brush.Parse("#181818");
            SidebarBrush = Brush.Parse("#181818");
            ButtonBrush = Brush.Parse("#252526");
            ButtonHoverBrush = Brush.Parse("#313437");
            EditorBackgroundBrush = Brush.Parse("#1E1E1E");
            CardBrush = Brush.Parse("#252526");
            PrimaryTextBrush = Brush.Parse("#F4F4F4");
            MutedTextBrush = Brush.Parse("#A0A0A0");
            SurfaceBorderBrush = Brush.Parse("#2B2B2B");
            AccentBrush = Brush.Parse("#0E639C");
        }

        // Notify the UI that all the brushes have changed
        OnPropertyChanged(nameof(WindowBackgroundBrush));
        OnPropertyChanged(nameof(TopBarBrush));
        OnPropertyChanged(nameof(SidebarBrush));
        OnPropertyChanged(nameof(ButtonBrush));
        OnPropertyChanged(nameof(ButtonHoverBrush));
        OnPropertyChanged(nameof(EditorBackgroundBrush));
        OnPropertyChanged(nameof(CardBrush));
        OnPropertyChanged(nameof(PrimaryTextBrush));
        OnPropertyChanged(nameof(MutedTextBrush));
        OnPropertyChanged(nameof(SurfaceBorderBrush));
        OnPropertyChanged(nameof(AccentBrush));
        ApplyThemeToEditor();
        RefreshState();
        RefreshExtensionTheme();
    }

    private LoadedExtension? GetLanguageExtension(string filePath)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        return LoadedExtensions.FirstOrDefault(e => 
            e.Type == "language" && 
            e.Extensions.Any(ex => ex.Equals(ext, StringComparison.OrdinalIgnoreCase)));
    }

    private void ApplyThemeToEditor()
    {
        if (EditorTextBox is null) return;
        EditorTextBox.Background = EditorBackgroundBrush;
        EditorTextBox.Foreground = PrimaryTextBrush;
        EditorTextBox.LineNumbersForeground = MutedTextBrush;
        // The TextArea's selection and caret colors
        EditorTextBox.TextArea.SelectionBrush = AccentBrush.ToImmutable() is ISolidColorBrush b
            ? new SolidColorBrush(b.Color, 0.3)
            : new SolidColorBrush(Color.Parse("#0E639C"), 0.3);
        EditorTextBox.TextArea.SelectionForeground = PrimaryTextBrush;
    }

    private void ApplySyntaxHighlighting(LoadedExtension ext)
    {
        if (EditorTextBox is null) return;
        EditorTextBox.SyntaxHighlighting = new KodoHighlightingDefinition(ext);
    }

    // Opens a file picker dialog and loads the selected file into the editor
    private async Task OpenFileAsync()
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Open File",
            AllowMultiple = false
        });

        var file = files.Count > 0 ? files[0] : null;
        if (file is null) return;

        var path = file.TryGetLocalPath();
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;

        await OpenFileFromPathAsync(path);
    }

    private async Task OpenFileFromPathAsync(string path)
    {
        _currentFilePath = path;
        _hasUntitledDocument = false;
        var content = await File.ReadAllTextAsync(path);
        
        var langExt = GetLanguageExtension(path);
        CurrentLanguageExtension = langExt;
        
        if (langExt is not null)
            ApplySyntaxHighlighting(langExt);
        else
            EditorTextBox.SyntaxHighlighting = null;

        SetEditorContent(content);
        
        _isDirty = false;
        IsSettingsPageVisible = false;
        RefreshState();
        FocusEditor();
    }

    private void NewFile()
    {
        _currentFilePath = null;
        _hasUntitledDocument = true;
        SetEditorContent(string.Empty);
        CurrentLanguageExtension = null;
        EditorTextBox.SyntaxHighlighting = null;
        _isDirty = false;
        IsSettingsPageVisible = false;
        RefreshState();
        FocusEditor();
    }

    // Opens a folder picker, populates the file tree, and shows the explorer panel
    private async Task OpenFolderAsync()
    {
        var folders = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "Open Folder",
            AllowMultiple = false
        });

        var folder = folders.Count > 0 ? folders[0] : null;
        if (folder is null) return;

        var path = folder.TryGetLocalPath();
        if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path)) return;

        _currentFolderPath = path;
        PopulateFileTree(path);
        IsFileExplorerVisible = true;
        RefreshState();
    }

    // Clears the folder, tree, and hides the explorer panel
    private void CloseFolder()
    {
        _currentFolderPath = null;
        FileTreeItems.Clear();
        IsFileExplorerVisible = false;
        RefreshState();
    }

    // Builds the top-level contents of the folder into FileTreeItems.
    // Sub-directories start collapsed; their children are inserted lazily when expanded.
    private void PopulateFileTree(string folderPath)
    {
        FileTreeItems.Clear();
        AppendDirectoryContents(folderPath, depth: 0);
    }

    // Appends the immediate children of a directory into the flat list at the correct position.
    // insertAfterIndex == -1 means append to the end of the list.
    private void AppendDirectoryContents(string dirPath, int depth, int insertAfterIndex = -1)
    {
        var entries = GetSortedEntries(dirPath);
        var pos = insertAfterIndex + 1;

        foreach (var entry in entries)
        {
            var item = new FileTreeItem
            {
                Name = Path.GetFileName(entry),
                FullPath = entry,
                IsDirectory = Directory.Exists(entry),
                Depth = depth,
            };

            if (insertAfterIndex < 0)
                FileTreeItems.Add(item);
            else
            {
                FileTreeItems.Insert(pos, item);
                pos++;
            }
        }
    }

    // Returns subdirectories first (sorted), then files (sorted), skipping hidden entries
    private static string[] GetSortedEntries(string dirPath)
    {
        try
        {
            var dirs = Directory.GetDirectories(dirPath)
                .Where(d => !Path.GetFileName(d).StartsWith('.'))
                .OrderBy(d => Path.GetFileName(d), StringComparer.OrdinalIgnoreCase)
                .ToArray();

            var files = Directory.GetFiles(dirPath)
                .Where(f => !Path.GetFileName(f).StartsWith('.'))
                .OrderBy(f => Path.GetFileName(f), StringComparer.OrdinalIgnoreCase)
                .ToArray();

            return [.. dirs, .. files];
        }
        catch
        {
            return [];
        }
    }

    // Saves the current editor content, prompting for a path if needed
    private async Task SaveAsync()
    {
        if (_currentFilePath is null)
        {
            var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "Save File",
                SuggestedFileName = "untitled.txt"
            });

            var newPath = file?.TryGetLocalPath();
            if (string.IsNullOrWhiteSpace(newPath)) return;

            _currentFilePath = newPath;
            _hasUntitledDocument = false;
        }

        await File.WriteAllTextAsync(_currentFilePath, EditorContent);
        _isDirty = false;
        RefreshState();
    }

    // ── Event handlers ──────────────────────────────────────────────────────

    private void EditorButton_OnClick(object? sender, RoutedEventArgs e)
    {
        IsSettingsPageVisible = false;
        IsExtensionsPageVisible = false;
        FocusEditor();
    }

    private async void OpenFileButton_OnClick(object? sender, RoutedEventArgs e) =>
        await OpenFileAsync();

    private async void SaveButton_OnClick(object? sender, RoutedEventArgs e) =>
        await SaveAsync();

    private void NewFileButton_OnClick(object? sender, RoutedEventArgs e) =>
        NewFile();

    private async void OpenFolderButton_OnClick(object? sender, RoutedEventArgs e) =>
        await OpenFolderAsync();

    private void CloseFolderButton_OnClick(object? sender, RoutedEventArgs e) =>
        CloseFolder();

    // Hides the explorer panel without closing the folder
    private void CollapseExplorerButton_OnClick(object? sender, RoutedEventArgs e) =>
        IsFileExplorerVisible = false;

    private void SettingsButton_OnClick(object? sender, RoutedEventArgs e) =>
        IsSettingsPageVisible = true;

    private void ExtensionsButton_OnClick(object? sender, RoutedEventArgs e) =>
        IsExtensionsPageVisible = true;

    private void RefreshExtensionsButton_OnClick(object? sender, RoutedEventArgs e)
    {
        LoadExtensions();
    }

    private void BackToEditorButton_OnClick(object? sender, RoutedEventArgs e)
    {
        IsSettingsPageVisible = false;
        IsExtensionsPageVisible = false;
        FocusEditor();
    }

    // Handles clicks on file tree rows.
    // Directories toggle their children; files are opened in the editor.
    private async void FileTreeItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: FileTreeItem item })
        {
            if (item.IsDirectory)
                ToggleDirectoryExpansion(item);
            else
                await OpenFileFromTreeAsync(item.FullPath);
        }
    }

    // Expands or collapses a directory row in the flat list
    private void ToggleDirectoryExpansion(FileTreeItem dirItem)
    {
        var index = FileTreeItems.IndexOf(dirItem);
        if (index < 0) return;

        if (dirItem.IsExpanded)
        {
            // Collapse: remove all descendants (items with greater depth that follow consecutively)
            dirItem.IsExpanded = false;
            var toRemove = FileTreeItems
                .Skip(index + 1)
                .TakeWhile(i => i.Depth > dirItem.Depth)
                .ToList();
            foreach (var child in toRemove)
                FileTreeItems.Remove(child);
        }
        else
        {
            // Expand: insert immediate children right after this item
            dirItem.IsExpanded = true;
            AppendDirectoryContents(dirItem.FullPath, dirItem.Depth + 1, insertAfterIndex: index);
        }
    }

    // Opens a file from the tree into the editor
    private async Task OpenFileFromTreeAsync(string filePath)
    {
        if (!File.Exists(filePath)) return;
        await OpenFileFromPathAsync(filePath);
    }

    // Handles theme selection buttons in the settings page
    private void ThemeButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Control { Tag: string themeName })
            ApplyTheme(themeName);
    }

    // Marks the editor content as dirty (unsaved) whenever it changes
    private void EditorTextBox_OnTextChanged(object? sender, EventArgs e)
    {
        if (_suppressDirtyTracking) return;
        _editorContent = EditorTextBox.Text ?? string.Empty;
        _isDirty = true;
        RefreshState();
    }

    private void FocusEditor()
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (IsEditorPageVisible && HasDocumentOpen)
                EditorTextBox.TextArea.Focus();
        }, DispatcherPriority.Background);
    }

    private async void MainWindow_OnKeyDown(object? sender, KeyEventArgs e)
    {
        var hasControl = (e.KeyModifiers & KeyModifiers.Control) == KeyModifiers.Control;
        if (!hasControl) return;

        switch (e.Key)
        {
            case Key.N:
                NewFile();
                e.Handled = true;
                break;
            case Key.E:
                if ((e.KeyModifiers & KeyModifiers.Shift) == KeyModifiers.Shift)
                {
                    IsSettingsPageVisible = false;
                    FocusEditor();
                    e.Handled = true;
                }
                break;
            case Key.S:
                e.Handled = true;
                await SaveAsync();
                break;
            case Key.O:
                e.Handled = true;
                await OpenFileAsync();
                break;
        }
    }

    private string GetDocumentDisplayName() =>
        HasFileOpen ? Path.GetFileName(_currentFilePath!) : "untitled.txt";

    private void SetEditorContent(string content)
    {
        _suppressDirtyTracking = true;
        _editorContent = content;
        EditorTextBox.Text = content;
        _suppressDirtyTracking = false;
    }
}

// ── Syntax highlighting ─────────────────────────────────────────────────────
// Builds an AvaloniaEdit IHighlightingDefinition at runtime from the data
// declared in a LoadedExtension (keywords, types, comment markers, color tokens).
public sealed class KodoHighlightingDefinition : IHighlightingDefinition
{
    private readonly HighlightingRuleSet _mainRuleSet;

    public string Name { get; }
    public HighlightingRuleSet MainRuleSet => _mainRuleSet;
    public IEnumerable<HighlightingColor> NamedHighlightingColors => [];
    public IDictionary<string, string> Properties { get; } = new Dictionary<string, string>();

    public KodoHighlightingDefinition(LoadedExtension ext)
    {
        Name = ext.Name;
        _mainRuleSet = BuildRuleSet(ext);
    }

    private static HighlightingColor ColorFor(LoadedExtension ext, string tokenName, string fallback)
    {
        var hex = ext.ColorTokens.TryGetValue(tokenName, out var h) ? h : fallback;
        return new HighlightingColor
        {
            Foreground = new SimpleHighlightingBrush(Color.Parse(hex))
        };
    }

    private static HighlightingRuleSet BuildRuleSet(LoadedExtension ext)
    {
        var ruleSet = new HighlightingRuleSet();

        var commentColor = ColorFor(ext, "comment", "#6A9955");
        var stringColor  = ColorFor(ext, "string",  "#CE9178");
        var keywordColor = ColorFor(ext, "keyword",  "#569CD6");
        var typeColor    = ColorFor(ext, "type",     "#4EC9B0");
        var numberColor  = ColorFor(ext, "number",   "#B5CEA8");

        // ── Spans (multi-character constructs, evaluated before word rules) ──

        // Single-line comment  e.g.  // ...
        if (!string.IsNullOrEmpty(ext.CommentLine))
        {
            ruleSet.Spans.Add(new HighlightingSpan
            {
                StartExpression        = new Regex(Regex.Escape(ext.CommentLine), RegexOptions.Compiled),
                EndExpression          = new Regex("$", RegexOptions.Compiled | RegexOptions.Multiline),
                SpanColor              = commentColor,
                SpanColorIncludesStart = true,
                SpanColorIncludesEnd   = true,
                RuleSet                = new HighlightingRuleSet() // no nested rules inside comments
            });
        }

        // Block comment  e.g.  /* ... */
        if (!string.IsNullOrEmpty(ext.CommentBlockStart) && !string.IsNullOrEmpty(ext.CommentBlockEnd))
        {
            ruleSet.Spans.Add(new HighlightingSpan
            {
                StartExpression        = new Regex(Regex.Escape(ext.CommentBlockStart), RegexOptions.Compiled),
                EndExpression          = new Regex(Regex.Escape(ext.CommentBlockEnd),   RegexOptions.Compiled),
                SpanColor              = commentColor,
                SpanColorIncludesStart = true,
                SpanColorIncludesEnd   = true,
                RuleSet                = new HighlightingRuleSet()
            });
        }

        // Double-quoted string  "..."
        ruleSet.Spans.Add(new HighlightingSpan
        {
            StartExpression        = new Regex("\"", RegexOptions.Compiled),
            EndExpression          = new Regex("\"", RegexOptions.Compiled),
            SpanColor              = stringColor,
            SpanColorIncludesStart = true,
            SpanColorIncludesEnd   = true,
            RuleSet                = new HighlightingRuleSet()
        });

        // Single-quoted char/string  '...'
        ruleSet.Spans.Add(new HighlightingSpan
        {
            StartExpression        = new Regex("'", RegexOptions.Compiled),
            EndExpression          = new Regex("'", RegexOptions.Compiled),
            SpanColor              = stringColor,
            SpanColorIncludesStart = true,
            SpanColorIncludesEnd   = true,
            RuleSet                = new HighlightingRuleSet()
        });

        // ── Word rules ───────────────────────────────────────────────────────

        // Keywords  — only match whole words so "string" != "substring"
        if (ext.Keywords.Length > 0)
        {
            var pattern = @"\b(" + string.Join("|", ext.Keywords.Select(Regex.Escape)) + @")\b";
            ruleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(pattern, RegexOptions.Compiled),
                Color = keywordColor
            });
        }

        // Types
        if (ext.Types.Length > 0)
        {
            var pattern = @"\b(" + string.Join("|", ext.Types.Select(Regex.Escape)) + @")\b";
            ruleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(pattern, RegexOptions.Compiled),
                Color = typeColor
            });
        }

        // Numbers (integer and floating-point literals)
        ruleSet.Rules.Add(new HighlightingRule
        {
            Regex = new Regex(@"\b\d+(\.\d+)?\b", RegexOptions.Compiled),
            Color = numberColor
        });

        return ruleSet;
    }

    // IHighlightingDefinition plumbing
    public HighlightingColor GetNamedColor(string name) => new();
    public HighlightingRuleSet GetNamedRuleSet(string name) =>
        name == string.Empty ? _mainRuleSet : throw new KeyNotFoundException(name);
}
