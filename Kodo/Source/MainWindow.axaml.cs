// Licensed under the Kodo Public License v1.0
// April 19th, 2026 - KerbalMissile - Changed "One file at a time" note to "No file open"
// April 19th, 2026 - KerbalMissile - Added proper comments
// April 19th, 2026 - KerbalMissile - Changed "No File Open" at top bar to be empty when no file is open, and to show the file name when a file is open, also shows "unsaved" if there are unsaved changes
// April 19th, 2026 - KerbalMissile - Changed open file icon to fit in better
// April 19th, 2026 - KerbalMissile - Re-added SS-YYC's changes to improve Kodo's UI, did some changes the New File buttons but they still do not work
// April 20th, 2026 - SS-YYC - Added collapsible file explorer panel with folder tree and expand/collapse
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Platform;
using Avalonia.Platform.Storage;
using Avalonia.Styling;
using Avalonia.Threading;
using DiscordAssetsModel = DiscordRPC.Assets;
using DiscordRpcClient = DiscordRPC.DiscordRpcClient;
using DiscordRichPresenceModel = DiscordRPC.RichPresence;

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
    public string Icon => IsDirectory ? (_isExpanded ? "📂" : "📁") : GetFileIcon(Name);

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

public partial class MainWindow : Window, INotifyPropertyChanged
{
    private const string DefaultDiscordClientId = "1495509170756255744";
    private const string DefaultDiscordLargeImageKey = "kodo_logo";
    private const string DefaultDiscordLargeImageText = "Kodo";
    private const string SettingsFileName = "settings.json";
    private const string DiscordClientIdEnvironmentVariable = "KODO_DISCORD_CLIENT_ID";
    private string? _currentFilePath;
    private string? _currentFolderPath;
    private DiscordRpcClient? _discordRpcClient;
    private readonly DispatcherTimer _autoSaveTimer = new() { Interval = TimeSpan.FromSeconds(2) };
    private string _editorContent = string.Empty;
    private bool _isAutoSaveEnabled;
    private bool _isDirty;
    private bool _isSaving;
    private bool _isDiscordRichPresenceEnabled;
    private bool _hasUntitledDocument;
    private bool _isSettingsPageVisible;
    private bool _isFileExplorerVisible;
    private bool _suppressDirtyTracking;
    private string _currentThemeName = "Dark";
    private string _editorStatsText = "0 lines";
    private event PropertyChangedEventHandler? ViewModelPropertyChanged;

    // Flat list that backs the ItemsControl – directories insert/remove their children in-place
    public ObservableCollection<FileTreeItem> FileTreeItems { get; } = new();

    public MainWindow()
    {
        InitializeComponent();
        LoadWindowIcon();
        var settings = LoadSettings();
        _isAutoSaveEnabled = settings.AutoSaveEnabled;
        _isDiscordRichPresenceEnabled = settings.DiscordRichPresenceEnabled;
        _autoSaveTimer.Tick += AutoSaveTimer_OnTick;
        DataContext = this;
        ApplyTheme(settings.ThemeName);
        UpdateDiscordRichPresenceLifecycle();
        Closed += MainWindow_OnClosed;
        RefreshState();
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

    public bool IsEditorPageVisible => !IsSettingsPageVisible;

    public bool HasDocumentOpen => _currentFilePath is not null || _hasUntitledDocument;

    public bool HasFileOpen => _currentFilePath is not null;

    public bool IsFolderOpen => _currentFolderPath is not null;

    public bool IsEmptyStateVisible => !HasDocumentOpen;

    public bool IsDiscordRichPresenceEnabled
    {
        get => _isDiscordRichPresenceEnabled;
        set
        {
            if (_isDiscordRichPresenceEnabled == value) return;
            _isDiscordRichPresenceEnabled = value;
            OnPropertyChanged();
            SaveSettings();
            UpdateDiscordRichPresenceLifecycle();
            OnPropertyChanged(nameof(DiscordRichPresenceStatusText));
        }
    }

    public bool IsAutoSaveEnabled
    {
        get => _isAutoSaveEnabled;
        set
        {
            if (_isAutoSaveEnabled == value) return;
            _isAutoSaveEnabled = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(AutoSaveStatusText));
            if (!_isAutoSaveEnabled)
            {
                _autoSaveTimer.Stop();
            }
            else
            {
                RestartAutoSaveTimerIfNeeded();
            }

            SaveSettings();
        }
    }

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

    public string DiscordRichPresenceStatusText
    {
        get
        {
            if (!IsDiscordRichPresenceEnabled)
            {
                return "Discord Rich Presence is turned off.";
            }

            return "Discord Rich Presence is on when the Discord desktop app is running.";
        }
    }

    public string AutoSaveStatusText =>
        !IsAutoSaveEnabled
            ? "Autosave is turned off."
            : HasFileOpen
                ? "Changes are saved automatically a couple seconds after you stop typing."
                : "Autosave will start working after the file has been saved once.";

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
        if (HasDocumentOpen)
        {
            var lines = EditorContent.Length == 0 ? 1 : EditorContent.Count(static c => c == '\n') + 1;
            EditorStatsText = $"{lines} lines  |  {EditorContent.Length} characters";
        }
        else
        {
            EditorStatsText = string.Empty;
        }

        Title = HasDocumentOpen ? $"{GetDocumentDisplayName()} - Kodo" : "Kodo";
        OnPropertyChanged(nameof(HasDocumentOpen));
        OnPropertyChanged(nameof(HasFileOpen));
        OnPropertyChanged(nameof(IsFolderOpen));
        OnPropertyChanged(nameof(IsEmptyStateVisible));
        OnPropertyChanged(nameof(FileSummaryText));
        OnPropertyChanged(nameof(FilePathText));
        OnPropertyChanged(nameof(ExplorerHeaderText));
        OnPropertyChanged(nameof(ThemeStatusText));
        OnPropertyChanged(nameof(DiscordRichPresenceStatusText));
        OnPropertyChanged(nameof(AutoSaveStatusText));
        UpdateDiscordPresence();
    }

    private void LoadWindowIcon()
    {
        using var iconStream = AssetLoader.Open(new Uri("avares://Kodo/Assets/kodo-logo.png"));
        Icon = new WindowIcon(iconStream);
    }

    private string? GetDiscordApplicationId()
    {
        var overrideClientId = Environment.GetEnvironmentVariable(DiscordClientIdEnvironmentVariable);
        return string.IsNullOrWhiteSpace(overrideClientId) ? DefaultDiscordClientId : overrideClientId;
    }

    private void UpdateDiscordRichPresenceLifecycle()
    {
        try
        {
            var clientId = GetDiscordApplicationId();
            if (!IsDiscordRichPresenceEnabled || string.IsNullOrWhiteSpace(clientId))
            {
                DisposeDiscordPresence();
                return;
            }

            if (_discordRpcClient is null)
            {
                _discordRpcClient = new DiscordRpcClient(clientId);
                _discordRpcClient.Initialize();
            }

            UpdateDiscordPresence();
        }
        catch
        {
            DisposeDiscordPresence();
        }
    }

    private void UpdateDiscordPresence()
    {
        if (_discordRpcClient is null || !IsDiscordRichPresenceEnabled)
        {
            return;
        }

        try
        {
            _discordRpcClient.SetPresence(new DiscordRichPresenceModel
            {
                Details = GetDiscordPresenceDetails(),
                State = GetDiscordPresenceState(),
                Assets = new DiscordAssetsModel
                {
                    LargeImageKey = DefaultDiscordLargeImageKey,
                    LargeImageText = DefaultDiscordLargeImageText
                }
            });
        }
        catch
        {
            DisposeDiscordPresence();
        }
    }

    private string GetDiscordPresenceDetails()
    {
        if (HasDocumentOpen)
        {
            return $"Editing {GetDocumentDisplayName()}";
        }

        return IsFolderOpen ? "Browsing project files" : "Idle in Kodo";
    }

    private string GetDiscordPresenceState()
    {
        if (HasFileOpen)
        {
            return "Working in editor";
        }

        if (_hasUntitledDocument)
        {
            return "Editing an unsaved file";
        }

        if (IsFolderOpen)
        {
            return Path.GetFileName(_currentFolderPath!.TrimEnd(Path.DirectorySeparatorChar));
        }

        return "Waiting for a file";
    }

    private void DisposeDiscordPresence()
    {
        if (_discordRpcClient is null)
        {
            return;
        }

        try
        {
            _discordRpcClient.ClearPresence();
            _discordRpcClient.Dispose();
        }
        catch
        {
            // Ignore cleanup failures on shutdown/disable.
        }
        finally
        {
            _discordRpcClient = null;
        }
    }

    private string SettingsFilePath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Kodo", SettingsFileName);

    private AppSettings LoadSettings()
    {
        try
        {
            if (!File.Exists(SettingsFilePath))
            {
                return new AppSettings();
            }

            var json = File.ReadAllText(SettingsFilePath);
            var settings = JsonSerializer.Deserialize<AppSettings>(json);
            if (settings is null)
            {
                return new AppSettings();
            }

            settings.ThemeName = settings.ThemeName is "Light" or "Dark" ? settings.ThemeName : "Dark";
            return settings;
        }
        catch
        {
            return new AppSettings();
        }
    }

    private void SaveSettings()
    {
        try
        {
            var settingsDirectory = Path.GetDirectoryName(SettingsFilePath);
            if (!string.IsNullOrWhiteSpace(settingsDirectory))
            {
                Directory.CreateDirectory(settingsDirectory);
            }

            var json = JsonSerializer.Serialize(new AppSettings
            {
                ThemeName = CurrentThemeName,
                AutoSaveEnabled = IsAutoSaveEnabled,
                DiscordRichPresenceEnabled = IsDiscordRichPresenceEnabled
            });
            File.WriteAllText(SettingsFilePath, json);
        }
        catch
        {
            // Ignore settings persistence failures and keep the UI responsive.
        }
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
        SaveSettings();
        RefreshState();
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

        _currentFilePath = path;
        _hasUntitledDocument = false;
        _autoSaveTimer.Stop();
        SetEditorContent(await File.ReadAllTextAsync(path));
        _isDirty = false;
        IsSettingsPageVisible = false;
        RefreshState();
        FocusEditor();
    }

    private void NewFile()
    {
        _currentFilePath = null;
        _hasUntitledDocument = true;
        _autoSaveTimer.Stop();
        SetEditorContent(string.Empty);
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
    private async Task SaveAsync(bool allowPromptForPath = true)
    {
        _autoSaveTimer.Stop();

        if (_isSaving)
        {
            return;
        }

        if (_currentFilePath is null)
        {
            if (!allowPromptForPath)
            {
                return;
            }

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

        try
        {
            _isSaving = true;
            await File.WriteAllTextAsync(_currentFilePath, EditorContent);
            _isDirty = false;
            RefreshState();
        }
        finally
        {
            _isSaving = false;
        }
    }

    // ── Event handlers ──────────────────────────────────────────────────────

    private void EditorButton_OnClick(object? sender, RoutedEventArgs e)
    {
        IsSettingsPageVisible = false;
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

    private void BackToEditorButton_OnClick(object? sender, RoutedEventArgs e)
    {
        IsSettingsPageVisible = false;
        FocusEditor();
    }

    private void MainWindow_OnClosed(object? sender, EventArgs e)
    {
        _autoSaveTimer.Stop();
        DisposeDiscordPresence();
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

        _currentFilePath = filePath;
        _hasUntitledDocument = false;
        _autoSaveTimer.Stop();
        SetEditorContent(await File.ReadAllTextAsync(filePath));
        _isDirty = false;
        IsSettingsPageVisible = false;
        RefreshState();
        FocusEditor();
    }

    // Handles theme selection buttons in the settings page
    private void ThemeButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Control { Tag: string themeName })
            ApplyTheme(themeName);
    }

    // Marks the editor content as dirty (unsaved) whenever it changes
    private void EditorTextBox_OnTextChanged(object? sender, TextChangedEventArgs e)
    {
        if (_suppressDirtyTracking) return;
        _isDirty = true;
        RefreshState();
        RestartAutoSaveTimerIfNeeded();
    }

    private void RestartAutoSaveTimerIfNeeded()
    {
        if (IsAutoSaveEnabled && HasFileOpen && _isDirty)
        {
            _autoSaveTimer.Stop();
            _autoSaveTimer.Start();
        }
    }

    private async void AutoSaveTimer_OnTick(object? sender, EventArgs e)
    {
        _autoSaveTimer.Stop();

        if (!IsAutoSaveEnabled || !HasFileOpen || !_isDirty)
        {
            return;
        }

        await SaveAsync(allowPromptForPath: false);
    }

    private void FocusEditor()
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (IsEditorPageVisible && HasDocumentOpen)
                EditorTextBox.Focus();
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
        EditorContent = content;
        _suppressDirtyTracking = false;
    }

    private sealed class AppSettings
    {
        public string ThemeName { get; set; } = "Dark";
        public bool AutoSaveEnabled { get; set; }
        public bool DiscordRichPresenceEnabled { get; set; }
    }
}
