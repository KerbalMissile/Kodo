// Licensed under the Kodo Public License v1.0
// April 19th, 2026 - KerbalMissile - Changed "One file at a time" note to "No file open"
// April 19th, 2026 - KerbalMissile - Added proper comments
// April 19th, 2026 - KerbalMissile - Changed "No File Open" at top bar to be empty when no file is open, and to show the file name when a file is open, also shows "unsaved" if there are unsaved changes
// April 19th, 2026 - KerbalMissile - Changed open file icon to fit in better
// April 19th, 2026 - KerbalMissile - Re-added SS-YYC's changes to improve Kodo's UI, did some changes the New File buttons but they still do not work
// April 20th, 2026 - SS-YYC - Added collapsible file explorer panel with folder tree and expand/collapse
// April 20th, 2026 - KerbalMissile - Added extension support, re-added full screen by default, updated open / close folder icon
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Net.Http;
using System.Text.Json;
using System.Text.RegularExpressions;
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
using AvaloniaEdit.Highlighting;
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
			    ".cs" or ".csproj" or ".axaml.cs" or ".csx" => "📄",
			    ".axaml" or ".xaml" => "🪟",
			    ".xml" or ".html" or ".htm" => "📋",
			    ".json" or ".yaml" or ".yml" or ".toml" => "📝",
			    ".md" or ".txt" or ".rst" => "📃",
			    ".png" or ".jpg" or ".jpeg" or ".gif" or ".svg" or ".ico" => "🖼",
			    ".py" => "🐍",
			    ".js" or ".ts" or ".jsx" or ".tsx" => "📜",
			    ".vue" or ".svelte" => "📜",
			    ".css" or ".scss" or ".less" => "🎨",
			    ".sh" or ".bat" or ".ps1" => "⚡",
			    ".zip" or ".tar" or ".gz" or ".rar" => "📦",
			    ".cpp" or ".c" or ".h" or ".hpp" => "📄",
			    ".rs" => "📄",
			    ".go" => "📄",
			    ".rb" => "📄",
			    ".java" or ".kt" or ".kts" => "📄",
			    ".swift" => "📄",
			    ".fs" or ".fsi" or ".fsx" => "📄",
			    ".sql" => "🗃",
			    ".lua" => "📄",
			    ".r" => "📄",
			    ".lock" => "🔒",
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
    public IBrush AccentBrush { get; set; } = Brush.Parse("#8C00FF");
    public IBrush CardBrush { get; set; } = Brush.Parse("#252526");
    public IBrush PrimaryTextBrush { get; set; } = Brush.Parse("#F4F4F4");
    public IBrush SurfaceBorderBrush { get; set; } = Brush.Parse("#2B2B2B");
    public IBrush MutedTextBrush { get; set; } = Brush.Parse("#A0A0A0");
    public string SourcePath { get; set; } = string.Empty;
    public bool IsDirectorySource { get; set; }
    public ExtensionThemeDefinition? ThemeDefinition { get; set; }

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

public class ExtensionThemeDefinition
{
    public string ThemeId { get; init; } = string.Empty;
    public string DisplayName { get; init; } = string.Empty;
    public string BaseTheme { get; init; } = "Dark";
    public string WindowBackground { get; init; } = "#000000";
    public string TopBar { get; init; } = "#0E0E0E";
    public string Sidebar { get; init; } = "#0E0E0E";
    public string Button { get; init; } = "#242424";
    public string ButtonHover { get; init; } = "#343434";
    public string EditorBackground { get; init; } = "#000000";
    public string Card { get; init; } = "#121212";
    public string PrimaryText { get; init; } = "#FFFFFF";
    public string MutedText { get; init; } = "#BDBDBD";
    public string SurfaceBorder { get; init; } = "#4A4A4A";
    public string Accent { get; init; } = "#8C00FF";
    public string PreviewBackground { get; init; } = "#000000";
    public string PreviewBorder { get; init; } = "#4A4A4A";
}

public class MarketplaceExtension : INotifyPropertyChanged
{
    private bool _isInstalling;
    private string _installButtonText = "Install";

    public string Id { get; init; } = string.Empty;
    public string Version { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public string Type { get; init; } = string.Empty;
    public string Author { get; init; } = string.Empty;
    public string Description { get; init; } = string.Empty;
    public string DownloadUrl { get; init; } = string.Empty;
    public string FileName { get; init; } = string.Empty;

    public bool IsInstalling
    {
        get => _isInstalling;
        set
        {
            if (_isInstalling == value) return;
            _isInstalling = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsInstallEnabled));
        }
    }

    public bool IsInstalled { get; private set; }

    public bool IsInstallEnabled => !IsInstalling && !IsInstalled && !string.IsNullOrWhiteSpace(DownloadUrl);

    public string InstallButtonText
    {
        get => _installButtonText;
        set
        {
            if (_installButtonText == value) return;
            _installButtonText = value;
            OnPropertyChanged();
        }
    }

    public void SetInstalledState(bool isInstalled)
    {
        if (IsInstalled == isInstalled)
        {
            if (!isInstalled && !IsInstalling && InstallButtonText != "Install")
                InstallButtonText = "Install";
            return;
        }

        IsInstalled = isInstalled;
        InstallButtonText = isInstalled ? "Installed" : "Install";
        OnPropertyChanged(nameof(IsInstalled));
        OnPropertyChanged(nameof(IsInstallEnabled));
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}

public partial class MainWindow : Window, INotifyPropertyChanged
{
    private const int MaxRecentFiles = 5;
    private const string DefaultDiscordClientId = "1495509170756255744";
    private const string DefaultDiscordLargeImageKey = "kodo_logo";
    private const string DefaultDiscordLargeImageText = "Kodo";
    private const string SettingsFileName = "settings.json";
    private const string DiscordClientIdEnvironmentVariable = "KODO_DISCORD_CLIENT_ID";
    private const string AutoSaveSavedMessage = "Saved.";
    private const string AutoSaveSavingMessage = "Saving...";
    private const string DefaultMarketplaceIndexUrl = "https://raw.githubusercontent.com/KerbalMissile/Kodo/main/Indexs/ExtensionsIndex.json";

    private string? _currentFilePath;
    private string? _currentFolderPath;
    private DiscordRpcClient? _discordRpcClient;
    private readonly DispatcherTimer _autoSaveTimer = new() { Interval = TimeSpan.FromSeconds(2) };
    private readonly DispatcherTimer _autoSaveStatusTimer = new() { Interval = TimeSpan.FromSeconds(3) };
    private string? _autoSaveStatusMessage;
    private bool _isAutoSaveEnabled;
    private bool _isDirty;
    private bool _isSaving;
    private bool _isDiscordRichPresenceEnabled;
    private bool _hasUntitledDocument;
    private bool _isRefreshingExtensions;
    private bool _isSettingsPageVisible;
    private bool _isExtensionsPageVisible;
    private bool _isFileExplorerVisible;
    private bool _isMarketplaceTabSelected;
    private bool _suppressDirtyTracking;
    private string _currentThemeName = "Dark";
    private string _requestedThemeName = "Dark";
    private string _editorStatsText = "0 lines";
    private string _extensionsStatusText = "Drop .kox extension files into the Extensions folder to install them.";
    private LoadedExtension? _currentLanguageExtension;
    private event PropertyChangedEventHandler? ViewModelPropertyChanged;
    private static readonly HttpClient MarketplaceHttpClient = new();

    private string ExtensionsFolderPath => Path.Combine(Directory.GetCurrentDirectory(), "Extensions");
    private string IndexFolderPath => Path.Combine(Directory.GetCurrentDirectory(), "Indexs");
    private string ProjectExtensionsFolderPath =>
        Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "Extensions"));

    // Flat list that backs the ItemsControl – directories insert/remove their children in-place
    public ObservableCollection<FileTreeItem> FileTreeItems { get; } = new();
    public ObservableCollection<RecentFileItem> RecentFiles { get; } = new();
    public ObservableCollection<LoadedExtension> LoadedExtensions { get; } = new();
    public ObservableCollection<MarketplaceExtension> MarketplaceExtensions { get; } = new();
    public ObservableCollection<string> ExtensionLoadErrors { get; } = new();

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
        LoadWindowIcon();
        // TextEditor uses EventHandler (not RoutedEventHandler), so hook up in code-behind
        EditorTextBox.TextChanged += EditorTextBox_OnTextChanged;
        var settings = LoadSettings();
        _requestedThemeName = string.IsNullOrWhiteSpace(settings.ThemeName) ? "Dark" : settings.ThemeName;
        _isAutoSaveEnabled = settings.AutoSaveEnabled;
        _isDiscordRichPresenceEnabled = settings.DiscordRichPresenceEnabled;
        LoadRecentFiles(settings.RecentFiles);
        _autoSaveTimer.Tick += AutoSaveTimer_OnTick;
        _autoSaveStatusTimer.Tick += AutoSaveStatusTimer_OnTick;
        DataContext = this;
        ApplyTheme(_requestedThemeName);
        EnsureExtensionsFolder();
        _ = RefreshExtensionsDataAsync();
        UpdateDiscordRichPresenceLifecycle();
        Closed += MainWindow_OnClosed;
        RefreshState();
    }

    // ── Extension loading ────────────────────────────────────────────────────

    private void EnsureExtensionsFolder()
    {
        if (!Directory.Exists(ExtensionsFolderPath))
            Directory.CreateDirectory(ExtensionsFolderPath);
    }

    private async Task RefreshExtensionsDataAsync()
    {
        if (_isRefreshingExtensions)
            return;

        IsRefreshingExtensions = true;
        ExtensionsStatusText = "Refreshing installed extensions and marketplace...";

        try
        {
            LoadExtensions();
            await LoadMarketplaceExtensionsAsync();
            if (!string.Equals(CurrentThemeName, _requestedThemeName, StringComparison.OrdinalIgnoreCase) &&
                (string.Equals(_requestedThemeName, "Light", StringComparison.OrdinalIgnoreCase) ||
                 string.Equals(_requestedThemeName, "Dark", StringComparison.OrdinalIgnoreCase) ||
                 ThemeExtensions.Any(t => string.Equals(t.ThemeDefinition!.ThemeId, _requestedThemeName, StringComparison.OrdinalIgnoreCase))))
            {
                ApplyTheme(_requestedThemeName);
            }
            ExtensionsStatusText = $"Refreshed {LoadedExtensions.Count} installed and {MarketplaceExtensions.Count} marketplace extension(s).";
        }
        catch (Exception ex)
        {
            ExtensionsStatusText = $"Refresh failed: {ex.Message}";
        }
        finally
        {
            IsRefreshingExtensions = false;
        }
    }

    private void LoadExtensions()
    {
        LoadedExtensions.Clear();
        ExtensionLoadErrors.Clear();

        var searchPaths = new List<string>();
        var binPath = ExtensionsFolderPath;
        searchPaths.Add(binPath);

        // Also search the project source tree when running from the build output directory
        var projectRoot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..");
        var srcPath = Path.GetFullPath(Path.Combine(projectRoot, "Extensions"));
        if (!string.Equals(srcPath, binPath, StringComparison.OrdinalIgnoreCase))
            searchPaths.Add(srcPath);

        var anyFolderFound = false;
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
                    if (File.Exists(Path.Combine(dir, "manifest.json")))
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
            ExtensionLoadErrors.Add($"Extensions folder not found. Expected: {binPath}");

        OnPropertyChanged(nameof(ExtensionLoadErrors));
        OnPropertyChanged(nameof(IsNoExtensionsVisible));
        OnPropertyChanged(nameof(ThemeExtensions));
        OnPropertyChanged(nameof(HasThemeExtensions));
        RefreshExtensionTheme();
        SyncMarketplaceInstallStates();
    }

    private async Task LoadMarketplaceExtensionsAsync()
    {
        MarketplaceExtensions.Clear();
        var loadedAny = false;

        foreach (var indexPath in GetExtensionsIndexSearchPaths())
        {
            if (!File.Exists(indexPath))
                continue;

            try
            {
                using var doc = JsonDocument.Parse(File.ReadAllText(indexPath));
                if (!doc.RootElement.TryGetProperty("extensions", out var extensionsElement) ||
                    extensionsElement.ValueKind != JsonValueKind.Array)
                    continue;

                foreach (var item in extensionsElement.EnumerateArray())
                {
                    var entry = ParseMarketplaceExtension(item);
                    if (string.IsNullOrWhiteSpace(entry.Id) || MarketplaceExtensions.Any(e => e.Id == entry.Id))
                        continue;

                    MarketplaceExtensions.Add(entry);
                }

                loadedAny = MarketplaceExtensions.Count > 0;
                break;
            }
            catch (Exception ex)
            {
                ExtensionLoadErrors.Add($"Failed to load marketplace index '{Path.GetFileName(indexPath)}': {ex.Message}");
            }
        }

        if (!loadedAny)
        {
            try
            {
                var remoteJson = await MarketplaceHttpClient.GetStringAsync(DefaultMarketplaceIndexUrl);
                using var doc = JsonDocument.Parse(remoteJson);
                if (doc.RootElement.TryGetProperty("extensions", out var extensionsElement) &&
                    extensionsElement.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in extensionsElement.EnumerateArray())
                    {
                        var entry = ParseMarketplaceExtension(item);
                        if (string.IsNullOrWhiteSpace(entry.Id) || MarketplaceExtensions.Any(e => e.Id == entry.Id))
                            continue;

                        MarketplaceExtensions.Add(entry);
                    }

                    if (MarketplaceExtensions.Count > 0)
                        ExtensionsStatusText = $"Loaded marketplace from {DefaultMarketplaceIndexUrl}";
                }
            }
            catch (Exception ex)
            {
                ExtensionLoadErrors.Add($"Failed to load remote marketplace index: {ex.Message}");
            }
        }

        SyncMarketplaceInstallStates();
        OnPropertyChanged(nameof(ExtensionLoadErrors));
        OnPropertyChanged(nameof(IsMarketplaceEmptyVisible));
    }

    private IEnumerable<string> GetExtensionsIndexSearchPaths()
    {
        yield return Path.Combine(IndexFolderPath, "ExtensionsIndex.json");

        var projectRoot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..");
        yield return Path.GetFullPath(Path.Combine(projectRoot, "Indexs", "ExtensionsIndex.json"));
    }

    private static MarketplaceExtension ParseMarketplaceExtension(JsonElement item) => new()
    {
        Id = item.TryGetProperty("id", out var id) ? id.GetString() ?? string.Empty : string.Empty,
        Version = item.TryGetProperty("version", out var version) ? version.GetString() ?? string.Empty : string.Empty,
        Name = item.TryGetProperty("name", out var name) ? name.GetString() ?? string.Empty : string.Empty,
        Type = item.TryGetProperty("type", out var type) ? type.GetString() ?? string.Empty : string.Empty,
        Author = item.TryGetProperty("author", out var author) ? author.GetString() ?? string.Empty : string.Empty,
        Description = item.TryGetProperty("description", out var description) ? description.GetString() ?? string.Empty : string.Empty,
        DownloadUrl = item.TryGetProperty("downloadUrl", out var downloadUrl) ? downloadUrl.GetString() ?? string.Empty : string.Empty,
        FileName = item.TryGetProperty("fileName", out var fileName) ? fileName.GetString() ?? string.Empty : string.Empty
    };

    private void SyncMarketplaceInstallStates()
    {
        foreach (var entry in MarketplaceExtensions)
            entry.SetInstalledState(LoadedExtensions.Any(ext => ext.Id.Equals(entry.Id, StringComparison.OrdinalIgnoreCase)));
    }

    private async Task InstallMarketplaceExtensionAsync(MarketplaceExtension marketplaceExtension)
    {
        if (marketplaceExtension.IsInstalling || marketplaceExtension.IsInstalled)
            return;

        marketplaceExtension.IsInstalling = true;
        marketplaceExtension.InstallButtonText = "Installing...";
        ExtensionsStatusText = $"Installing {marketplaceExtension.Name}...";

        try
        {
            EnsureExtensionsFolder();

            var fileName = string.IsNullOrWhiteSpace(marketplaceExtension.FileName)
                ? Path.GetFileName(new Uri(marketplaceExtension.DownloadUrl).AbsolutePath)
                : marketplaceExtension.FileName;

            var outputPath = Path.Combine(ExtensionsFolderPath, fileName);
            var bytes = await MarketplaceHttpClient.GetByteArrayAsync(marketplaceExtension.DownloadUrl);
            await File.WriteAllBytesAsync(outputPath, bytes);

            await RefreshExtensionsDataAsync();
            ExtensionsStatusText = $"{marketplaceExtension.Name} installed.";
        }
        catch (Exception ex)
        {
            marketplaceExtension.SetInstalledState(false);
            ExtensionsStatusText = $"Failed to install {marketplaceExtension.Name}: {ex.Message}";
        }
        finally
        {
            marketplaceExtension.IsInstalling = false;
            if (!marketplaceExtension.IsInstalled)
                marketplaceExtension.InstallButtonText = "Install";
        }
    }

    private async Task UninstallExtensionAsync(LoadedExtension extension)
    {
        if (string.IsNullOrWhiteSpace(extension.SourcePath))
        {
            ExtensionsStatusText = $"Cannot uninstall {extension.Name}: missing source path.";
            return;
        }

        try
        {
            var resolvedPath = Path.GetFullPath(extension.SourcePath);
            if (!IsPathInsideDirectory(resolvedPath, ExtensionsFolderPath) &&
                !IsPathInsideDirectory(resolvedPath, ProjectExtensionsFolderPath))
            {
                ExtensionsStatusText = $"Cannot uninstall {extension.Name}: source is outside the Extensions folders.";
                return;
            }

            if (extension.IsDirectorySource)
            {
                if (Directory.Exists(resolvedPath))
                    Directory.Delete(resolvedPath, recursive: true);
            }
            else
            {
                if (File.Exists(resolvedPath))
                    File.Delete(resolvedPath);
            }

            await RefreshExtensionsDataAsync();
            ExtensionsStatusText = $"{extension.Name} uninstalled.";
        }
        catch (Exception ex)
        {
            ExtensionsStatusText = $"Failed to uninstall {extension.Name}: {ex.Message}";
        }
    }

    private static bool IsPathInsideDirectory(string path, string directory)
    {
        var normalizedPath = Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalizedDirectory = Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        return normalizedPath.StartsWith(normalizedDirectory + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
            || string.Equals(normalizedPath, normalizedDirectory, StringComparison.OrdinalIgnoreCase);
    }

    private LoadedExtension? LoadExtensionFromFolder(string folderPath)
    {
        var manifestPath = Path.Combine(folderPath, "manifest.json");
        if (!File.Exists(manifestPath)) return null;

        using var manifestDoc = JsonDocument.Parse(File.ReadAllText(manifestPath));
        var ext = ParseManifest(manifestDoc.RootElement);

        var languagePath = Path.Combine(folderPath, "language.json");
        if (File.Exists(languagePath))
        {
            using var langDoc = JsonDocument.Parse(File.ReadAllText(languagePath));
            ParseLanguage(langDoc.RootElement, ext);
        }

        var themePath = Path.Combine(folderPath, "theme.json");
        if (File.Exists(themePath))
        {
            using var themeDoc = JsonDocument.Parse(File.ReadAllText(themePath));
            ext.ThemeDefinition = ParseTheme(themeDoc.RootElement, ext);
        }

        ext.SourcePath = folderPath;
        ext.IsDirectorySource = true;
        return ext;
    }

    private LoadedExtension? LoadExtensionFromKox(string koxPath)
    {
        using var archive = ZipFile.OpenRead(koxPath);
        var manifestEntry = archive.GetEntry("manifest.json");
        if (manifestEntry is null) return null;

        using var manifestStream = manifestEntry.Open();
        using var manifestDoc = JsonDocument.Parse(manifestStream);
        var ext = ParseManifest(manifestDoc.RootElement);

        var languageEntry = archive.GetEntry("language.json");
        if (languageEntry is not null)
        {
            using var langStream = languageEntry.Open();
            using var langDoc = JsonDocument.Parse(langStream);
            ParseLanguage(langDoc.RootElement, ext);
        }

        var themeEntry = archive.GetEntry("theme.json");
        if (themeEntry is not null)
        {
            using var themeStream = themeEntry.Open();
            using var themeDoc = JsonDocument.Parse(themeStream);
            ext.ThemeDefinition = ParseTheme(themeDoc.RootElement, ext);
        }

        ext.SourcePath = koxPath;
        ext.IsDirectorySource = false;
        return ext;
    }

    private static LoadedExtension ParseManifest(JsonElement manifest) => new()
    {
        Id          = manifest.TryGetProperty("id",          out var id)   ? id.GetString()   ?? "" : "",
        Version     = manifest.TryGetProperty("version",     out var ver)  ? ver.GetString()  ?? "" : "",
        Name        = manifest.TryGetProperty("name",        out var name) ? name.GetString() ?? "" : "",
        Type        = manifest.TryGetProperty("type",        out var type) ? type.GetString() ?? "" : "",
        Author      = manifest.TryGetProperty("author",      out var auth) ? auth.GetString() ?? "" : "",
        Description = manifest.TryGetProperty("description", out var desc) ? desc.GetString() ?? "" : "",
        Extensions  = manifest.TryGetProperty("extensions",  out var exts)
            ? exts.EnumerateArray().Select(e => e.GetString() ?? "").ToArray()
            : []
    };

    private static void ParseLanguage(JsonElement lang, LoadedExtension ext)
    {
        ext.Keywords          = lang.TryGetProperty("keywords",          out var kw)  ? kw.EnumerateArray().Select(e => e.GetString() ?? "").ToArray() : [];
        ext.Types             = lang.TryGetProperty("types",             out var tp)  ? tp.EnumerateArray().Select(e => e.GetString() ?? "").ToArray() : [];
        ext.CommentLine       = lang.TryGetProperty("commentLine",       out var cl)  ? cl.GetString()  ?? "//" : "//";
        ext.CommentBlockStart = lang.TryGetProperty("commentBlockStart", out var cbs) ? cbs.GetString() ?? "/*" : "/*";
        ext.CommentBlockEnd   = lang.TryGetProperty("commentBlockEnd",   out var cbe) ? cbe.GetString() ?? "*/" : "*/";

        if (lang.TryGetProperty("colorTokens", out var ct))
            foreach (var prop in ct.EnumerateObject())
                ext.ColorTokens[prop.Name] = prop.Value.GetString() ?? "#FFFFFF";
    }

    private static ExtensionThemeDefinition ParseTheme(JsonElement theme, LoadedExtension ext) => new()
    {
        ThemeId = theme.TryGetProperty("themeId", out var themeId) ? themeId.GetString() ?? ext.Id : ext.Id,
        DisplayName = theme.TryGetProperty("displayName", out var displayName) ? displayName.GetString() ?? ext.Name : ext.Name,
        BaseTheme = theme.TryGetProperty("baseTheme", out var baseTheme) ? baseTheme.GetString() ?? "Dark" : "Dark",
        WindowBackground = GetThemeColor(theme, "windowBackground", "#000000"),
        TopBar = GetThemeColor(theme, "topBar", "#0E0E0E"),
        Sidebar = GetThemeColor(theme, "sidebar", "#0E0E0E"),
        Button = GetThemeColor(theme, "button", "#242424"),
        ButtonHover = GetThemeColor(theme, "buttonHover", "#343434"),
        EditorBackground = GetThemeColor(theme, "editorBackground", "#000000"),
        Card = GetThemeColor(theme, "card", "#121212"),
        PrimaryText = GetThemeColor(theme, "primaryText", "#FFFFFF"),
        MutedText = GetThemeColor(theme, "mutedText", "#BDBDBD"),
        SurfaceBorder = GetThemeColor(theme, "surfaceBorder", "#4A4A4A"),
        Accent = GetThemeColor(theme, "accent", "#8C00FF"),
        PreviewBackground = GetThemeColor(theme, "previewBackground", GetThemeColor(theme, "editorBackground", "#000000")),
        PreviewBorder = GetThemeColor(theme, "previewBorder", GetThemeColor(theme, "surfaceBorder", "#4A4A4A"))
    };

    private static string GetThemeColor(JsonElement theme, string propertyName, string fallback) =>
        theme.TryGetProperty(propertyName, out var value) ? value.GetString() ?? fallback : fallback;

    private void RefreshExtensionTheme()
    {
        foreach (var ext in LoadedExtensions)
        {
            ext.AccentBrush        = AccentBrush;
            ext.CardBrush          = CardBrush;
            ext.PrimaryTextBrush   = PrimaryTextBrush;
            ext.SurfaceBorderBrush = SurfaceBorderBrush;
            ext.MutedTextBrush     = MutedTextBrush;
            ext.NotifyAllBrushesChanged();
        }
    }

    private LoadedExtension? GetLanguageExtension(string filePath)
    {
        var fileExt = Path.GetExtension(filePath).ToLowerInvariant();
        return LoadedExtensions.FirstOrDefault(e =>
            e.Type == "language" &&
            e.Extensions.Any(ex => ex.Equals(fileExt, StringComparison.OrdinalIgnoreCase)));
    }

    // ── Theme / editor appearance ────────────────────────────────────────────

    private void ApplyThemeToEditor()
    {
        if (EditorTextBox is null) return;
        EditorTextBox.Background = EditorBackgroundBrush;
        EditorTextBox.Foreground = PrimaryTextBrush;
        EditorTextBox.LineNumbersForeground = MutedTextBrush;
        EditorTextBox.TextArea.SelectionBrush = AccentBrush.ToImmutable() is ISolidColorBrush b
            ? new SolidColorBrush(b.Color, 0.3)
            : new SolidColorBrush(Color.Parse("#8C00FF"), 0.3);
        EditorTextBox.TextArea.SelectionForeground = PrimaryTextBrush;
    }

    private void ApplySyntaxHighlighting(LoadedExtension ext)
    {
        if (EditorTextBox is null) return;
        EditorTextBox.SyntaxHighlighting = new KodoHighlightingDefinition(ext);
    }

    // ── Properties ───────────────────────────────────────────────────────────

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

    public bool IsEditorPageVisible => !IsSettingsPageVisible && !IsExtensionsPageVisible;

    public bool HasDocumentOpen => _currentFilePath is not null || _hasUntitledDocument;

    public bool HasFileOpen => _currentFilePath is not null;

    public bool IsFolderOpen => _currentFolderPath is not null;

    public bool IsEmptyStateVisible => !HasDocumentOpen;

    public bool HasRecentFiles => RecentFiles.Count > 0;

    public bool IsNoExtensionsVisible => LoadedExtensions.Count == 0;

    public IEnumerable<LoadedExtension> ThemeExtensions =>
        LoadedExtensions.Where(e => e.Type.Equals("theme", StringComparison.OrdinalIgnoreCase) && e.ThemeDefinition is not null);

    public bool HasThemeExtensions => ThemeExtensions.Any();

    public bool IsRefreshingExtensions
    {
        get => _isRefreshingExtensions;
        private set
        {
            if (_isRefreshingExtensions == value) return;
            _isRefreshingExtensions = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(RefreshExtensionsButtonText));
        }
    }

    public bool IsInstalledTabSelected
    {
        get => !_isMarketplaceTabSelected;
        private set
        {
            var shouldSelectMarketplace = !value;
            if (_isMarketplaceTabSelected == shouldSelectMarketplace) return;
            _isMarketplaceTabSelected = shouldSelectMarketplace;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsMarketplaceTabSelected));
        }
    }

    public bool IsMarketplaceTabSelected
    {
        get => _isMarketplaceTabSelected;
        private set
        {
            if (_isMarketplaceTabSelected == value) return;
            _isMarketplaceTabSelected = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsInstalledTabSelected));
        }
    }

    public string ExtensionsStatusText
    {
        get => _extensionsStatusText;
        private set
        {
            if (_extensionsStatusText == value) return;
            _extensionsStatusText = value;
            OnPropertyChanged();
        }
    }

    public string RefreshExtensionsButtonText => IsRefreshingExtensions ? "Refreshing..." : "Refresh";

    public bool IsMarketplaceEmptyVisible => MarketplaceExtensions.Count == 0;

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
                _autoSaveTimer.Stop();
            else
                RestartAutoSaveTimerIfNeeded();
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

    // Displays the file name and unsaved/autosave status in the top bar
    public string FileSummaryText => HasDocumentOpen
        ? $"{GetDocumentDisplayName()}{GetDocumentStatusSuffix()}"
        : "Open A File";

    public string FilePathText => HasFileOpen
        ? _currentFilePath!
        : HasDocumentOpen ? "Unsaved file"
        : IsFolderOpen ? $"📂 {_currentFolderPath}"
        : "No file open";

    public string ExplorerHeaderText => IsFolderOpen
        ? Path.GetFileName(_currentFolderPath!.TrimEnd(Path.DirectorySeparatorChar)).ToUpperInvariant()
        : "EXPLORER";

    public string ThemeStatusText => $"Current theme: {CurrentThemeName}";

    public string DiscordRichPresenceStatusText => !IsDiscordRichPresenceEnabled
        ? "Discord Rich Presence is turned off."
        : "Discord Rich Presence is on when the Discord desktop app is running.";

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
    public IBrush TopBarBrush           { get; private set; } = Brush.Parse("#181818");
    public IBrush SidebarBrush          { get; private set; } = Brush.Parse("#181818");
    public IBrush ButtonBrush           { get; private set; } = Brush.Parse("#252526");
    public IBrush ButtonHoverBrush      { get; private set; } = Brush.Parse("#313437");
    public IBrush EditorBackgroundBrush { get; private set; } = Brush.Parse("#1E1E1E");
    public IBrush CardBrush             { get; private set; } = Brush.Parse("#252526");
    public IBrush PrimaryTextBrush      { get; private set; } = Brush.Parse("#F4F4F4");
    public IBrush MutedTextBrush        { get; private set; } = Brush.Parse("#A0A0A0");
    public IBrush SurfaceBorderBrush    { get; private set; } = Brush.Parse("#2B2B2B");
    public IBrush AccentBrush           { get; private set; } = Brush.Parse("#8C00FF");

    event PropertyChangedEventHandler? INotifyPropertyChanged.PropertyChanged
    {
        add    => ViewModelPropertyChanged += value;
        remove => ViewModelPropertyChanged -= value;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null) =>
        ViewModelPropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

    // ── State management ─────────────────────────────────────────────────────

    private void RefreshState()
    {
        var content = EditorTextBox?.Document?.Text ?? string.Empty;
        if (HasDocumentOpen)
        {
            var lines = content.Length == 0 ? 1 : content.Count(static c => c == '\n') + 1;
            EditorStatsText = $"{lines} lines  |  {content.Length} characters";
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
        OnPropertyChanged(nameof(HasRecentFiles));
        OnPropertyChanged(nameof(FileSummaryText));
        OnPropertyChanged(nameof(FilePathText));
        OnPropertyChanged(nameof(ExplorerHeaderText));
        OnPropertyChanged(nameof(ThemeStatusText));
        OnPropertyChanged(nameof(ThemeExtensions));
        OnPropertyChanged(nameof(HasThemeExtensions));
        OnPropertyChanged(nameof(DiscordRichPresenceStatusText));
        OnPropertyChanged(nameof(AutoSaveStatusText));
        UpdateDiscordPresence();
    }

    private string GetDocumentStatusSuffix()
    {
        var text = GetDocumentStatusText();
        return string.IsNullOrWhiteSpace(text) ? string.Empty : $" • {text}";
    }

    private string? GetDocumentStatusText()
    {
        if (!HasDocumentOpen) return null;

        if (IsAutoSaveEnabled && HasFileOpen)
        {
            if (!string.IsNullOrWhiteSpace(_autoSaveStatusMessage))
                return _autoSaveStatusMessage;

            if (_isSaving || _isDirty || _autoSaveTimer.IsEnabled)
                return AutoSaveSavingMessage;
        }

        return _isDirty ? "unsaved" : null;
    }

    // ── Window icon ──────────────────────────────────────────────────────────

    private void LoadWindowIcon()
    {
        using var iconStream = AssetLoader.Open(new Uri("avares://Kodo/Assets/kodo-logo.png"));
        Icon = new WindowIcon(iconStream);
    }

    // ── Discord Rich Presence ────────────────────────────────────────────────

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
        catch { DisposeDiscordPresence(); }
    }

    private void UpdateDiscordPresence()
    {
        if (_discordRpcClient is null || !IsDiscordRichPresenceEnabled) return;

        try
        {
            _discordRpcClient.SetPresence(new DiscordRichPresenceModel
            {
                Details = GetDiscordPresenceDetails(),
                State   = GetDiscordPresenceState(),
                Assets  = new DiscordAssetsModel
                {
                    LargeImageKey  = DefaultDiscordLargeImageKey,
                    LargeImageText = DefaultDiscordLargeImageText
                }
            });
        }
        catch { DisposeDiscordPresence(); }
    }

    private string GetDiscordPresenceDetails()
    {
        if (HasDocumentOpen) return $"Editing {GetDocumentDisplayName()}";
        return IsFolderOpen ? "Browsing project files" : "Idle in Kodo";
    }

    private string GetDiscordPresenceState()
    {
        if (HasFileOpen)          return GetDiscordWorkspaceLabel();
        if (_hasUntitledDocument) return GetDiscordWorkspaceLabel("Editing an unsaved file");
        if (IsFolderOpen)         return GetDiscordWorkspaceLabel();
        return "Waiting for a file";
    }

    private string GetDiscordWorkspaceLabel(string fallback = "Working in editor")
    {
        if (!IsFolderOpen) return fallback;
        var folderName = Path.GetFileName(_currentFolderPath!.TrimEnd(Path.DirectorySeparatorChar));
        return string.IsNullOrWhiteSpace(folderName) ? fallback : $"Workspace: {folderName}";
    }

    private void DisposeDiscordPresence()
    {
        if (_discordRpcClient is null) return;
        try
        {
            _discordRpcClient.ClearPresence();
            _discordRpcClient.Dispose();
        }
        catch { /* Ignore cleanup failures. */ }
        finally { _discordRpcClient = null; }
    }

    // ── Settings persistence ─────────────────────────────────────────────────

    private string SettingsFilePath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Kodo", SettingsFileName);

    private AppSettings LoadSettings()
    {
        try
        {
            if (!File.Exists(SettingsFilePath)) return new AppSettings();

            var settings = JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(SettingsFilePath));
            if (settings is null) return new AppSettings();

            settings.ThemeName = string.IsNullOrWhiteSpace(settings.ThemeName) ? "Dark" : settings.ThemeName;
            settings.RecentFiles = settings.RecentFiles?
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? [];
            return settings;
        }
        catch { return new AppSettings(); }
    }

    private void SaveSettings()
    {
        try
        {
            var dir = Path.GetDirectoryName(SettingsFilePath);
            if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);

            File.WriteAllText(SettingsFilePath, JsonSerializer.Serialize(new AppSettings
            {
                ThemeName                  = CurrentThemeName,
                AutoSaveEnabled            = IsAutoSaveEnabled,
                DiscordRichPresenceEnabled = IsDiscordRichPresenceEnabled,
                RecentFiles                = RecentFiles.Select(f => f.Path).ToList()
            }));
        }
        catch { /* Ignore settings persistence failures. */ }
    }

    // ── Theme application ────────────────────────────────────────────────────

    private void ApplyTheme(string themeName)
    {
        _requestedThemeName = themeName;
        var extensionTheme = ThemeExtensions
            .Select(e => e.ThemeDefinition!)
            .FirstOrDefault(t => string.Equals(t.ThemeId, themeName, StringComparison.OrdinalIgnoreCase));

        if (extensionTheme is not null)
        {
            CurrentThemeName = extensionTheme.ThemeId;
            Application.Current!.RequestedThemeVariant = string.Equals(extensionTheme.BaseTheme, "Light", StringComparison.OrdinalIgnoreCase)
                ? ThemeVariant.Light
                : ThemeVariant.Dark;

            WindowBackgroundBrush = Brush.Parse(extensionTheme.WindowBackground);
            TopBarBrush           = Brush.Parse(extensionTheme.TopBar);
            SidebarBrush          = Brush.Parse(extensionTheme.Sidebar);
            ButtonBrush           = Brush.Parse(extensionTheme.Button);
            ButtonHoverBrush      = Brush.Parse(extensionTheme.ButtonHover);
            EditorBackgroundBrush = Brush.Parse(extensionTheme.EditorBackground);
            CardBrush             = Brush.Parse(extensionTheme.Card);
            PrimaryTextBrush      = Brush.Parse(extensionTheme.PrimaryText);
            MutedTextBrush        = Brush.Parse(extensionTheme.MutedText);
            SurfaceBorderBrush    = Brush.Parse(extensionTheme.SurfaceBorder);
            AccentBrush           = Brush.Parse(extensionTheme.Accent);
        }
        else
        {
            CurrentThemeName = themeName == "Light" ? "Light" : "Dark";
            Application.Current!.RequestedThemeVariant = CurrentThemeName == "Light"
                ? ThemeVariant.Light
                : ThemeVariant.Dark;

            if (CurrentThemeName == "Light")
            {
                WindowBackgroundBrush = Brush.Parse("#F3F3F3");
                TopBarBrush           = Brush.Parse("#FFFFFF");
                SidebarBrush          = Brush.Parse("#EFF2F7");
                ButtonBrush           = Brush.Parse("#E3E8F1");
                ButtonHoverBrush      = Brush.Parse("#D5DDE9");
                EditorBackgroundBrush = Brush.Parse("#FFFFFF");
                CardBrush             = Brush.Parse("#F7F9FC");
                PrimaryTextBrush      = Brush.Parse("#202124");
                MutedTextBrush        = Brush.Parse("#5F6B7A");
                SurfaceBorderBrush    = Brush.Parse("#D7DCE5");
                AccentBrush           = Brush.Parse("#8C00FF");
            }
            else
            {
                WindowBackgroundBrush = Brush.Parse("#1E1E1E");
                TopBarBrush           = Brush.Parse("#181818");
                SidebarBrush          = Brush.Parse("#181818");
                ButtonBrush           = Brush.Parse("#252526");
                ButtonHoverBrush      = Brush.Parse("#313437");
                EditorBackgroundBrush = Brush.Parse("#1E1E1E");
                CardBrush             = Brush.Parse("#252526");
                PrimaryTextBrush      = Brush.Parse("#F4F4F4");
                MutedTextBrush        = Brush.Parse("#A0A0A0");
                SurfaceBorderBrush    = Brush.Parse("#2B2B2B");
                AccentBrush           = Brush.Parse("#8C00FF");
            }
        }

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
        SaveSettings();
        RefreshState();
        RefreshExtensionTheme();
    }

    // ── File operations ──────────────────────────────────────────────────────

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

    // Central method used by Open File, Open From Tree, and Open Recent
    private async Task OpenFileFromPathAsync(string path)
    {
        _currentFilePath = path;
        _hasUntitledDocument = false;
        _autoSaveTimer.Stop();
        ClearAutoSaveStatus();

        var langExt = GetLanguageExtension(path);
        CurrentLanguageExtension = langExt;
        if (langExt is not null)
            ApplySyntaxHighlighting(langExt);
        else
            EditorTextBox.SyntaxHighlighting = null;

        SetEditorContent(await File.ReadAllTextAsync(path));
        _isDirty = false;
        IsSettingsPageVisible = false;
        IsExtensionsPageVisible = false;
        RefreshState();
        FocusEditor();
    }

    private void NewFile()
    {
        _currentFilePath = null;
        _hasUntitledDocument = true;
        _autoSaveTimer.Stop();
        ClearAutoSaveStatus();
        CurrentLanguageExtension = null;
        EditorTextBox.SyntaxHighlighting = null;
        SetEditorContent(string.Empty);
        _isDirty = false;
        IsSettingsPageVisible = false;
        IsExtensionsPageVisible = false;
        RefreshState();
        FocusEditor();
    }

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

    private void CloseFolder()
    {
        _currentFolderPath = null;
        FileTreeItems.Clear();
        IsFileExplorerVisible = false;
        RefreshState();
    }

    private async Task SaveAsync(bool allowPromptForPath = true)
    {
        _autoSaveTimer.Stop();
        if (_isSaving) return;

        if (_currentFilePath is null)
        {
            if (!allowPromptForPath) return;

            var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "Save File",
                SuggestedFileName = "untitled.txt"
            });

            var newPath = file?.TryGetLocalPath();
            if (string.IsNullOrWhiteSpace(newPath)) return;

            _currentFilePath = newPath;
            _hasUntitledDocument = false;
            ClearAutoSaveStatus();
        }

        try
        {
            _isSaving = true;
            if (IsAutoSaveEnabled && HasFileOpen)
            {
                _autoSaveStatusMessage = AutoSaveSavingMessage;
                _autoSaveStatusTimer.Stop();
                OnPropertyChanged(nameof(FileSummaryText));
            }

            // Read content directly from the TextEditor document
            await File.WriteAllTextAsync(_currentFilePath, EditorTextBox.Document.Text);
            _isDirty = false;
            AddRecentFile(_currentFilePath);

            if (IsAutoSaveEnabled && HasFileOpen)
            {
                _autoSaveStatusMessage = AutoSaveSavedMessage;
                _autoSaveStatusTimer.Stop();
                _autoSaveStatusTimer.Start();
            }

            RefreshState();
        }
        finally
        {
            _isSaving = false;
            OnPropertyChanged(nameof(FileSummaryText));
        }
    }

    // ── File tree ────────────────────────────────────────────────────────────

    private void PopulateFileTree(string folderPath)
    {
        FileTreeItems.Clear();
        AppendDirectoryContents(folderPath, depth: 0);
    }

    private void AppendDirectoryContents(string dirPath, int depth, int insertAfterIndex = -1)
    {
        var entries = GetSortedEntries(dirPath);
        var pos = insertAfterIndex + 1;

        foreach (var entry in entries)
        {
            var item = new FileTreeItem
            {
                Name        = Path.GetFileName(entry),
                FullPath    = entry,
                IsDirectory = Directory.Exists(entry),
                Depth       = depth,
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
        catch { return []; }
    }

    private void ToggleDirectoryExpansion(FileTreeItem dirItem)
    {
        var index = FileTreeItems.IndexOf(dirItem);
        if (index < 0) return;

        if (dirItem.IsExpanded)
        {
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
            dirItem.IsExpanded = true;
            AppendDirectoryContents(dirItem.FullPath, dirItem.Depth + 1, insertAfterIndex: index);
        }
    }

    // ── Recent files ─────────────────────────────────────────────────────────

    private void LoadRecentFiles(IEnumerable<string>? recentFiles)
    {
        RecentFiles.Clear();
        foreach (var path in (recentFiles ?? []).Take(MaxRecentFiles))
            if (File.Exists(path))
                RecentFiles.Add(new RecentFileItem(path));
    }

    private void AddRecentFile(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return;

        var existing = RecentFiles.FirstOrDefault(f => string.Equals(f.Path, path, StringComparison.OrdinalIgnoreCase));
        if (existing is not null) RecentFiles.Remove(existing);

        RecentFiles.Insert(0, new RecentFileItem(path));
        while (RecentFiles.Count > MaxRecentFiles)
            RecentFiles.RemoveAt(RecentFiles.Count - 1);

        SaveSettings();
        OnPropertyChanged(nameof(HasRecentFiles));
    }

    private void RemoveRecentFile(string path)
    {
        var existing = RecentFiles.FirstOrDefault(f => string.Equals(f.Path, path, StringComparison.OrdinalIgnoreCase));
        if (existing is null) return;
        RecentFiles.Remove(existing);
        SaveSettings();
        OnPropertyChanged(nameof(HasRecentFiles));
    }

    // ── Autosave helpers ─────────────────────────────────────────────────────

    private void RestartAutoSaveTimerIfNeeded()
    {
        if (IsAutoSaveEnabled && HasFileOpen && _isDirty)
        {
            _autoSaveTimer.Stop();
            _autoSaveTimer.Start();
        }
    }

    private void ClearAutoSaveStatus()
    {
        if (string.IsNullOrWhiteSpace(_autoSaveStatusMessage)) return;
        _autoSaveStatusTimer.Stop();
        _autoSaveStatusMessage = null;
        OnPropertyChanged(nameof(FileSummaryText));
    }

    // ── Editor helpers ───────────────────────────────────────────────────────

    private string GetDocumentDisplayName() =>
        HasFileOpen ? Path.GetFileName(_currentFilePath!) : "untitled.txt";

    // Writes content into the TextEditor document without triggering dirty tracking
    private void SetEditorContent(string content)
    {
        _suppressDirtyTracking = true;
        EditorTextBox.Document.Text = content;
        _suppressDirtyTracking = false;
    }

    private void FocusEditor()
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (IsEditorPageVisible && HasDocumentOpen)
                EditorTextBox.TextArea.Focus();
        }, DispatcherPriority.Background);
    }

    // ── Event handlers ───────────────────────────────────────────────────────

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

    private void CollapseExplorerButton_OnClick(object? sender, RoutedEventArgs e) =>
        IsFileExplorerVisible = false;

    private void SettingsButton_OnClick(object? sender, RoutedEventArgs e)
    {
        IsExtensionsPageVisible = false;
        IsSettingsPageVisible = true;
    }

    private void ExtensionsButton_OnClick(object? sender, RoutedEventArgs e)
    {
        IsSettingsPageVisible = false;
        IsExtensionsPageVisible = true;
        _ = RefreshExtensionsDataAsync();
    }

    private async void RefreshExtensionsButton_OnClick(object? sender, RoutedEventArgs e) =>
        await RefreshExtensionsDataAsync();

    private void InstalledTabButton_OnClick(object? sender, RoutedEventArgs e) =>
        IsMarketplaceTabSelected = false;

    private void MarketplaceTabButton_OnClick(object? sender, RoutedEventArgs e) =>
        IsMarketplaceTabSelected = true;

    private void BackToEditorButton_OnClick(object? sender, RoutedEventArgs e)
    {
        IsSettingsPageVisible = false;
        IsExtensionsPageVisible = false;
        FocusEditor();
    }

    private void ThemeButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Control { Tag: string themeName })
            ApplyTheme(themeName);
    }

    private async void FileTreeItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: FileTreeItem item })
        {
            if (item.IsDirectory)
                ToggleDirectoryExpansion(item);
            else
                await OpenFileFromPathAsync(item.FullPath);
        }
    }

    private async void RecentFileButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: string filePath })
        {
            if (!File.Exists(filePath))
            {
                RemoveRecentFile(filePath);
                return;
            }
            await OpenFileFromPathAsync(filePath);
        }
    }

    private async void InstallMarketplaceExtensionButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: MarketplaceExtension marketplaceExtension })
            await InstallMarketplaceExtensionAsync(marketplaceExtension);
    }

    private async void UninstallExtensionButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: LoadedExtension extension })
            await UninstallExtensionAsync(extension);
    }

    // TextEditor fires EventHandler (not RoutedEventHandler) — signature must match exactly
    private void EditorTextBox_OnTextChanged(object? sender, EventArgs e)
    {
        if (_suppressDirtyTracking) return;
        ClearAutoSaveStatus();
        _isDirty = true;
        RefreshState();
        RestartAutoSaveTimerIfNeeded();
    }

    private async void AutoSaveTimer_OnTick(object? sender, EventArgs e)
    {
        _autoSaveTimer.Stop();
        if (!IsAutoSaveEnabled || !HasFileOpen || !_isDirty) return;
        await SaveAsync(allowPromptForPath: false);
    }

    private void AutoSaveStatusTimer_OnTick(object? sender, EventArgs e)
    {
        _autoSaveStatusTimer.Stop();
        ClearAutoSaveStatus();
    }

    private void MainWindow_OnClosed(object? sender, EventArgs e)
    {
        _autoSaveTimer.Stop();
        _autoSaveStatusTimer.Stop();
        DisposeDiscordPresence();
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
                    IsExtensionsPageVisible = false;
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

    // ── Nested types ─────────────────────────────────────────────────────────

    private sealed class AppSettings
    {
        public string ThemeName { get; set; } = "Dark";
        public bool AutoSaveEnabled { get; set; }
        public bool DiscordRichPresenceEnabled { get; set; }
        public List<string> RecentFiles { get; set; } = [];
    }
}

public sealed class RecentFileItem
{
    public RecentFileItem(string path) => Path = path;
    public string Path { get; }
    public string DisplayName  => System.IO.Path.GetFileName(Path);
    public string DirectoryPath => System.IO.Path.GetDirectoryName(Path) ?? string.Empty;
}

// ── Syntax highlighting ──────────────────────────────────────────────────────
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
        return new HighlightingColor { Foreground = new SimpleHighlightingBrush(Color.Parse(hex)) };
    }

    private static HighlightingRuleSet BuildRuleSet(LoadedExtension ext)
    {
        var ruleSet = new HighlightingRuleSet();

        var commentColor = ColorFor(ext, "comment", "#6A9955");
        var stringColor  = ColorFor(ext, "string",  "#CE9178");
        var keywordColor = ColorFor(ext, "keyword",  "#569CD6");
        var typeColor    = ColorFor(ext, "type",     "#4EC9B0");
        var numberColor  = ColorFor(ext, "number",   "#B5CEA8");

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
                RuleSet                = new HighlightingRuleSet()
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

        // Keywords — only match whole words so "string" != "substring"
        if (ext.Keywords.Length > 0)
        {
            ruleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(@"\b(" + string.Join("|", ext.Keywords.Select(Regex.Escape)) + @")\b", RegexOptions.Compiled),
                Color = keywordColor
            });
        }

        // Types
        if (ext.Types.Length > 0)
        {
            ruleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(@"\b(" + string.Join("|", ext.Types.Select(Regex.Escape)) + @")\b", RegexOptions.Compiled),
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

    public HighlightingColor GetNamedColor(string name) => new();
    public HighlightingRuleSet GetNamedRuleSet(string name) =>
        name == string.Empty ? _mainRuleSet : throw new KeyNotFoundException(name);
}
