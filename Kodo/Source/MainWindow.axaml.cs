// Licensed under the Kodo Public License v1.1
// May 8th, 2026 - Fixed settings.json resetting by using atomic write (temp file + replace)
// May 8th, 2026 - Added ShowWarningDialogAsync for non-fatal recoverable errors
// May 5th, 2026 - SS-YYC - Added Unsaved dot indicator on editor tabs, close other/all tabs context menu, last-modified time on recent files, language indicator in status bar, font size setting, collapse all tree button, extension search filter, active theme indicator, find in file panel (Ctrl+F), file tree right-click context menu (copy path, delete)
// April 29th, 2026 - SS-YYC - Fixed syntax highlighting: keywords/numbers no longer colour inside comments or strings by isolating rules into a code-only ruleset that comment and string spans cannot inherit from
// April 29th, 2026 - SS-YYC - Fixed language.json and language2.json now both apply and merge correctly for both .kox and folder extensions
// April 29th, 2026 - SS-YYC - Fixed marketplace icons always showing abbreviation boxes instead of index icons
// April 29th, 2026 - SS-YYC - Fixed installed extensions not showing index icons in the Installed tab
// April 29th, 2026 - SS-YYC - Fixed install/uninstall not refreshing due to cooldown, safe URI parsing in install path, HttpClient timeout, removed redundant binPath variable
// April 29th, 2026 - SS-YYC - Marketplace now pulls exclusively from the web, removed local index file lookup
// April 24th, 2026 - SS-YYC - Fixed multi-theme support: theme.json arrays now create one LoadedExtension per theme entry
// April 19th, 2026 - KerbalMissile - Changed "One file at a time" note to "No file open"
// April 19th, 2026 - KerbalMissile - Added proper comments
// April 19th, 2026 - KerbalMissile - Changed "No File Open" at top bar to be empty when no file is open, and to show the file name when a file is open, also shows "Unsaved" if there are Unsaved changes
// April 19th, 2026 - KerbalMissile - Changed open file icon to fit in better
// April 19th, 2026 - KerbalMissile - Re-added SS-YYC's changes to improve Kodo's UI, did some changes the New File buttons but they still do not work
// April 20th, 2026 - SS-YYC - Added collapsible file explorer panel with folder tree and expand/collapse
// April 20th, 2026 - KerbalMissile - Added extension support, re-added full screen by default, updated open / close folder icon
using System;
using System.Reflection;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using Avalonia.Platform.Storage;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using AvaloniaEdit.Editing;
using AvaloniaEdit.Highlighting;
using AvaloniaEdit.Rendering;
using DiscordAssetsModel = DiscordRPC.Assets;
using DiscordRpcClient = DiscordRPC.DiscordRpcClient;
using DiscordRichPresenceModel = DiscordRPC.RichPresence;
using Kodo.Models;

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
    public string Icon => IsDirectory ? (_isExpanded ? "\U0001F4C2" : "\U0001F4C1") : GetFileIcon(Name);

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

    // Returns a simple file-type icon based on extension
    private static string GetFileIcon(string fileName)
    {
        var ext = Path.GetExtension(fileName).ToLowerInvariant();
        return ext switch
			{
			    ".cs" or ".csproj" or ".axaml.cs" or ".csx" => "C#",
			    ".axaml" or ".xaml" => "XM",
			    ".xml" or ".html" or ".htm" => "HTML",
			    ".json" or ".yaml" or ".yml" or ".toml" => "JSON",
			    ".md" or ".txt" or ".rst" => "MD",
			    ".png" => "PNG",
			    ".jpg" or ".jpeg" => "JPG",
			    ".gif" => "GIF",
			    ".svg" => "SVG",
			    ".ico" => "ICO",
			    ".webp" => "WBP",
			    ".bmp" => "BMP",
			    ".py" => "Py",
			    ".js" or ".jsx" => "JS",
			    ".ts" or ".tsx" => "TS",
			    ".vue" or ".svelte" => "UI",
			    ".css" or ".scss" or ".less" => "CSS",
			    ".sh" or ".bat" or ".ps1" => ">_",
			    ".zip" or ".tar" or ".gz" or ".rar" => "ZIP",
			    ".cpp" or ".cc" or ".cxx" => "C++",
			    ".c" => "C",
			    ".h" or ".hpp" or ".hxx" => "C++",
			    ".rs" => "Rs",
			    ".go" => "Go",
			    ".rb" => "Rb",
			    ".java" => "JAVA",
			    ".kt" or ".kts" => "Kt",
			    ".swift" => "Sw",
			    ".fs" or ".fsi" or ".fsx" => "F#",
			    ".sql" => "Db",
			    ".lua" => "Lu",
			    ".r" => "R",
			    ".lock" => "Lk",
				".csv" or ".tsv" => "CSV",
				".nova" => "NOVA",
				".kox" => "KOX",
			    _ => "..",
			};
    }
}

public record class LoadedExtension : INotifyPropertyChanged
{
    private bool _isUpdateAvailable;

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
    public string[] StringDelimiters { get; set; } = ["\"", "'"];
    public string[] MultiLineStringDelimiters { get; set; } = [];
    // When true, the open-ended single-quote span is replaced with a precise
    // char-literal regex. Set via "disableSingleQuoteStrings": true in language.json.
    // Use for languages like C# where ' appears in non-string contexts.
    public bool DisableSingleQuoteStrings { get; set; }
    public Dictionary<string, string> ColorTokens { get; set; } = new();
    public List<LanguageSyntaxProfile> SyntaxProfiles { get; } = [];
    public IBrush AccentBrush { get; set; } = Brush.Parse("#8C00FF");
    public IBrush CardBrush { get; set; } = Brush.Parse("#252526");
    public IBrush PrimaryTextBrush { get; set; } = Brush.Parse("#F4F4F4");
    public IBrush SurfaceBorderBrush { get; set; } = Brush.Parse("#2B2B2B");
    public IBrush MutedTextBrush { get; set; } = Brush.Parse("#A0A0A0");
    public string SourcePath { get; set; } = string.Empty;
    public bool IsDirectorySource { get; set; }
    public ExtensionThemeDefinition? ThemeDefinition { get; set; }
    public string ThemeCardThemeId => ThemeDefinition?.ThemeId ?? string.Empty;
    public string ThemeCardDisplayName => ThemeDefinition?.DisplayName ?? Name;
    public string ThemeCardPreviewBackground => ThemeDefinition?.PreviewBackground ?? "#000000";
    public string ThemeCardPreviewBorder => ThemeDefinition?.PreviewBorder ?? "#4A4A4A";
    public string ThemeCardAccent => ThemeDefinition?.Accent ?? "#8C00FF";
    // True for the 2nd, 3rd, etc. entries split out of a multi-theme array —
    // they appear in ThemeExtensions but are hidden from the Installed list.
    public bool IsThemeSubEntry { get; init; }
    // Optional icon loaded from icon.png inside the .kox / folder
    public Bitmap? IconImage { get; set; }
    // Fallback: first two letters of the name, shown when no icon is present
    public string NameAbbreviation => Name.Length >= 2 ? Name[..2] : Name;
    public bool HasIcon => IconImage is not null;
    public bool IsUpdateAvailable
    {
        get => _isUpdateAvailable;
        set
        {
            if (_isUpdateAvailable == value) return;
            _isUpdateAvailable = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsUpdateAvailable)));
        }
    }

    private bool _isActiveTheme;
    public bool IsActiveTheme
    {
        get => _isActiveTheme;
        set
        {
            if (_isActiveTheme == value) return;
            _isActiveTheme = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsActiveTheme)));
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public void NotifyAllBrushesChanged()
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(AccentBrush)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CardBrush)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(PrimaryTextBrush)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SurfaceBorderBrush)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(MutedTextBrush)));
    }

    public void NotifyIconChanged()
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IconImage)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasIcon)));
    }
}

public sealed class LanguageSyntaxProfile
{
    public string[] Extensions { get; init; } = [];
    public string[] Keywords { get; init; } = [];
    public string[] Types { get; init; } = [];
    public string? CommentLine { get; init; }
    public string? CommentBlockStart { get; init; }
    public string? CommentBlockEnd { get; init; }
    public string[]? StringDelimiters { get; init; }
    public string[]? MultiLineStringDelimiters { get; init; }
    public bool? DisableSingleQuoteStrings { get; init; }
    public Dictionary<string, string> ColorTokens { get; init; } = new();
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

public class ReleaseInfo
{
    public string Name { get; init; } = string.Empty;
    public string Tag { get; init; } = string.Empty;
    public string Notes { get; init; } = string.Empty;
    public string Url { get; init; } = string.Empty;
}

public class ReleaseLinkItem
{
    public string Label { get; init; } = string.Empty;
    public string Url { get; init; } = string.Empty;
}

public enum ExplorerClipboardMode
{
    Copy,
    Cut
}

public class MarketplaceExtension : INotifyPropertyChanged
{
    private bool _isInstalling;
    private string _installButtonText = "Install";
    private bool _isUpdateAvailable;
    private string _installedVersion = string.Empty;

    public string Id { get; init; } = string.Empty;
    public string Version { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public string Type { get; init; } = string.Empty;
    public string Author { get; init; } = string.Empty;
    public string Description { get; init; } = string.Empty;
    public string DownloadUrl { get; init; } = string.Empty;
    public string FileName { get; init; } = string.Empty;
    public string IconUrl { get; init; } = string.Empty;

    private Bitmap? _iconImage;
    public Bitmap? IconImage
    {
        get => _iconImage;
        set
        {
            if (_iconImage == value) return;
            _iconImage = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(HasIcon));
        }
    }

    public bool HasIcon => IconImage is not null;
    public string NameAbbreviation => Name.Length >= 2 ? Name[..2] : Name;

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

    public bool IsUpdateAvailable
    {
        get => _isUpdateAvailable;
        private set
        {
            if (_isUpdateAvailable == value) return;
            _isUpdateAvailable = value;
            OnPropertyChanged();
        }
    }

    public string InstalledVersion
    {
        get => _installedVersion;
        private set
        {
            if (_installedVersion == value) return;
            _installedVersion = value;
            OnPropertyChanged();
        }
    }

    public bool IsInstallEnabled => !IsInstalling && (!IsInstalled || IsUpdateAvailable) && !string.IsNullOrWhiteSpace(DownloadUrl);

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

    public void SetInstalledState(LoadedExtension? installedExtension, bool isUpdateAvailable)
    {
        var isInstalled = installedExtension is not null;
        InstalledVersion = installedExtension?.Version ?? string.Empty;
        IsUpdateAvailable = isUpdateAvailable;

        if (IsInstalled != isInstalled)
        {
            IsInstalled = isInstalled;
            OnPropertyChanged(nameof(IsInstalled));
        }

        if (!IsInstalling)
        {
            InstallButtonText = isInstalled
                ? (isUpdateAvailable ? "Update" : "Installed")
                : "Install";
        }

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
    private const string AutoSaveFailedMessagePrefix = "Save failed:";
    // Read from <InformationalVersion> in Kodo.csproj (e.g. "v1.3.0-BETA").
    // To update the app version, change only that tag in the csproj.
    private static readonly string CurrentAppVersion =
        Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion ?? "v0.0.0";
    private const string DefaultMarketplaceIndexUrl = "https://raw.githubusercontent.com/KerbalMissile/Kodo/main/Indexs/ExtensionsIndex.json";
    private const string LatestReleaseApiUrl = "https://api.github.com/repos/KerbalMissile/Kodo/releases/latest";
    private const string ReleasesApiUrl = "https://api.github.com/repos/KerbalMissile/Kodo/releases";
    private const string ReleasesPageUrl = "https://github.com/KerbalMissile/Kodo/releases";

    private string? _currentFilePath;
    private string? _currentFolderPath;
    private DiscordRpcClient? _discordRpcClient;
    private readonly DispatcherTimer _autoSaveTimer = new() { Interval = TimeSpan.FromSeconds(2) };
    private readonly DispatcherTimer _autoSaveStatusTimer = new() { Interval = TimeSpan.FromSeconds(3) };
    private readonly DispatcherTimer _discordReconnectTimer = new() { Interval = TimeSpan.FromSeconds(10) };
    private readonly DispatcherTimer _editorStateRefreshTimer = new() { Interval = TimeSpan.FromMilliseconds(75) };
    private readonly RainbowBracketColorizer _rainbowBracketColorizer = new();
    private EditorTab? _activeEditorTab;
    private int _nextUntitledTabNumber = 1;
    private string? _autoSaveStatusMessage;
    private bool _isAutoSaveEnabled;
    private bool _isDirty;
    private bool _isSaving;
    private bool _isDiscordRichPresenceEnabled;
    private bool _hasUntitledDocument;
    private bool _isRefreshingExtensions;
    private bool _isRefreshingLatestRelease;
    private bool _isSettingsPageVisible;
    private bool _isExtensionsPageVisible;
    private bool _isWhatsNewPageVisible;
    private bool _isWhatsNewExpanded;
    private bool _isHomePageVisible;
    private bool _isFileExplorerVisible;
    private bool _isFileTreeExpanded;
    private bool _isStatusBarFilePathVisible = true;
    private bool _isWordWrapEnabled;
    private bool _isConfirmBeforeClosingUnsavedTabsEnabled = true;
    private bool _isRestoreOpenTabsOnLaunchEnabled;
    private bool _isMarketplaceTabSelected;
    private bool _suppressDirtyTracking;
    private int _tabSize = 4;
    private int _editorFontSize = 14;
    private string _currentThemeName = "Dark";
    private string _requestedThemeName = "Dark";
    private string _editorStatsText = "0 lines";
    private string _lastDiscordPresenceDetails = string.Empty;
    private string _lastDiscordPresenceState = string.Empty;
    private readonly DateTime _sessionStart = DateTime.UtcNow;
    private string _extensionsStatusText = "Drop .kox extension files into the Extensions folder to install them.";
    private string _latestReleaseStatusText = "Loading latest release...";
    private ReleaseInfo? _latestRelease;
    private LoadedExtension? _currentLanguageExtension;
    private Bitmap? _currentImagePreview;
    private FileSystemWatcher? _extensionsFolderWatcher;
    private FileSystemWatcher? _projectExtensionsFolderWatcher;
    private readonly DispatcherTimer _extensionsRefreshDebounceTimer = new() { Interval = TimeSpan.FromMilliseconds(250) };
    private readonly IndentGuideBackgroundRenderer _indentGuideRenderer = new();
    private readonly List<string> _startupOpenTabPaths = [];
    private static readonly TimeSpan ExtensionsRefreshCooldown = TimeSpan.FromSeconds(8);
    private static readonly HashSet<string> ImagePreviewExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".png", ".apng", ".jpg", ".jpeg", ".jpe", ".jfif", ".bmp", ".dib", ".gif",
        ".webp", ".ico", ".cur", ".tif", ".tiff"
    };
    private DateTime _lastExtensionsRefreshUtc = DateTime.MinValue;
    private string? _startupActiveTabPath;
    private string? _startupFilePath;
    private string _extensionSearchText = string.Empty;
    private bool _isFindPanelVisible;
    private string _findText = string.Empty;

    // File tree clipboard state (cut/copy/paste)
    private string? _clipboardItemPath;
    private bool _clipboardItemIsDirectory;
    private bool _clipboardIsCut;
    private event PropertyChangedEventHandler? ViewModelPropertyChanged;
    private static readonly HttpClient MarketplaceHttpClient = CreateHttpClient();

    // ── Auto-completion ──────────────────────────────────────────────────────

    // Maps each opening character to its closing pair
    private static readonly Dictionary<char, char> BracketPairs = new()
    {
        { '(', ')' },
        { '[', ']' },
        { '{', '}' },
        { '<', '>' },
        { '"', '"' },
        { '\'', '\'' },
        { '`', '`' },
    };

    // Closing characters — when typed over an existing auto-inserted closer, skip past it
    private static readonly HashSet<char> ClosingChars = new() { ')', ']', '}', '>', '"', '\'', '`' };

    private static HttpClient CreateHttpClient()
    {
        var client = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
        client.DefaultRequestHeaders.UserAgent.ParseAdd("Kodo/1.0 (https://github.com/KerbalMissile/Kodo)");
        return client;
    }

    private string ExtensionsFolderPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "Kodo", "Extensions");
    private string ProjectExtensionsFolderPath =>
        Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "Extensions"));

    // Flat list that backs the ItemsControl – directories insert/remove their children in-place
    public ObservableCollection<FileTreeItem> FileTreeItems { get; } = new();
    public ObservableCollection<RecentFileItem> RecentFiles { get; } = new();
    public ObservableCollection<EditorTab> OpenTabs { get; } = new();
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

    public Bitmap? CurrentImagePreview
    {
        get => _currentImagePreview;
        private set
        {
            if (ReferenceEquals(_currentImagePreview, value))
                return;

            var previousPreview = _currentImagePreview;
            _currentImagePreview = value;
            previousPreview?.Dispose();
            OnPropertyChanged();
            OnPropertyChanged(nameof(HasImagePreview));
            OnPropertyChanged(nameof(IsImagePreviewVisible));
            OnPropertyChanged(nameof(IsTextEditorVisible));
        }
    }

    public EditorTab? ActiveEditorTab
    {
        get => _activeEditorTab;
        private set
        {
            if (ReferenceEquals(_activeEditorTab, value))
                return;

            if (_activeEditorTab is not null)
                _activeEditorTab.IsSelected = false;

            _activeEditorTab = value;

            if (_activeEditorTab is not null)
                _activeEditorTab.IsSelected = true;

            OnPropertyChanged();
            OnPropertyChanged(nameof(HasOpenEditors));
            OnPropertyChanged(nameof(IsEditorTabsVisible));
            SaveSettings();
        }
    }

    public MainWindow() : this(null) { }

    public MainWindow(string? startupFilePath)
    {
        var trimmedStartupPath = startupFilePath?.Trim().Trim('"');
        _startupFilePath = !string.IsNullOrWhiteSpace(trimmedStartupPath) && File.Exists(trimmedStartupPath)
            ? trimmedStartupPath
            : null;
        InitializeComponent();
        LoadWindowIcon();
        EditorTextBox.LineNumbersMargin = new Thickness(8, 0, 8, 0);
        EditorTextBox.TextArea.LeftMargins.Add(DottedLineMargin.Create());
        EditorTextBox.TextArea.TextView.BackgroundRenderers.Add(_indentGuideRenderer);
        EditorTextBox.TextArea.TextView.LineTransformers.Add(_rainbowBracketColorizer);
        EditorTextBox.TextArea.TextView.LinkTextForegroundBrush = Brush.Parse("#5BA3D9");
        EditorTextBox.TextArea.TextView.LinkTextBackgroundBrush = Brushes.Transparent;
        OpenTabs.CollectionChanged += OpenTabs_CollectionChanged;
        FileTreeItems.CollectionChanged += (_, _) => OnPropertyChanged(nameof(ExplorerPanelWidth));
        // TextEditor uses EventHandler (not RoutedEventHandler), so hook up in code-behind
        EditorTextBox.TextChanged += EditorTextBox_OnTextChanged;
        EditorTextBox.TextArea.Caret.PositionChanged += (_, _) => QueueRefreshState();
		// Auto-completion: insert closing bracket/quote after opener, skip-over when typing a closer
        EditorTextBox.TextArea.TextEntering += EditorTextArea_OnTextEntering;
        EditorTextBox.TextArea.TextEntered  += EditorTextArea_OnTextEntered;
        AddHandler(InputElement.KeyDownEvent, MainWindow_EditorKeyIntercept_OnKeyDown, RoutingStrategies.Tunnel, handledEventsToo: true);
        var settings = LoadSettings();
        _requestedThemeName = string.IsNullOrWhiteSpace(settings.ThemeName) ? "Dark" : settings.ThemeName;
        _isAutoSaveEnabled = settings.AutoSaveEnabled;
        _isDiscordRichPresenceEnabled = settings.DiscordRichPresenceEnabled;
        _isStatusBarFilePathVisible = settings.StatusBarFilePathVisible;
        _isWordWrapEnabled = settings.WordWrapEnabled;
        _isConfirmBeforeClosingUnsavedTabsEnabled = settings.ConfirmBeforeClosingUnsavedTabsEnabled;
        _isRestoreOpenTabsOnLaunchEnabled = settings.RestoreOpenTabsOnLaunchEnabled;
        _tabSize = NormalizeTabSize(settings.TabSize);
        _editorFontSize = settings.EditorFontSize is >= 8 and <= 32 ? settings.EditorFontSize : 14;
        _startupOpenTabPaths.AddRange(settings.OpenTabPaths
            .Where(path => File.Exists(path))
            .Distinct(StringComparer.OrdinalIgnoreCase));
        _startupActiveTabPath = settings.ActiveTabPath;
        LoadRecentFiles(settings.RecentFiles);
        _autoSaveTimer.Tick += AutoSaveTimer_OnTick;
        _autoSaveStatusTimer.Tick += AutoSaveStatusTimer_OnTick;
        _discordReconnectTimer.Tick += DiscordReconnectTimer_OnTick;
        _editorStateRefreshTimer.Tick += EditorStateRefreshTimer_OnTick;
        _extensionsRefreshDebounceTimer.Tick += ExtensionsRefreshDebounceTimer_OnTick;
        DataContext = this;
        IsHomePageVisible = true;
        ApplyTheme(_requestedThemeName);
        EnsureExtensionsFolder();
        SetupExtensionFolderWatchers();
        _ = RefreshExtensionsDataAsync();
        _ = RefreshLatestReleaseAsync();
        UpdateDiscordRichPresenceLifecycle();
        ApplyEditorSettings();
        Opened += MainWindow_OnOpened;
        Closing += MainWindow_OnClosing;
        Closed += MainWindow_OnClosed;
        RefreshState();
    }

    // ── Extension loading ────────────────────────────────────────────────────

    private void EnsureExtensionsFolder()
    {
        if (!Directory.Exists(ExtensionsFolderPath))
            Directory.CreateDirectory(ExtensionsFolderPath);
    }

    private void SetupExtensionFolderWatchers()
    {
        DisposeExtensionFolderWatchers();
        _extensionsFolderWatcher = CreateExtensionFolderWatcher(ExtensionsFolderPath);

        if (Directory.Exists(ProjectExtensionsFolderPath) &&
            !string.Equals(ProjectExtensionsFolderPath, ExtensionsFolderPath, StringComparison.OrdinalIgnoreCase))
        {
            _projectExtensionsFolderWatcher = CreateExtensionFolderWatcher(ProjectExtensionsFolderPath);
        }
    }

    private FileSystemWatcher CreateExtensionFolderWatcher(string path)
    {
        var watcher = new FileSystemWatcher(path)
        {
            IncludeSubdirectories = false,
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.CreationTime
        };

        watcher.Created += ExtensionFolderWatcher_OnChanged;
        watcher.Deleted += ExtensionFolderWatcher_OnChanged;
        watcher.Renamed += ExtensionFolderWatcher_OnRenamed;
        watcher.EnableRaisingEvents = true;
        return watcher;
    }

    private void DisposeExtensionFolderWatchers()
    {
        DisposeExtensionFolderWatcher(_extensionsFolderWatcher);
        DisposeExtensionFolderWatcher(_projectExtensionsFolderWatcher);
        _extensionsFolderWatcher = null;
        _projectExtensionsFolderWatcher = null;
    }

    private void DisposeExtensionFolderWatcher(FileSystemWatcher? watcher)
    {
        if (watcher is null)
            return;

        watcher.EnableRaisingEvents = false;
        watcher.Created -= ExtensionFolderWatcher_OnChanged;
        watcher.Deleted -= ExtensionFolderWatcher_OnChanged;
        watcher.Renamed -= ExtensionFolderWatcher_OnRenamed;
        watcher.Dispose();
    }

    private static bool IsExtensionFilePath(string path)
    {
        var extension = Path.GetExtension(path);
        return extension.Equals(".kox", StringComparison.OrdinalIgnoreCase) ||
               string.IsNullOrWhiteSpace(extension);
    }

    private void ExtensionFolderWatcher_OnChanged(object sender, FileSystemEventArgs e)
    {
        if (!IsExtensionFilePath(e.FullPath))
            return;

        QueueExtensionsRefresh();
    }

    private void ExtensionFolderWatcher_OnRenamed(object sender, RenamedEventArgs e)
    {
        if (!IsExtensionFilePath(e.OldFullPath) && !IsExtensionFilePath(e.FullPath))
            return;

        QueueExtensionsRefresh();
    }

    private void QueueExtensionsRefresh()
    {
        Dispatcher.UIThread.Post(() =>
        {
            _extensionsRefreshDebounceTimer.Stop();
            _extensionsRefreshDebounceTimer.Start();
        });
    }

    private async void ExtensionsRefreshDebounceTimer_OnTick(object? sender, EventArgs e)
    {
        _extensionsRefreshDebounceTimer.Stop();
        await RefreshExtensionsDataAsync();
    }

    private async Task RefreshExtensionsDataAsync(bool force = false)
    {
        if (_isRefreshingExtensions)
            return;

        if (!force && DateTime.UtcNow - _lastExtensionsRefreshUtc < ExtensionsRefreshCooldown)
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
            var updateCount = MarketplaceExtensions.Count(e => e.IsUpdateAvailable);
            ExtensionsStatusText = updateCount > 0
                ? $"Refreshed {VisibleLoadedExtensions.Count()} installed and {MarketplaceExtensions.Count} marketplace extension(s). {updateCount} update(s) available."
                : $"Refreshed {VisibleLoadedExtensions.Count()} installed and {MarketplaceExtensions.Count} marketplace extension(s).";
            _lastExtensionsRefreshUtc = DateTime.UtcNow;
        }
        catch (Exception ex)
        {
            ExtensionsStatusText = $"Refresh failed: {ex.Message}";
            await ShowWarningDialogAsync("Extension refresh", ex);
        }
        finally
        {
            IsRefreshingExtensions = false;
        }
    }

    private void LoadExtensions()
    {
        var loadedExtensions = new List<LoadedExtension>();
        var extensionLoadErrors = new List<string>();
        var searchPaths = GetExtensionSearchPaths().ToList();

        var anyFolderFound = false;
        foreach (var searchPath in searchPaths)
        {
            if (!Directory.Exists(searchPath)) continue;
            anyFolderFound = true;

            foreach (var koxFile in Directory.GetFiles(searchPath, "*.kox"))
            {
                try
                {
                    foreach (var ext in LoadExtensionsFromKox(koxFile))
                    {
                        AddOrReplaceLoadedExtension(loadedExtensions, ext);
                    }
                }
                catch (Exception ex)
                {
                    extensionLoadErrors.Add($"Failed to load '{Path.GetFileName(koxFile)}': {ex.Message}");
                }
            }

            foreach (var dir in Directory.GetDirectories(searchPath))
            {
                try
                {
                    if (File.Exists(Path.Combine(dir, "manifest.json")))
                    {
                        foreach (var ext in LoadExtensionsFromFolder(dir))
                        {
                            AddOrReplaceLoadedExtension(loadedExtensions, ext);
                        }
                    }
                }
                catch (Exception ex)
                {
                    extensionLoadErrors.Add($"Failed to load folder extension '{Path.GetFileName(dir)}': {ex.Message}");
                }
            }
        }

        if (!anyFolderFound)
            extensionLoadErrors.Add($"Extensions folder not found. Expected: {ExtensionsFolderPath}");

        SyncObservableCollection(LoadedExtensions, loadedExtensions, ext => ext.Id);
        SyncObservableCollection(ExtensionLoadErrors, extensionLoadErrors, error => error);

        OnPropertyChanged(nameof(ExtensionLoadErrors));
        OnPropertyChanged(nameof(VisibleLoadedExtensions));
        OnPropertyChanged(nameof(FilteredInstalledExtensions));
        OnPropertyChanged(nameof(IsNoExtensionsVisible));
        OnPropertyChanged(nameof(ThemeExtensions));
        OnPropertyChanged(nameof(HasThemeExtensions));
        RefreshExtensionTheme();
        SyncMarketplaceInstallStates();
    }

    private async Task LoadMarketplaceExtensionsAsync()
    {
        var marketplaceExtensions = new List<MarketplaceExtension>();
        var extensionLoadErrors = new List<string>();

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
                    if (string.IsNullOrWhiteSpace(entry.Id) || marketplaceExtensions.Any(e => e.Id == entry.Id))
                        continue;

                    marketplaceExtensions.Add(entry);
                }
            }
        }
        catch (Exception ex)
        {
            extensionLoadErrors.Add($"Failed to load remote marketplace index: {ex.Message}");
            await ShowWarningDialogAsync("Marketplace fetch", ex);
        }

        SyncObservableCollection(MarketplaceExtensions, marketplaceExtensions, ext => ext.Id);
        SyncObservableCollection(
            ExtensionLoadErrors,
            ExtensionLoadErrors.Concat(extensionLoadErrors).Distinct().ToList(),
            error => error);

        SyncMarketplaceInstallStates();
        OnPropertyChanged(nameof(ExtensionLoadErrors));
        OnPropertyChanged(nameof(IsMarketplaceEmptyVisible));
        await FetchMarketplaceIconsAsync();
        await FetchInstalledExtensionIconsAsync();
    }

    private async Task FetchInstalledExtensionIconsAsync()
    {
        // For each installed extension, look up its IconUrl from the marketplace index
        // and fetch it so the Installed tab shows the same artwork as the Marketplace tab.
        var tasks = LoadedExtensions
            .Select(ext =>
            {
                var marketplaceEntry = MarketplaceExtensions.FirstOrDefault(m =>
                    m.Id.Equals(ext.Id, StringComparison.OrdinalIgnoreCase));
                return (ext, iconUrl: marketplaceEntry?.IconUrl ?? string.Empty);
            })
            .Where(pair => !string.IsNullOrWhiteSpace(pair.iconUrl))
            .Select(async pair =>
            {
                try
                {
                    var bytes = await MarketplaceHttpClient.GetByteArrayAsync(pair.iconUrl);
                    await Dispatcher.UIThread.InvokeAsync(() =>
                    {
                        try
                        {
                            using var ms = new MemoryStream(bytes);
                            pair.ext.IconImage = new Bitmap(ms);
                            pair.ext.NotifyIconChanged();
                        }
                        catch { /* Bad image data — keep existing icon or abbreviation. */ }
                    });
                }
                catch { /* Network failure — keep existing icon or abbreviation. */ }
            });

        await Task.WhenAll(tasks);
    }

    private async Task FetchMarketplaceIconsAsync()
    {
        // Fetch the icon from IconUrl for every entry that has one.
        // For installed extensions, the remote icon takes priority over the local
        // .kox icon so the marketplace always shows the index-supplied artwork.
        var tasks = MarketplaceExtensions
            .Where(e => !string.IsNullOrWhiteSpace(e.IconUrl))
            .Select(async entry =>
            {
                try
                {
                    var bytes = await MarketplaceHttpClient.GetByteArrayAsync(entry.IconUrl);
                    await Dispatcher.UIThread.InvokeAsync(() =>
                    {
                        try
                        {
                            using var ms = new MemoryStream(bytes);
                            entry.IconImage = new Bitmap(ms);
                        }
                        catch { /* Silently ignore bad image data — fallback to abbreviation. */ }
                    });
                }
                catch { /* Network failure — fallback to local icon or abbreviation. */ }
            });

        await Task.WhenAll(tasks);
    }

    private async Task RefreshLatestReleaseAsync()
    {
        if (_isRefreshingLatestRelease)
            return;

        _isRefreshingLatestRelease = true;
        OnPropertyChanged(nameof(IsRefreshingLatestRelease));
        OnPropertyChanged(nameof(RefreshLatestReleaseButtonText));
        LatestReleaseStatusText = "Loading latest release...";

        try
        {
            LatestRelease = await FetchLatestReleaseInfoAsync();

            LatestReleaseStatusText = HasLatestRelease
                ? $"Latest release: {LatestReleaseDisplayName}"
                : "No releases found.";
        }
        catch (Exception ex)
        {
            LatestRelease = null;
            LatestReleaseStatusText = $"Could not load release info: {ex.Message}";
        }
        finally
        {
            _isRefreshingLatestRelease = false;
            OnPropertyChanged(nameof(IsRefreshingLatestRelease));
            OnPropertyChanged(nameof(RefreshLatestReleaseButtonText));
        }
    }

    private async Task<ReleaseInfo?> FetchLatestReleaseInfoAsync()
    {
        var latestRelease = await TryFetchLatestStableReleaseAsync();
        if (latestRelease is not null)
            return latestRelease;

        return await TryFetchLatestListedReleaseAsync();
    }

    private async Task<ReleaseInfo?> TryFetchLatestStableReleaseAsync()
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, LatestReleaseApiUrl);
        request.Headers.Accept.ParseAdd("application/vnd.github+json");

        using var response = await MarketplaceHttpClient.SendAsync(request);
        if (!response.IsSuccessStatusCode)
            return null;

        var json = await response.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(json);
        return ParseReleaseInfo(doc.RootElement);
    }

    private async Task<ReleaseInfo?> TryFetchLatestListedReleaseAsync()
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, ReleasesApiUrl);
        request.Headers.Accept.ParseAdd("application/vnd.github+json");

        using var response = await MarketplaceHttpClient.SendAsync(request);
        response.EnsureSuccessStatusCode();

        var json = await response.Content.ReadAsStringAsync();
        using var doc = JsonDocument.Parse(json);
        if (doc.RootElement.ValueKind != JsonValueKind.Array)
            return null;

        foreach (var release in doc.RootElement.EnumerateArray())
        {
            var parsedRelease = ParseReleaseInfo(release);
            if (parsedRelease is not null)
                return parsedRelease;
        }

        return null;
    }

    private static ReleaseInfo? ParseReleaseInfo(JsonElement releaseElement)
    {
        if (releaseElement.ValueKind != JsonValueKind.Object)
            return null;

        var name = releaseElement.TryGetProperty("name", out var nameElement)
            ? nameElement.GetString() ?? string.Empty
            : string.Empty;
        var tag = releaseElement.TryGetProperty("tag_name", out var tagElement)
            ? tagElement.GetString() ?? string.Empty
            : string.Empty;
        var notes = releaseElement.TryGetProperty("body", out var bodyElement)
            ? bodyElement.GetString() ?? string.Empty
            : string.Empty;
        var url = releaseElement.TryGetProperty("html_url", out var urlElement)
            ? urlElement.GetString() ?? ReleasesPageUrl
            : ReleasesPageUrl;

        if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(tag) && string.IsNullOrWhiteSpace(notes))
            return null;

        return new ReleaseInfo
        {
            Name = name,
            Tag = tag,
            Notes = notes,
            Url = url
        };
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
        FileName = item.TryGetProperty("fileName", out var fileName) ? fileName.GetString() ?? string.Empty : string.Empty,
        IconUrl = item.TryGetProperty("iconUrl", out var iconUrl) ? iconUrl.GetString() ?? string.Empty : string.Empty
    };

    private void SyncMarketplaceInstallStates()
    {
        foreach (var installedExtension in LoadedExtensions)
            installedExtension.IsUpdateAvailable = false;

        foreach (var entry in MarketplaceExtensions)
        {
            var localExt = GetPreferredLoadedExtension(entry.Id);
            var isUpdateAvailable = localExt is not null && CompareExtensionVersions(entry.Version, localExt.Version) > 0;

            entry.SetInstalledState(localExt, isUpdateAvailable);
            if (localExt is not null)
                localExt.IsUpdateAvailable = isUpdateAvailable;

            // When installed, prefer the local icon from the .kox file over the remote IconUrl.
            if (localExt?.IconImage is not null)
                entry.IconImage = localExt.IconImage;
        }

        OnPropertyChanged(nameof(AvailableExtensionUpdatesCount));
        OnPropertyChanged(nameof(IsExtensionUpdateBannerVisible));
        OnPropertyChanged(nameof(ExtensionUpdatesBannerText));
    }

    private MarketplaceExtension? GetMarketplaceExtensionForInstalled(LoadedExtension extension) =>
        MarketplaceExtensions.FirstOrDefault(entry =>
            entry.Id.Equals(extension.Id, StringComparison.OrdinalIgnoreCase));

    private void AddOrReplaceLoadedExtension(IList<LoadedExtension> extensions, LoadedExtension extension)
    {
        var existingIndex = extensions
            .Select((item, index) => new { item, index })
            .FirstOrDefault(x => x.item.Id.Equals(extension.Id, StringComparison.OrdinalIgnoreCase));

        if (existingIndex is null)
        {
            extensions.Add(extension);
            return;
        }

        if (ShouldReplaceLoadedExtension(existingIndex.item, extension))
            extensions[existingIndex.index] = extension;
    }

    private static void SyncObservableCollection<T, TKey>(
        ObservableCollection<T> target,
        IList<T> source,
        Func<T, TKey> keySelector)
        where TKey : notnull
    {
        for (var i = target.Count - 1; i >= 0; i--)
        {
            var key = keySelector(target[i]);
            if (!source.Any(item => EqualityComparer<TKey>.Default.Equals(keySelector(item), key)))
                target.RemoveAt(i);
        }

        for (var i = 0; i < source.Count; i++)
        {
            var item = source[i];
            var key = keySelector(item);
            var existingIndex = -1;
            for (var j = 0; j < target.Count; j++)
            {
                if (EqualityComparer<TKey>.Default.Equals(keySelector(target[j]), key))
                {
                    existingIndex = j;
                    break;
                }
            }

            if (existingIndex == -1)
            {
                target.Insert(Math.Min(i, target.Count), item);
                continue;
            }

            if (existingIndex != i)
                target.Move(existingIndex, i);

            if (!ReferenceEquals(target[i], item))
                target[i] = item;
        }
    }

    private LoadedExtension? GetPreferredLoadedExtension(string extensionId) =>
        LoadedExtensions
            .Where(ext => ext.Id.Equals(extensionId, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(GetLoadedExtensionSourcePriority)
            .ThenByDescending(ext => ParseVersionNumbers(ext.Version), VersionNumberSequenceComparer.Instance)
            .FirstOrDefault();

    private bool ShouldReplaceLoadedExtension(LoadedExtension current, LoadedExtension candidate)
    {
        var currentPriority = GetLoadedExtensionSourcePriority(current);
        var candidatePriority = GetLoadedExtensionSourcePriority(candidate);
        if (candidatePriority != currentPriority)
            return candidatePriority > currentPriority;

        return CompareExtensionVersions(candidate.Version, current.Version) > 0;
    }

    private int GetLoadedExtensionSourcePriority(LoadedExtension extension)
    {
        if (!string.IsNullOrWhiteSpace(extension.SourcePath))
        {
            var sourcePath = Path.GetFullPath(extension.SourcePath);
            if (IsPathInsideDirectory(sourcePath, ExtensionsFolderPath))
                return 2;

            if (IsPathInsideDirectory(sourcePath, ProjectExtensionsFolderPath))
                return 1;
        }

        return 0;
    }

    private async Task InstallMarketplaceExtensionAsync(MarketplaceExtension marketplaceExtension)
    {
        if (marketplaceExtension.IsInstalling || (marketplaceExtension.IsInstalled && !marketplaceExtension.IsUpdateAvailable))
            return;

        marketplaceExtension.IsInstalling = true;
        var action = marketplaceExtension.IsUpdateAvailable ? "Updating" : "Installing";
        marketplaceExtension.InstallButtonText = $"{action}...";
        ExtensionsStatusText = $"{action} {marketplaceExtension.Name}...";

        try
        {
            EnsureExtensionsFolder();
            var wasUpdate = marketplaceExtension.IsUpdateAvailable;
            var installedExtension = GetPreferredLoadedExtension(marketplaceExtension.Id);
            var outputPath = ResolveExtensionInstallPath(marketplaceExtension, installedExtension);
            var bytes = await MarketplaceHttpClient.GetByteArrayAsync(marketplaceExtension.DownloadUrl);
            DeleteInstalledExtensionSources(marketplaceExtension.Id, outputPath);
            await File.WriteAllBytesAsync(outputPath, bytes);
            NormalizeKoxManifestVersion(outputPath);

            await RefreshExtensionsDataAsync(force: true);
            ExtensionsStatusText = $"{marketplaceExtension.Name} {(wasUpdate ? "updated" : "installed")}.";
        }
        catch (Exception ex)
        {
            marketplaceExtension.SetInstalledState(
                GetPreferredLoadedExtension(marketplaceExtension.Id),
                marketplaceExtension.IsUpdateAvailable);
            ExtensionsStatusText = $"Failed to install {marketplaceExtension.Name}: {ex.Message}";
            await ShowWarningDialogAsync($"Extension install — {marketplaceExtension.Name}", ex);
        }
        finally
        {
            marketplaceExtension.IsInstalling = false;
            SyncMarketplaceInstallStates();
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

            await RefreshExtensionsDataAsync(force: true);
            ExtensionsStatusText = $"{extension.Name} uninstalled.";
        }
        catch (Exception ex)
        {
            ExtensionsStatusText = $"Failed to uninstall {extension.Name}: {ex.Message}";
            await ShowWarningDialogAsync($"Extension uninstall — {extension.Name}", ex);
        }
    }

    private static bool IsPathInsideDirectory(string path, string directory)
    {
        var normalizedPath = Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        var normalizedDirectory = Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        return normalizedPath.StartsWith(normalizedDirectory + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
            || string.Equals(normalizedPath, normalizedDirectory, StringComparison.OrdinalIgnoreCase);
    }

    private IEnumerable<string> GetExtensionSearchPaths()
    {
        yield return ExtensionsFolderPath;

        // Also search the project source tree when running from the build output directory
        var projectRoot = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..");
        var srcPath = Path.GetFullPath(Path.Combine(projectRoot, "Extensions"));
        if (!string.Equals(srcPath, ExtensionsFolderPath, StringComparison.OrdinalIgnoreCase))
            yield return srcPath;
    }

    private IEnumerable<(string Path, bool IsDirectory)> EnumerateInstalledExtensionSources(string extensionId)
    {
        foreach (var searchPath in GetExtensionSearchPaths())
        {
            if (!Directory.Exists(searchPath))
                continue;

            foreach (var koxFile in Directory.GetFiles(searchPath, "*.kox"))
            {
                if (ExtensionSourceMatchesId(koxFile, extensionId, isDirectory: false))
                    yield return (koxFile, false);
            }

            foreach (var dir in Directory.GetDirectories(searchPath))
            {
                if (ExtensionSourceMatchesId(dir, extensionId, isDirectory: true))
                    yield return (dir, true);
            }
        }
    }

    private void DeleteInstalledExtensionSources(string extensionId, string? pathToKeep = null)
    {
        foreach (var source in EnumerateInstalledExtensionSources(extensionId))
        {
            var resolvedPath = Path.GetFullPath(source.Path);
            if (!string.IsNullOrWhiteSpace(pathToKeep) &&
                string.Equals(resolvedPath, Path.GetFullPath(pathToKeep), StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!IsPathInsideDirectory(resolvedPath, ExtensionsFolderPath) &&
                !IsPathInsideDirectory(resolvedPath, ProjectExtensionsFolderPath))
            {
                continue;
            }

            if (source.IsDirectory)
            {
                if (Directory.Exists(resolvedPath))
                    Directory.Delete(resolvedPath, recursive: true);
            }
            else
            {
                if (File.Exists(resolvedPath))
                    File.Delete(resolvedPath);
            }
        }
    }

    private string ResolveExtensionInstallPath(MarketplaceExtension marketplaceExtension, LoadedExtension? installedExtension)
    {
        if (installedExtension is not null &&
            !installedExtension.IsDirectorySource &&
            !string.IsNullOrWhiteSpace(installedExtension.SourcePath))
        {
            var sourcePath = Path.GetFullPath(installedExtension.SourcePath);
            if (IsPathInsideDirectory(sourcePath, ExtensionsFolderPath) ||
                IsPathInsideDirectory(sourcePath, ProjectExtensionsFolderPath))
            {
                return sourcePath;
            }
        }

        var fileName = string.IsNullOrWhiteSpace(marketplaceExtension.FileName)
            ? TryGetFileNameFromUrl(marketplaceExtension.DownloadUrl)
            : marketplaceExtension.FileName;

        return Path.Combine(ExtensionsFolderPath, fileName);
    }

    private static string TryGetFileNameFromUrl(string url)
    {
        try
        {
            return Path.GetFileName(new Uri(url).AbsolutePath);
        }
        catch
        {
            return "extension.kox";
        }
    }

    private static bool ExtensionSourceMatchesId(string path, string extensionId, bool isDirectory)
    {
        try
        {
            var manifest = isDirectory
                ? ReadManifestFromFolder(path)
                : ReadManifestFromKox(path);

            return manifest.TryGetProperty("id", out var id) &&
                   string.Equals(id.GetString(), extensionId, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static JsonElement ReadManifestFromFolder(string folderPath)
    {
        using var manifestDoc = JsonDocument.Parse(File.ReadAllText(Path.Combine(folderPath, "manifest.json")));
        return manifestDoc.RootElement.Clone();
    }

    private static JsonElement ReadManifestFromKox(string koxPath)
    {
        using var archive = ZipFile.OpenRead(koxPath);
        var manifestEntry = archive.GetEntry("manifest.json")
            ?? throw new InvalidDataException($"Missing manifest.json in '{koxPath}'.");
        using var manifestStream = manifestEntry.Open();
        using var manifestDoc = JsonDocument.Parse(manifestStream);
        return manifestDoc.RootElement.Clone();
    }

    private static void NormalizeKoxManifestVersion(string koxPath)
    {
        try
        {
            var inferredVersion = ExtractVersionFromName(Path.GetFileName(koxPath));
            if (string.IsNullOrWhiteSpace(inferredVersion) || !File.Exists(koxPath))
                return;

            using var archive = ZipFile.Open(koxPath, ZipArchiveMode.Update);
            var manifestEntry = archive.GetEntry("manifest.json");
            if (manifestEntry is null)
                return;

            JsonObject? manifestObject;
            using (var manifestStream = manifestEntry.Open())
            using (var reader = new StreamReader(manifestStream))
            {
                manifestObject = JsonNode.Parse(reader.ReadToEnd()) as JsonObject;
            }

            if (manifestObject is null)
                return;

            var manifestVersion = manifestObject["version"]?.GetValue<string>() ?? string.Empty;
            if (CompareExtensionVersions(inferredVersion, manifestVersion) <= 0)
                return;

            manifestObject["version"] = inferredVersion;
            manifestEntry.Delete();
            var newManifestEntry = archive.CreateEntry("manifest.json");
            using var outputStream = newManifestEntry.Open();
            using var writer = new Utf8JsonWriter(outputStream, new JsonWriterOptions { Indented = true });
            manifestObject.WriteTo(writer);
            writer.Flush();
        }
        catch
        {
            // If we cannot normalize the package metadata, fall back to reading it as-is.
        }
    }

    private static int CompareExtensionVersions(string left, string right)
    {
        var leftParts = ParseVersionNumbers(left);
        var rightParts = ParseVersionNumbers(right);
        return VersionNumberSequenceComparer.Instance.Compare(leftParts, rightParts);
    }

    private static string GetBestKnownExtensionVersion(string manifestVersion, string sourcePath)
    {
        var inferredVersion = ExtractVersionFromName(Path.GetFileName(sourcePath));
        return CompareExtensionVersions(inferredVersion, manifestVersion) > 0
            ? inferredVersion
            : manifestVersion;
    }

    private static string ExtractVersionFromName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return string.Empty;

        var match = Regex.Match(name, @"(?i)(v\d+(?:\.\d+)+)");
        return match.Success ? match.Groups[1].Value : string.Empty;
    }

    private static int[] ParseVersionNumbers(string version)
    {
        if (string.IsNullOrWhiteSpace(version))
            return [0];

        var matches = Regex.Matches(version, @"\d+");
        if (matches.Count == 0)
            return [0];

        return matches
            .Select(match => int.TryParse(match.Value, out var part) ? part : 0)
            .ToArray();
    }

    private sealed class VersionNumberSequenceComparer : IComparer<int[]>
    {
        public static VersionNumberSequenceComparer Instance { get; } = new();

        public int Compare(int[]? left, int[]? right)
        {
            left ??= [0];
            right ??= [0];

            var maxLength = Math.Max(left.Length, right.Length);
            for (var i = 0; i < maxLength; i++)
            {
                var leftPart = i < left.Length ? left[i] : 0;
                var rightPart = i < right.Length ? right[i] : 0;
                var comparison = leftPart.CompareTo(rightPart);
                if (comparison != 0)
                    return comparison;
            }

            return 0;
        }
    }

    private IEnumerable<LoadedExtension> LoadExtensionsFromFolder(string folderPath)
    {
        var manifestPath = Path.Combine(folderPath, "manifest.json");
        if (!File.Exists(manifestPath)) yield break;

        using var manifestDoc = JsonDocument.Parse(File.ReadAllText(manifestPath));
        var baseExt = ParseManifest(manifestDoc.RootElement);
        baseExt = baseExt with { Version = GetBestKnownExtensionVersion(baseExt.Version, folderPath) };
        baseExt.SourcePath = folderPath;
        baseExt.IsDirectorySource = true;

        var languagePath = Path.Combine(folderPath, "language.json");
        if (File.Exists(languagePath))
        {
            using var langDoc = JsonDocument.Parse(File.ReadAllText(languagePath));
            ParseLanguage(langDoc.RootElement, baseExt);
        }

        var language2Path = Path.Combine(folderPath, "language2.json");
        if (File.Exists(language2Path))
        {
            using var lang2Doc = JsonDocument.Parse(File.ReadAllText(language2Path));
            ParseLanguage(lang2Doc.RootElement, baseExt);
        }

        var iconPath = Path.Combine(folderPath, "icon.png");
        if (File.Exists(iconPath))
        {
            using var iconStream = File.OpenRead(iconPath);
            baseExt.IconImage = LoadIconFromStream(iconStream);
        }

        var themePath = Path.Combine(folderPath, "theme.json");
        if (!File.Exists(themePath))
        {
            // No theme file — yield the extension as-is (language extension, etc.)
            yield return baseExt;
            yield break;
        }

        using var themeDoc = JsonDocument.Parse(File.ReadAllText(themePath));
        var root = themeDoc.RootElement;

        if (root.ValueKind == JsonValueKind.Array)
        {
            // One LoadedExtension per theme entry in the array
            var index = 0;
            foreach (var themeElement in root.EnumerateArray())
            {
                var def = ParseTheme(themeElement, baseExt);
                var entry = CloneBaseExtension(baseExt);
                entry.ThemeDefinition = def;
                // Make the Id unique so duplicate-checking works correctly.
                // Mark index > 0 entries so they're hidden from the Installed list.
                if (index > 0)
                    entry = entry with { Id = $"{baseExt.Id}_{def.ThemeId}", IsThemeSubEntry = true };
                yield return entry;
                index++;
            }
        }
        else
        {
            baseExt.ThemeDefinition = ParseTheme(root, baseExt);
            yield return baseExt;
        }
    }

    // Shallow-clones a LoadedExtension so each theme entry gets its own object
    private static LoadedExtension CloneBaseExtension(LoadedExtension src) => new()
    {
        Id                = src.Id,
        Version           = src.Version,
        Name              = src.Name,
        Type              = src.Type,
        Author            = src.Author,
        Description       = src.Description,
        Extensions        = src.Extensions,
        Keywords          = src.Keywords,
        Types             = src.Types,
        CommentLine       = src.CommentLine,
        CommentBlockStart = src.CommentBlockStart,
        CommentBlockEnd   = src.CommentBlockEnd,
        StringDelimiters  = src.StringDelimiters.ToArray(),
        MultiLineStringDelimiters = src.MultiLineStringDelimiters.ToArray(),
        DisableSingleQuoteStrings = src.DisableSingleQuoteStrings,
        ColorTokens       = new Dictionary<string, string>(src.ColorTokens),
        SourcePath        = src.SourcePath,
        IsDirectorySource = src.IsDirectorySource,
        IconImage         = src.IconImage,
    };

    // Loads a PNG from a stream and scales it to 48x48 if it is square,
    // otherwise returns null so the text fallback is used.
    private static Bitmap? LoadIconFromStream(Stream stream)
    {
        try
        {
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            ms.Position = 0;
            var bmp = new Bitmap(ms);
            if (bmp.PixelSize.Width != bmp.PixelSize.Height) return null;
            // Scale down to 48x48 — Avalonia Bitmap doesn't resize on load,
            // but we can let the Image control handle it via Width/Height binding.
            // Just return the bitmap as-is; sizing is done in AXAML.
            return bmp;
        }
        catch
        {
            return null;
        }
    }

    private IEnumerable<LoadedExtension> LoadExtensionsFromKox(string koxPath)
    {
        using var archive = ZipFile.OpenRead(koxPath);
        var manifestEntry = archive.GetEntry("manifest.json");
        if (manifestEntry is null) yield break;

        using var manifestStream = manifestEntry.Open();
        using var manifestDoc = JsonDocument.Parse(manifestStream);
        var baseExt = ParseManifest(manifestDoc.RootElement);
        baseExt = baseExt with { Version = GetBestKnownExtensionVersion(baseExt.Version, koxPath) };
        baseExt.SourcePath = koxPath;
        baseExt.IsDirectorySource = false;

        var languageEntry = archive.GetEntry("language.json");
        if (languageEntry is not null)
        {
            using var langStream = languageEntry.Open();
            using var langDoc = JsonDocument.Parse(langStream);
            ParseLanguage(langDoc.RootElement, baseExt);
        }

        var language2Entry = archive.GetEntry("language2.json");
        if (language2Entry is not null)
        {
            using var lang2Stream = language2Entry.Open();
            using var lang2Doc = JsonDocument.Parse(lang2Stream);
            ParseLanguage(lang2Doc.RootElement, baseExt);
        }

        var iconEntry = archive.GetEntry("icon.png");
        if (iconEntry is not null)
        {
            using var iconStream = iconEntry.Open();
            baseExt.IconImage = LoadIconFromStream(iconStream);
        }

        var themeEntry = archive.GetEntry("theme.json");
        if (themeEntry is null)
        {
            yield return baseExt;
            yield break;
        }

        // ZipArchiveEntry streams are forward-only — read to memory first so we can
        // enumerate the JSON array without the stream closing under us.
        using var themeStream = themeEntry.Open();
        using var ms = new MemoryStream();
        themeStream.CopyTo(ms);
        ms.Position = 0;
        using var themeDoc = JsonDocument.Parse(ms);
        var root = themeDoc.RootElement;

        if (root.ValueKind == JsonValueKind.Array)
        {
            var index = 0;
            foreach (var themeElement in root.EnumerateArray())
            {
                var def = ParseTheme(themeElement, baseExt);
                var entry = CloneBaseExtension(baseExt);
                entry.ThemeDefinition = def;
                if (index > 0)
                    entry = entry with { Id = $"{baseExt.Id}_{def.ThemeId}", IsThemeSubEntry = true };
                yield return entry;
                index++;
            }
        }
        else
        {
            baseExt.ThemeDefinition = ParseTheme(root, baseExt);
            yield return baseExt;
        }
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
        var profile = ParseLanguageProfile(lang);
        if (profile.Extensions.Length > 0)
        {
            ext.SyntaxProfiles.Add(profile);
            return;
        }

        ApplyLanguageProfile(ext, profile);
    }

    private static LanguageSyntaxProfile ParseLanguageProfile(JsonElement lang)
    {
        var profile = new LanguageSyntaxProfile
        {
            Extensions = ReadStringArray(lang, "extensions"),
            Keywords = ReadStringArray(lang, "keywords"),
            Types = ReadStringArray(lang, "types"),
            CommentLine = lang.TryGetProperty("commentLine", out var cl) ? cl.GetString() ?? string.Empty : null,
            CommentBlockStart = lang.TryGetProperty("commentBlockStart", out var cbs) ? cbs.GetString() ?? string.Empty : null,
            CommentBlockEnd = lang.TryGetProperty("commentBlockEnd", out var cbe) ? cbe.GetString() ?? string.Empty : null,
            StringDelimiters = lang.TryGetProperty("stringDelimiters", out var sd) ? ReadStringArray(sd) : null,
            MultiLineStringDelimiters = lang.TryGetProperty("multiLineStringDelimiters", out var msd) ? ReadStringArray(msd) : null,
            DisableSingleQuoteStrings = lang.TryGetProperty("disableSingleQuoteStrings", out var dsqs) && dsqs.ValueKind is JsonValueKind.True or JsonValueKind.False
                ? dsqs.GetBoolean()
                : null,
            ColorTokens = ReadColorTokens(lang)
        };

        return profile;
    }

    private static void ApplyLanguageProfile(LoadedExtension ext, LanguageSyntaxProfile profile)
    {
        ext.Keywords = ext.Keywords.Union(profile.Keywords).ToArray();
        ext.Types = ext.Types.Union(profile.Types).ToArray();

        if (profile.CommentLine is not null)
            ext.CommentLine = profile.CommentLine;
        if (profile.CommentBlockStart is not null)
            ext.CommentBlockStart = profile.CommentBlockStart;
        if (profile.CommentBlockEnd is not null)
            ext.CommentBlockEnd = profile.CommentBlockEnd;
        if (profile.StringDelimiters is not null)
            ext.StringDelimiters = profile.StringDelimiters.ToArray();
        if (profile.MultiLineStringDelimiters is not null)
            ext.MultiLineStringDelimiters = profile.MultiLineStringDelimiters.ToArray();
        if (profile.DisableSingleQuoteStrings.HasValue)
            ext.DisableSingleQuoteStrings = profile.DisableSingleQuoteStrings.Value;

        foreach (var (key, value) in profile.ColorTokens)
            ext.ColorTokens[key] = value;
    }

    private static string[] ReadStringArray(JsonElement root, string propertyName) =>
        root.TryGetProperty(propertyName, out var value) ? ReadStringArray(value) : [];

    private static string[] ReadStringArray(JsonElement value) =>
        value.ValueKind == JsonValueKind.Array
            ? value.EnumerateArray()
                .Select(e => e.GetString() ?? string.Empty)
                .Where(e => !string.IsNullOrWhiteSpace(e))
                .ToArray()
            : [];

    private static Dictionary<string, string> ReadColorTokens(JsonElement lang)
    {
        var colorTokens = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!lang.TryGetProperty("colorTokens", out var ct) || ct.ValueKind != JsonValueKind.Object)
            return colorTokens;

        foreach (var prop in ct.EnumerateObject())
            colorTokens[prop.Name] = prop.Value.GetString() ?? "#FFFFFF";

        return colorTokens;
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
        if (IsPlainTextFile(filePath))
            return null;

        var fileExt = Path.GetExtension(filePath).ToLowerInvariant();
        var extension = LoadedExtensions.FirstOrDefault(e =>
            e.Type == "language" &&
            e.Extensions.Any(ex => ex.Equals(fileExt, StringComparison.OrdinalIgnoreCase)));

        if (extension is null)
            return null;

        var matchingProfiles = extension.SyntaxProfiles
            .Where(profile => profile.Extensions.Any(ex => ex.Equals(fileExt, StringComparison.OrdinalIgnoreCase)))
            .ToList();

        if (matchingProfiles.Count == 0)
            return extension;

        var effectiveExtension = CloneBaseExtension(extension);
        if (matchingProfiles.Any(profile => profile.Keywords.Length > 0))
            effectiveExtension.Keywords = [];
        if (matchingProfiles.Any(profile => profile.Types.Length > 0))
            effectiveExtension.Types = [];

        foreach (var profile in matchingProfiles)
            ApplyLanguageProfile(effectiveExtension, profile);

        return effectiveExtension;
    }

    private static bool IsPlainTextFile(string? filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            return false;

        var ext = Path.GetExtension(filePath);
        return ext.Equals(".txt", StringComparison.OrdinalIgnoreCase) ||
               ext.Equals(".text", StringComparison.OrdinalIgnoreCase) ||
               ext.Equals(".log", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsImagePreviewFile(string? filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            return false;

        var ext = Path.GetExtension(filePath);
        return ImagePreviewExtensions.Contains(ext);
    }

    private static Bitmap? TryLoadImagePreview(string? filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath) || !IsImagePreviewFile(filePath))
            return null;

        try
        {
            using var stream = File.OpenRead(filePath);
            return new Bitmap(stream);
        }
        catch
        {
            return null;
        }
    }

    private bool IsPlainTextMode()
    {
        if (HasImagePreview)
            return true;

        if (IsPlainTextFile(_currentFilePath))
            return true;

        return _currentFilePath is null && _hasUntitledDocument;
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
        EditorTextBox.TextArea.TextView.LinkTextForegroundBrush = Brush.Parse("#5BA3D9");
        EditorTextBox.TextArea.TextView.LinkTextBackgroundBrush = Brushes.Transparent;
        _indentGuideRenderer.GuideBrush = MutedTextBrush.ToImmutable() is ISolidColorBrush mutedBrush
            ? new SolidColorBrush(mutedBrush.Color, 0.4)
            : new SolidColorBrush(Color.Parse("#808080"), 0.4);
        EditorTextBox.TextArea.TextView.InvalidateLayer(KnownLayer.Background);
    }

    private void ApplyEditorSettings()
    {
        if (EditorTextBox is null)
        {
            return;
        }

        EditorTextBox.WordWrap = IsWordWrapEnabled;
        EditorTextBox.Options.IndentationSize = TabSize;
        EditorTextBox.FontSize = EditorFontSize;
        _indentGuideRenderer.TabSize = TabSize;
        EditorTextBox.TextArea.TextView.InvalidateLayer(KnownLayer.Background);
    }

    private void ApplySyntaxHighlighting(LoadedExtension ext)
    {
        if (EditorTextBox is null) return;
        EditorTextBox.SyntaxHighlighting = new KodoHighlightingDefinition(ext);
        ConfigureRainbowBrackets(ext);
    }

    private void RefreshCurrentFileSyntaxHighlighting()
    {
        if (EditorTextBox is null)
            return;

        if (string.IsNullOrWhiteSpace(_currentFilePath))
        {
            CurrentLanguageExtension = null;
            EditorTextBox.SyntaxHighlighting = null;
            ConfigureRainbowBrackets(null);
            return;
        }

        var langExt = GetLanguageExtension(_currentFilePath);
        CurrentLanguageExtension = langExt;

        if (langExt is null)
        {
            EditorTextBox.SyntaxHighlighting = null;
            ConfigureRainbowBrackets(null);
        }
        else
            ApplySyntaxHighlighting(langExt);
    }

    private void UpdateCurrentDocumentPresentation()
    {
        var imagePreview = TryLoadImagePreview(_currentFilePath);
        CurrentImagePreview = imagePreview;

        if (imagePreview is not null)
        {
            CurrentLanguageExtension = null;
            EditorTextBox.SyntaxHighlighting = null;
            ConfigureRainbowBrackets(null);
            return;
        }

        RefreshCurrentFileSyntaxHighlighting();
    }

    private void ConfigureRainbowBrackets(LoadedExtension? ext)
    {
        _rainbowBracketColorizer.UpdateSyntax(ext);
        EditorTextBox?.TextArea.TextView.InvalidateLayer(KnownLayer.Text);
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

    public bool IsWhatsNewPageVisible
    {
        get => _isWhatsNewPageVisible;
        set
        {
            if (_isWhatsNewPageVisible == value) return;
            _isWhatsNewPageVisible = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsEditorPageVisible));
        }
    }

    public bool IsHomePageVisible
    {
        get => _isHomePageVisible;
        set
        {
            if (_isHomePageVisible == value) return;
            _isHomePageVisible = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsEmptyStateVisible));
            OnPropertyChanged(nameof(IsDocumentViewVisible));
            OnPropertyChanged(nameof(FileSummaryText));
            OnPropertyChanged(nameof(FilePathText));
            OnPropertyChanged(nameof(LanguageDisplayText));
            OnPropertyChanged(nameof(CanShowSaveActions));
        }
    }

    public bool IsWhatsNewExpanded
    {
        get => _isWhatsNewExpanded;
        set
        {
            if (_isWhatsNewExpanded == value) return;
            _isWhatsNewExpanded = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(WhatsNewToggleText));
            OnPropertyChanged(nameof(WhatsNewToggleGlyph));
        }
    }

    // Base width 240 + extra pixels per character beyond 2 in the longest icon label.
    // 2-char labels ("Py","JS") → 240px; 3-char ("C++","F#") → 252px; 4-char ("Java") → 264px.
    public double ExplorerPanelWidth
    {
        get
        {
            var maxLen = FileTreeItems
                .Select(i => (i.Icon ?? string.Empty).Length)
                .DefaultIfEmpty(2)
                .Max();
            return 240 + Math.Max(0, maxLen - 2) * 12;
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

    public bool IsEditorPageVisible => !IsSettingsPageVisible && !IsExtensionsPageVisible && !IsWhatsNewPageVisible;

    public bool HasDocumentOpen => _currentFilePath is not null || _hasUntitledDocument;

    public bool HasOpenEditors => OpenTabs.Count > 0;

    public bool IsDocumentViewVisible => HasDocumentOpen && !IsHomePageVisible;

    public bool HasImagePreview => CurrentImagePreview is not null;

    public bool IsImagePreviewVisible => IsDocumentViewVisible && HasImagePreview;

    public bool IsTextEditorVisible => IsDocumentViewVisible && !HasImagePreview;

    public bool CanShowFindInFile => IsTextEditorVisible;

    public bool IsFindPanelActive => IsFindPanelVisible && CanShowFindInFile;

    public bool IsEditorTabsVisible => OpenTabs.Count >= 1 && !IsHomePageVisible;

    public bool CanShowSaveActions => IsTextEditorVisible;

    public string WhatsNewToggleText => IsWhatsNewExpanded ? "Hide release notes" : "Show release notes";

    public string WhatsNewToggleGlyph => IsWhatsNewExpanded ? "▾" : "▸";

    public bool HasFileOpen => _currentFilePath is not null;

    public bool IsFolderOpen => _currentFolderPath is not null;

    public bool IsEmptyStateVisible => IsHomePageVisible || !HasDocumentOpen;

    public bool HasRecentFiles => RecentFiles.Count > 0;

    // Sub-entries from multi-theme arrays are hidden from the Installed list
    public IEnumerable<LoadedExtension> VisibleLoadedExtensions =>
        LoadedExtensions.Where(e => !e.IsThemeSubEntry);

    public bool IsNoExtensionsVisible => !VisibleLoadedExtensions.Any();

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

    public string LatestReleaseStatusText
    {
        get => _latestReleaseStatusText;
        private set
        {
            if (_latestReleaseStatusText == value) return;
            _latestReleaseStatusText = value;
            OnPropertyChanged();
        }
    }

    public ReleaseInfo? LatestRelease
    {
        get => _latestRelease;
        private set
        {
            if (_latestRelease == value) return;
            _latestRelease = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(HasLatestRelease));
            OnPropertyChanged(nameof(LatestReleaseDisplayName));
            OnPropertyChanged(nameof(LatestReleaseTag));
            OnPropertyChanged(nameof(LatestReleaseNotes));
            OnPropertyChanged(nameof(LatestReleasePreview));
            OnPropertyChanged(nameof(LatestReleaseUrl));
            OnPropertyChanged(nameof(LatestReleaseLinks));
            OnPropertyChanged(nameof(HasLatestReleaseLinks));
            OnPropertyChanged(nameof(IsNewerVersionAvailable));
            OnPropertyChanged(nameof(IsAppUpdateAvailable));
        }
    }

    public bool HasLatestRelease => LatestRelease is not null;

    public string CurrentAppVersionDisplay => CurrentAppVersion;

    private bool _updateBannerDismissed;
    private bool _extensionUpdateBannerDismissed;

    // Returns true if the current build is a -DEV build.
    // DEV builds suppress all app update UI (but extension updating still works).
    private static bool IsDevBuild =>
        CurrentAppVersion.Contains("-DEV", StringComparison.OrdinalIgnoreCase);

    // Priority: stable (no suffix) = 2, -BETA = 1, -DEV = 0
    private static int VersionPriority(string tag)
    {
        if (tag.Contains("-DEV",  StringComparison.OrdinalIgnoreCase)) return 0;
        if (tag.Contains("-BETA", StringComparison.OrdinalIgnoreCase)) return 1;
        return 2;
    }

    private static string StripPreRelease(string tag)
    {
        var t = tag.TrimStart('v');
        var dash = t.IndexOf('-');
        return dash >= 0 ? t[..dash] : t;
    }

    // Core version check — dismissal has no effect here.
    // -DEV builds always return false (updates suppressed entirely).
    // Priority: stable > beta > dev. A stable latest beats a beta current
    // of the same numeric version; a beta latest beats a dev current, etc.
    public bool IsNewerVersionAvailable
    {
        get
        {
            if (IsDevBuild) return false;
            if (!HasLatestRelease || string.IsNullOrWhiteSpace(LatestReleaseTag)) return false;

            if (!Version.TryParse(StripPreRelease(CurrentAppVersion), out var current)) return false;
            if (!Version.TryParse(StripPreRelease(LatestReleaseTag),  out var latest))  return false;

            if (latest != current) return latest > current;

            // Same numeric version — compare by suffix priority
            return VersionPriority(LatestReleaseTag) > VersionPriority(CurrentAppVersion);
        }
    }

    // Banner visibility — collapses when dismissed, reappears if the app restarts.
    public bool IsAppUpdateAvailable => IsNewerVersionAvailable && !_updateBannerDismissed;
    public int AvailableExtensionUpdatesCount => MarketplaceExtensions.Count(e => e.IsUpdateAvailable);
    public bool IsExtensionUpdateBannerVisible => AvailableExtensionUpdatesCount > 0 && !_extensionUpdateBannerDismissed;
    public string ExtensionUpdatesBannerText =>
        $"{AvailableExtensionUpdatesCount} extension{(AvailableExtensionUpdatesCount == 1 ? string.Empty : "s")} {(AvailableExtensionUpdatesCount == 1 ? "has" : "have")} updates available";

    public string LatestReleaseDisplayName =>
        !string.IsNullOrWhiteSpace(LatestRelease?.Name)
            ? LatestRelease.Name
            : !string.IsNullOrWhiteSpace(LatestRelease?.Tag)
                ? LatestRelease.Tag
                : "Latest Release";

    public string LatestReleaseTag => LatestRelease?.Tag ?? string.Empty;

    public string LatestReleaseNotes => string.IsNullOrWhiteSpace(LatestRelease?.Notes)
        ? "No release notes available."
        : ConvertMarkdownToDisplayText(LatestRelease.Notes);

    public string LatestReleasePreview
    {
        get
        {
            var notes = LatestReleaseNotes.Replace("\r\n", "\n").Replace('\r', '\n');
            var preview = notes.Length > 220 ? notes[..220].TrimEnd() + "..." : notes;
            return preview;
        }
    }

    public string LatestReleaseUrl => LatestRelease?.Url ?? ReleasesPageUrl;

    public IReadOnlyList<ReleaseLinkItem> LatestReleaseLinks =>
        ExtractReleaseLinks(LatestRelease?.Notes ?? string.Empty);

    public bool HasLatestReleaseLinks => LatestReleaseLinks.Count > 0;

    private static string ConvertMarkdownToDisplayText(string markdown)
    {
        if (string.IsNullOrWhiteSpace(markdown))
            return string.Empty;

        var text = markdown.Replace("\r\n", "\n").Replace('\r', '\n');

        text = MdCodeFenceRegex.Replace(text, string.Empty);
        text = text.Replace("```", string.Empty);
        text = MdInlineCodeRegex.Replace(text, "$1");
        text = MdImageRegex.Replace(text, "$1");
        text = MdLinkRegex.Replace(text, "$1");
        text = MdHeadingRegex.Replace(text, string.Empty);
        text = MdBlockquoteRegex.Replace(text, string.Empty);
        text = MdHrRegex.Replace(text, string.Empty);
        text = MdBulletRegex.Replace(text, "• ");
        text = MdOrderedListRegex.Replace(text, "$1. ");
        text = MdBoldRegex.Replace(text, "$1");
        text = MdItalicRegex.Replace(text, "$1");
        text = MdBoldUnderscoreRegex.Replace(text, "$1");
        text = MdItalicUnderscoreRegex.Replace(text, "$1");
        text = MdStrikethroughRegex.Replace(text, "$1");
        text = MdTableLeadingPipeRegex.Replace(text, string.Empty);
        text = MdTableTrailingPipeRegex.Replace(text, string.Empty);
        text = MdTablePipeRegex.Replace(text, " | ");
        text = MdExcessNewlinesRegex.Replace(text, "\n\n");

        return text.Trim();
    }

    // Pre-compiled regex patterns used by ConvertMarkdownToDisplayText — compiled
    // once at class load time so repeated calls don't recompile on every access.
    private static readonly Regex MdCodeFenceRegex         = new(@"```(?:[\w#+.-]+)?\n?",         RegexOptions.Compiled);
    private static readonly Regex MdInlineCodeRegex        = new(@"`([^`]+)`",                    RegexOptions.Compiled);
    private static readonly Regex MdImageRegex             = new(@"!\[([^\]]*)\]\([^)]+\)",        RegexOptions.Compiled);
    private static readonly Regex MdLinkRegex              = new(@"\[(.*?)\]\((.*?)\)",            RegexOptions.Compiled);
    private static readonly Regex MdHeadingRegex           = new(@"(?m)^\s{0,3}#{1,6}\s*",        RegexOptions.Compiled);
    private static readonly Regex MdBlockquoteRegex        = new(@"(?m)^\s{0,3}>\s?",             RegexOptions.Compiled);
    private static readonly Regex MdHrRegex                = new(@"(?m)^\s*[-*_]{3,}\s*$",        RegexOptions.Compiled);
    private static readonly Regex MdBulletRegex            = new(@"(?m)^\s*[-*+]\s+",             RegexOptions.Compiled);
    private static readonly Regex MdOrderedListRegex       = new(@"(?m)^\s*(\d+)\.\s+",           RegexOptions.Compiled);
    private static readonly Regex MdBoldRegex              = new(@"(?<!\*)\*\*(?!\*)(.*?)\*\*(?<!\*)", RegexOptions.Compiled);
    private static readonly Regex MdItalicRegex            = new(@"(?<!\*)\*(?!\*)(.*?)\*(?<!\*)", RegexOptions.Compiled);
    private static readonly Regex MdBoldUnderscoreRegex    = new(@"__(.*?)__",                    RegexOptions.Compiled);
    private static readonly Regex MdItalicUnderscoreRegex  = new(@"(?<!_)_(?!_)(.*?)(?<!_)_(?!_)", RegexOptions.Compiled);
    private static readonly Regex MdStrikethroughRegex     = new(@"~~(.*?)~~",                    RegexOptions.Compiled);
    private static readonly Regex MdTableLeadingPipeRegex  = new(@"(?m)^\s*\|",                   RegexOptions.Compiled);
    private static readonly Regex MdTableTrailingPipeRegex = new(@"(?m)\|\s*$",                   RegexOptions.Compiled);
    private static readonly Regex MdTablePipeRegex         = new(@"\|",                           RegexOptions.Compiled);
    private static readonly Regex MdExcessNewlinesRegex    = new(@"\n{3,}",                       RegexOptions.Compiled);

    private static IReadOnlyList<ReleaseLinkItem> ExtractReleaseLinks(string markdown)
    {
        if (string.IsNullOrWhiteSpace(markdown))
            return [];

        var links = new List<ReleaseLinkItem>();
        var seenUrls = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (Match match in Regex.Matches(markdown, @"\[(.*?)\]\((https?://[^\s)]+)\)"))
        {
            var label = match.Groups[1].Value.Trim();
            var url = match.Groups[2].Value.Trim();
            if (string.IsNullOrWhiteSpace(url) || !seenUrls.Add(url))
                continue;

            links.Add(new ReleaseLinkItem
            {
                Label = string.IsNullOrWhiteSpace(label) ? url : label,
                Url = url
            });
        }

        foreach (Match match in Regex.Matches(markdown, @"https?://[^\s)]+"))
        {
            var url = match.Value.Trim().TrimEnd('.', ',', ';');
            if (string.IsNullOrWhiteSpace(url) || !seenUrls.Add(url))
                continue;

            links.Add(new ReleaseLinkItem
            {
                Label = url,
                Url = url
            });
        }

        return links;
    }

    public bool IsRefreshingLatestRelease => _isRefreshingLatestRelease;

    public string RefreshLatestReleaseButtonText => IsRefreshingLatestRelease ? "Refreshing..." : "Refresh";

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

    public bool IsStatusBarFilePathVisible
    {
        get => _isStatusBarFilePathVisible;
        set
        {
            if (_isStatusBarFilePathVisible == value) return;
            _isStatusBarFilePathVisible = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(FilePathText));
            OnPropertyChanged(nameof(LanguageDisplayText));
            OnPropertyChanged(nameof(StatusBarFilePathVisibilityText));
            SaveSettings();
        }
    }

    public bool IsWordWrapEnabled
    {
        get => _isWordWrapEnabled;
        set
        {
            if (_isWordWrapEnabled == value) return;
            _isWordWrapEnabled = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(EditorBehaviorStatusText));
            ApplyEditorSettings();
            SaveSettings();
        }
    }

    public int TabSize
    {
        get => _tabSize;
        set
        {
            var normalizedValue = NormalizeTabSize(value);
            if (_tabSize == normalizedValue) return;
            _tabSize = normalizedValue;
            OnPropertyChanged();
            OnPropertyChanged(nameof(EditorBehaviorStatusText));
            ApplyEditorSettings();
            SaveSettings();
        }
    }

    public int TabSizeIndex
    {
        get => TabSize switch
        {
            2 => 0,
            8 => 2,
            _ => 1
        };
        set => TabSize = value switch
        {
            0 => 2,
            2 => 8,
            _ => 4
        };
    }

    public int EditorFontSize
    {
        get => _editorFontSize;
        set
        {
            var clamped = Math.Clamp(value, 8, 32);
            if (_editorFontSize == clamped) return;
            _editorFontSize = clamped;
            OnPropertyChanged();
            ApplyEditorSettings();
            SaveSettings();
        }
    }

    public string ExtensionSearchText
    {
        get => _extensionSearchText;
        set
        {
            if (_extensionSearchText == value) return;
            _extensionSearchText = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(FilteredInstalledExtensions));
            OnPropertyChanged(nameof(FilteredMarketplaceExtensions));
        }
    }

    public IEnumerable<LoadedExtension> FilteredInstalledExtensions =>
        string.IsNullOrWhiteSpace(_extensionSearchText)
            ? VisibleLoadedExtensions
            : VisibleLoadedExtensions.Where(e =>
                e.Name.Contains(_extensionSearchText, StringComparison.OrdinalIgnoreCase) ||
                e.Description.Contains(_extensionSearchText, StringComparison.OrdinalIgnoreCase));

    public IEnumerable<MarketplaceExtension> FilteredMarketplaceExtensions =>
        string.IsNullOrWhiteSpace(_extensionSearchText)
            ? MarketplaceExtensions
            : MarketplaceExtensions.Where(e =>
                e.Name.Contains(_extensionSearchText, StringComparison.OrdinalIgnoreCase) ||
                e.Description.Contains(_extensionSearchText, StringComparison.OrdinalIgnoreCase) ||
                e.Author.Contains(_extensionSearchText, StringComparison.OrdinalIgnoreCase));

    public bool IsFindPanelVisible
    {
        get => _isFindPanelVisible;
        set
        {
            if (_isFindPanelVisible == value) return;
            _isFindPanelVisible = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsFindPanelActive));
        }
    }

    public string FindText
    {
        get => _findText;
        set
        {
            if (_findText == value) return;
            _findText = value;
            OnPropertyChanged();
        }
    }

    public bool IsConfirmBeforeClosingUnsavedTabsEnabled
    {
        get => _isConfirmBeforeClosingUnsavedTabsEnabled;
        set
        {
            if (_isConfirmBeforeClosingUnsavedTabsEnabled == value) return;
            _isConfirmBeforeClosingUnsavedTabsEnabled = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(TabRestoreStatusText));
            SaveSettings();
        }
    }

    public bool IsRestoreOpenTabsOnLaunchEnabled
    {
        get => _isRestoreOpenTabsOnLaunchEnabled;
        set
        {
            if (_isRestoreOpenTabsOnLaunchEnabled == value) return;
            _isRestoreOpenTabsOnLaunchEnabled = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(TabRestoreStatusText));
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
            OnPropertyChanged(nameof(ThemeStatusText));
            OnPropertyChanged(nameof(IsDarkThemeActive));
            OnPropertyChanged(nameof(IsLightThemeActive));
            foreach (var ext in ThemeExtensions)
                ext.IsActiveTheme = string.Equals(ext.ThemeCardThemeId, value, StringComparison.OrdinalIgnoreCase);
        }
    }

    // Displays the file name and Unsaved/autosave status in the top bar
    public string FileSummaryText => IsHomePageVisible
        ? "Home"
        : HasDocumentOpen
        ? $"{GetDocumentDisplayName()}{GetDocumentStatusSuffix()}"
        : "Editor";

    public string FilePathText => IsHomePageVisible
        ? "Welcome to Kodo!"
        : HasFileOpen
        ? (IsStatusBarFilePathVisible ? _currentFilePath! : GetDocumentDisplayName())
        : HasDocumentOpen ? "Unsaved file"
        : IsFolderOpen ? (IsStatusBarFilePathVisible
            ? $"📂 {_currentFolderPath}"
            : $"📂 {Path.GetFileName(_currentFolderPath!.TrimEnd(Path.DirectorySeparatorChar))}")
        : "No file open";

    public string ExplorerHeaderText => IsFolderOpen
        ? Path.GetFileName(_currentFolderPath!.TrimEnd(Path.DirectorySeparatorChar)).ToUpperInvariant()
        : "EXPLORER";

    public string ThemeStatusText => $"Current theme: {CurrentThemeName}";

    private static string TimeOfDay()
    {
        int hour = DateTime.Now.Hour;
        if (hour < 6)  return "night";
        if (hour < 12) return "morning";
        if (hour < 17) return "afternoon";
        if (hour < 22) return "evening";
        return "night";
    }

    private static readonly string[] _welcomeMessages = BuildWelcomeMessages();

    private static string[] BuildWelcomeMessages()
    {
        var tod = TimeOfDay();
        var messages = new List<string>
        {
            "Welcome back!",
            "Great to see you!",
            "Ready to code?",
            "Let's build something!",
            "What are we building today?",
            "Back at it again!",
            "Let's get to work!",
            "Hey there!",
            $"Good {tod}!",
            $"Good {tod}, ready to build?",
            $"Good {tod}, let's get to it!",
            $"It's a great {tod} to code!",
        };

        if (tod == "morning") messages.Add("Hey there, early bird!");
        if (tod == "night")   messages.Add("Hey there, night owl!");

        return messages.ToArray();
    }

    public static string WelcomeMessage { get; } =
        _welcomeMessages[Random.Shared.Next(_welcomeMessages.Length)];

    public bool IsDarkThemeActive  => string.Equals(CurrentThemeName, "Dark",  StringComparison.OrdinalIgnoreCase);
    public bool IsLightThemeActive => string.Equals(CurrentThemeName, "Light", StringComparison.OrdinalIgnoreCase);

    public string LanguageDisplayText
    {
        get
        {
            if (HasImagePreview) return "Image Preview";
            if (!HasDocumentOpen) return string.Empty;
            if (!string.IsNullOrWhiteSpace(_currentFilePath))
            {
                var ext = Path.GetExtension(_currentFilePath);
                if (!string.IsNullOrWhiteSpace(ext))
                    return $"{ext.ToLowerInvariant()} file";
                var name = Path.GetFileName(_currentFilePath);
                return string.IsNullOrWhiteSpace(name) ? "Plain Text" : $"{name} file";
            }
            return "Plain Text";
        }
    }

    public string DiscordRichPresenceStatusText => !IsDiscordRichPresenceEnabled
        ? "Discord Rich Presence is turned off."
        : "Discord Rich Presence is on when the Discord desktop app is running.";

    public string AutoSaveStatusText =>
        !IsAutoSaveEnabled
            ? "Autosave is turned off."
            : HasFileOpen
                ? "Changes are saved automatically a couple seconds after you stop typing."
                : "Autosave will start working after the file has been saved once.";

    public string StatusBarFilePathVisibilityText => IsStatusBarFilePathVisible
        ? "The status bar shows the full path for the current file or folder."
        : "The status bar keeps a shorter label instead of the full file or folder path.";

    public string EditorBehaviorStatusText => IsWordWrapEnabled
        ? $"Word wrap is on, and indentation guides are spaced every {TabSize} columns."
        : $"Word wrap is off, and indentation guides are spaced every {TabSize} columns.";

    public string TabRestoreStatusText => IsRestoreOpenTabsOnLaunchEnabled
        ? "File-backed tabs reopen on launch, and Unsaved tabs ask for confirmation before closing."
        : IsConfirmBeforeClosingUnsavedTabsEnabled
            ? "Unsaved tabs ask for confirmation before closing."
            : "Tabs close immediately, and launch starts with a fresh editor session.";

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

    private void OpenTabs_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.Action == NotifyCollectionChangedAction.Add)
            IsFileExplorerVisible = true;
        OnPropertyChanged(nameof(HasOpenEditors));
        OnPropertyChanged(nameof(IsEditorTabsVisible));
        SaveSettings();
    }

    // ── State management ─────────────────────────────────────────────────────

    private void RefreshState()
    {
        var document = EditorTextBox?.Document;
        if (HasDocumentOpen && !IsHomePageVisible)
        {
            if (CurrentImagePreview is not null)
            {
                EditorStatsText = $"{CurrentImagePreview.PixelSize.Width} x {CurrentImagePreview.PixelSize.Height}px";
            }
            else
            {
                var lines      = document?.LineCount ?? 1;
                var characters = document?.TextLength ?? 0;
                var caret      = EditorTextBox?.TextArea?.Caret;
                var ln         = caret?.Line ?? 1;
                var col        = caret?.Column ?? 1;
                EditorStatsText = $"Ln {ln}, Col {col}  |  {lines} lines  |  {characters} characters";
            }
        }
        else
        {
            EditorStatsText = string.Empty;
        }

        Title = HasDocumentOpen ? $"{GetDocumentDisplayName()} - Kodo" : "Kodo";
        OnPropertyChanged(nameof(HasDocumentOpen));
        OnPropertyChanged(nameof(IsDocumentViewVisible));
        OnPropertyChanged(nameof(HasImagePreview));
        OnPropertyChanged(nameof(IsImagePreviewVisible));
        OnPropertyChanged(nameof(IsTextEditorVisible));
        OnPropertyChanged(nameof(CanShowFindInFile));
        OnPropertyChanged(nameof(IsFindPanelActive));
        OnPropertyChanged(nameof(CanShowSaveActions));
        OnPropertyChanged(nameof(HasFileOpen));
        OnPropertyChanged(nameof(IsFolderOpen));
        OnPropertyChanged(nameof(IsEmptyStateVisible));
        OnPropertyChanged(nameof(HasRecentFiles));
        OnPropertyChanged(nameof(FileSummaryText));
        OnPropertyChanged(nameof(FilePathText));
        OnPropertyChanged(nameof(ExplorerHeaderText));
        OnPropertyChanged(nameof(DiscordRichPresenceStatusText));
        OnPropertyChanged(nameof(AutoSaveStatusText));
        OnPropertyChanged(nameof(LanguageDisplayText));
        UpdateDiscordPresence();
    }

    private void QueueRefreshState()
    {
        _editorStateRefreshTimer.Stop();
        _editorStateRefreshTimer.Start();
    }

    private void EditorStateRefreshTimer_OnTick(object? sender, EventArgs e)
    {
        _editorStateRefreshTimer.Stop();
        RefreshState();
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

            if (_isSaving)
                return AutoSaveSavingMessage;

            if (_isDirty || _autoSaveTimer.IsEnabled)
                return "Unsaved";
        }

        return _isDirty ? "Unsaved" : null;
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
        _discordReconnectTimer.Stop();
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
            ResetDiscordPresenceForReconnect();
            _discordReconnectTimer.Start();
        }
    }

    private void UpdateDiscordPresence()
    {
        if (_discordRpcClient is null || !IsDiscordRichPresenceEnabled) return;

        try
        {
            var details = GetDiscordPresenceDetails();
            var state = GetDiscordPresenceState();
            if (details == _lastDiscordPresenceDetails && state == _lastDiscordPresenceState)
                return;

            _discordRpcClient.SetPresence(new DiscordRichPresenceModel
            {
                Details    = details,
                State      = state,
                Assets     = new DiscordAssetsModel
                {
                    LargeImageKey  = DefaultDiscordLargeImageKey,
                    LargeImageText = DefaultDiscordLargeImageText
                },
                Timestamps = new DiscordRPC.Timestamps(_sessionStart)
            });
            _lastDiscordPresenceDetails = details;
            _lastDiscordPresenceState = state;
        }
        catch
        {
            ResetDiscordPresenceForReconnect();
            _discordReconnectTimer.Start();
        }
    }

    private string GetDiscordPresenceDetails()
    {
        if (HasDocumentOpen) return $"Editing {GetDocumentDisplayName()}";
        return IsFolderOpen ? "Browsing project files" : "Idle in Kodo";
    }

    private string GetDiscordPresenceState()
    {
        if (HasFileOpen)          return GetDiscordWorkspaceLabel();
        if (_hasUntitledDocument) return GetDiscordWorkspaceLabel("Editing an Unsaved file");
        if (IsFolderOpen)         return GetDiscordWorkspaceLabel();
        return "Waiting for a file";
    }

    private string GetDiscordWorkspaceLabel(string fallback = "Working in editor")
    {
        if (!IsFolderOpen) return fallback;
        var folderName = Path.GetFileName(_currentFolderPath!.TrimEnd(Path.DirectorySeparatorChar));
        return string.IsNullOrWhiteSpace(folderName) ? fallback : $"Workspace: {folderName}";
    }

    private void DisposeDiscordPresence(bool clearPresence = true)
    {
        _discordReconnectTimer.Stop();
        if (_discordRpcClient is null) return;
        try
        {
            if (clearPresence)
                _discordRpcClient.ClearPresence();
            _discordRpcClient.Dispose();
        }
        catch { /* Ignore cleanup failures. */ }
        finally
        {
            _discordRpcClient = null;
            _lastDiscordPresenceDetails = string.Empty;
            _lastDiscordPresenceState = string.Empty;
        }
    }

    private void ResetDiscordPresenceForReconnect()
    {
        if (_discordRpcClient is null)
            return;

        try
        {
            _discordRpcClient.Dispose();
        }
        catch { /* Ignore reconnect cleanup failures. */ }
        finally
        {
            _discordRpcClient = null;
        }
    }

    private void DiscordReconnectTimer_OnTick(object? sender, EventArgs e)
    {
        _discordReconnectTimer.Stop();
        UpdateDiscordRichPresenceLifecycle();
    }

    // ── Settings persistence ─────────────────────────────────────────────────

    private string SettingsFilePath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Kodo", SettingsFileName);

    private AppSettings LoadSettings()
    {
        try
        {
            if (!File.Exists(SettingsFilePath)) return new AppSettings();

            var json = File.ReadAllText(SettingsFilePath);

            // An empty or whitespace-only file means a previous write was interrupted.
            // Treat it like a missing file rather than overwriting with defaults.
            if (string.IsNullOrWhiteSpace(json)) return new AppSettings();

            var settings = JsonSerializer.Deserialize<AppSettings>(json);
            if (settings is null) return new AppSettings();

            settings.ThemeName = string.IsNullOrWhiteSpace(settings.ThemeName) ? "Dark" : settings.ThemeName;
            settings.RecentFiles = settings.RecentFiles?
                .Where(e => !string.IsNullOrWhiteSpace(e.Path))
                .GroupBy(e => e.Path, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList() ?? [];
            settings.OpenTabPaths = settings.OpenTabPaths?
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? [];
            settings.TabSize = NormalizeTabSize(settings.TabSize);
            return settings;
        }
        catch { return new AppSettings(); }
    }

    private void SaveSettings()
    {
        // Snapshot all UI-thread-owned state here, before the background task,
        // so we don't access ObservableCollections or bound properties from a background thread.
        var snapshot = new AppSettings
        {
            ThemeName                              = CurrentThemeName,
            AutoSaveEnabled                        = IsAutoSaveEnabled,
            DiscordRichPresenceEnabled             = IsDiscordRichPresenceEnabled,
            StatusBarFilePathVisible               = IsStatusBarFilePathVisible,
            WordWrapEnabled                        = IsWordWrapEnabled,
            TabSize                                = TabSize,
            EditorFontSize                         = EditorFontSize,
            ConfirmBeforeClosingUnsavedTabsEnabled  = IsConfirmBeforeClosingUnsavedTabsEnabled,
            RestoreOpenTabsOnLaunchEnabled          = IsRestoreOpenTabsOnLaunchEnabled,
            OpenTabPaths = OpenTabs
                .Where(tab => !tab.IsUntitled && !string.IsNullOrWhiteSpace(tab.Path))
                .Select(tab => tab.Path)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList(),
            ActiveTabPath = ActiveEditorTab is { IsUntitled: false } activeTab
                ? activeTab.Path
                : null,
            RecentFiles = RecentFiles
                .Select(f => new RecentFileEntry { Path = f.Path, IsFolder = f.IsFolder, LastOpened = f.LastOpened })
                .ToList()
        };

        Task.Run(() =>
        {
            try
            {
                var dir = Path.GetDirectoryName(SettingsFilePath);
                if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);

                // Write to a temp file first, then atomically replace the real file.
                // This prevents settings from being wiped if the app is killed or
                // crashes mid-write, which leaves File.WriteAllText with a zero-byte file.
                var tempPath = SettingsFilePath + ".tmp";
                File.WriteAllText(tempPath, JsonSerializer.Serialize(snapshot));
                File.Move(tempPath, SettingsFilePath, overwrite: true);
            }
            catch { /* Ignore settings persistence failures. */ }
        });
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

    private void SaveCurrentEditorStateIntoTab()
    {
        if (ActiveEditorTab is null || EditorTextBox?.Document is null)
            return;

        if (HasImagePreview)
        {
            ActiveEditorTab.IsDirty = false;
            if (!ActiveEditorTab.IsUntitled && !string.IsNullOrWhiteSpace(_currentFilePath))
                ActiveEditorTab.Path = _currentFilePath;
            return;
        }

        ActiveEditorTab.Content = EditorTextBox.Document.Text;
        ActiveEditorTab.IsDirty = _isDirty;
        if (!ActiveEditorTab.IsUntitled && !string.IsNullOrWhiteSpace(_currentFilePath))
            ActiveEditorTab.Path = _currentFilePath;
    }

    private EditorTab CreateUntitledTab()
    {
        var displayName = $"untitled-{_nextUntitledTabNumber++}.txt";
        return new EditorTab(displayName, displayName, string.Empty, isUntitled: true);
    }

    private void ActivateTab(EditorTab tab, bool focusEditor = true, bool preserveCurrentState = true)
    {
        if (ReferenceEquals(ActiveEditorTab, tab))
        {
            // Force page state even if NavigateTo bails early due to no change
            _isHomePageVisible = false;
            NavigateTo(Page.Editor);
            RefreshState();
            if (focusEditor)
                FocusEditor();
            return;
        }

        if (preserveCurrentState)
            SaveCurrentEditorStateIntoTab();
        ActiveEditorTab = tab;
        _currentFilePath = tab.IsUntitled ? null : tab.Path;
        _hasUntitledDocument = tab.IsUntitled;
        _isDirty = tab.IsDirty;
        _autoSaveTimer.Stop();
        ClearAutoSaveStatus();
        SetEditorContent(IsImagePreviewFile(_currentFilePath) ? string.Empty : tab.Content);
        UpdateCurrentDocumentPresentation();

        // Directly set the backing field before NavigateTo so the bail-early
        // check doesn't short-circuit when we're already on the editor page.
        _isHomePageVisible = false;
        NavigateTo(Page.Editor);
        RefreshState();

        if (focusEditor)
            FocusEditor();
    }

    private void CloseTab(EditorTab tab)
    {
        var closingActiveTab = ReferenceEquals(tab, ActiveEditorTab);
        var index = OpenTabs.IndexOf(tab);
        if (index < 0)
            return;

        OpenTabs.RemoveAt(index);

        if (!closingActiveTab)
        {
            RefreshState();
            return;
        }

        if (OpenTabs.Count > 0)
        {
            var nextIndex = Math.Min(index, OpenTabs.Count - 1);
            ActivateTab(OpenTabs[nextIndex], focusEditor: true, preserveCurrentState: false);
            return;
        }

        ActiveEditorTab = null;
        _currentFilePath = null;
        _hasUntitledDocument = false;
        _isDirty = false;
        CurrentLanguageExtension = null;
        CurrentImagePreview = null;
        EditorTextBox.SyntaxHighlighting = null;
        ConfigureRainbowBrackets(null);
        SetEditorContent(string.Empty);
        RefreshState();
    }

    private async Task<bool> RequestCloseTabAsync(EditorTab tab)
    {
        if (tab.IsDirty && IsConfirmBeforeClosingUnsavedTabsEnabled)
        {
            var originalActiveTab = ActiveEditorTab;
            var action = await ShowUnsavedTabDialogAsync(tab);
            switch (action)
            {
                case UnsavedTabAction.Cancel:
                    return false;
                case UnsavedTabAction.Save:
                    if (!ReferenceEquals(tab, ActiveEditorTab))
                    {
                        ActivateTab(tab, focusEditor: false);
                    }

                    await SaveAsync();

                    if (tab.IsDirty)
                    {
                        if (originalActiveTab is not null &&
                            !ReferenceEquals(originalActiveTab, tab) &&
                            OpenTabs.Contains(originalActiveTab))
                        {
                            ActivateTab(originalActiveTab, focusEditor: false);
                        }

                        return false;
                    }

                    if (originalActiveTab is not null &&
                        !ReferenceEquals(originalActiveTab, tab) &&
                        OpenTabs.Contains(originalActiveTab))
                    {
                        ActivateTab(originalActiveTab, focusEditor: false);
                    }
                    break;
            }
        }

        CloseTab(tab);
        return true;
    }

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
        EnsureCurrentDocumentHasTab();

        var existingTab = OpenTabs.FirstOrDefault(tab =>
            !tab.IsUntitled &&
            string.Equals(tab.Path, path, StringComparison.OrdinalIgnoreCase));
        if (existingTab is not null)
        {
            ActivateTab(existingTab);
            return;
        }

        string content;
        try
        {
            content = IsImagePreviewFile(path) ? string.Empty : await File.ReadAllTextAsync(path);
        }
        catch (Exception ex)
        {
            await ShowWarningDialogAsync("Open file", ex);
            return;
        }

        var tab = new EditorTab(path, Path.GetFileName(path), content);
        OpenTabs.Add(tab);
        ActivateTab(tab);
    }

    private void EnsureCurrentDocumentHasTab()
    {
        if (ActiveEditorTab is not null || !HasDocumentOpen)
        {
            return;
        }

        var displayName = _currentFilePath is not null
            ? Path.GetFileName(_currentFilePath)
            : $"untitled-{_nextUntitledTabNumber++}.txt";
        var path = _currentFilePath ?? displayName;
        var content = IsImagePreviewFile(_currentFilePath) ? string.Empty : EditorTextBox?.Document?.Text ?? string.Empty;
        var recoveredTab = new EditorTab(path, displayName, content, isUntitled: _currentFilePath is null)
        {
            IsDirty = _isDirty
        };

        OpenTabs.Add(recoveredTab);
        ActivateTab(recoveredTab, focusEditor: false);
    }

    private async void MainWindow_OnOpened(object? sender, EventArgs e)
    {
        Opened -= MainWindow_OnOpened;

        if (IsRestoreOpenTabsOnLaunchEnabled && _startupOpenTabPaths.Count > 0)
        {
            foreach (var path in _startupOpenTabPaths)
            {
                await OpenFileFromPathAsync(path);
            }

            if (!string.IsNullOrWhiteSpace(_startupActiveTabPath))
            {
                var activeTab = OpenTabs.FirstOrDefault(tab =>
                    !tab.IsUntitled &&
                    string.Equals(tab.Path, _startupActiveTabPath, StringComparison.OrdinalIgnoreCase));
                if (activeTab is not null)
                {
                    ActivateTab(activeTab);
                }
            }
        }

        // Open the file passed on the command line (e.g. via "Open with" or double-click)
        if (!string.IsNullOrWhiteSpace(_startupFilePath))
        {
            await OpenFileFromPathAsync(_startupFilePath);
        }
    }

    private void NewFile()
    {
        // Navigate away from home BEFORE adding the tab, so the CollectionChanged
        // notification evaluates IsEditorTabsVisible with IsHomePageVisible already false.
        NavigateTo(Page.Editor);

        var tab = CreateUntitledTab();
        OpenTabs.Add(tab);
        ActivateTab(tab);
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
        AddRecentFolder(path);
        await PopulateFileTreeAsync(path);
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

    private async Task<bool> SaveAsync(bool allowPromptForPath, bool forcePromptForPath)
    {
        _autoSaveTimer.Stop();
        if (_isSaving) return false;
        if (HasImagePreview) return false;

        var shouldPromptForPath = forcePromptForPath || _currentFilePath is null;
        if (shouldPromptForPath)
        {
            if (!allowPromptForPath) return false;

            var suggestedFileName = ActiveEditorTab?.DisplayName;
            if (string.IsNullOrWhiteSpace(suggestedFileName))
                suggestedFileName = HasFileOpen ? Path.GetFileName(_currentFilePath) : "untitled.txt";

            var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = forcePromptForPath ? "Save File As" : "Save File",
                SuggestedFileName = suggestedFileName
            });

            var newPath = file?.TryGetLocalPath();
            if (string.IsNullOrWhiteSpace(newPath)) return false;

            _currentFilePath = newPath;
            _hasUntitledDocument = false;
            if (ActiveEditorTab is not null)
                ActiveEditorTab.Rename(newPath, Path.GetFileName(newPath));
            ClearAutoSaveStatus();
            RefreshCurrentFileSyntaxHighlighting();
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
            await File.WriteAllTextAsync(_currentFilePath!, EditorTextBox.Document.Text);
            _isDirty = false;
            if (ActiveEditorTab is not null)
            {
                ActiveEditorTab.Content = EditorTextBox.Document.Text;
                ActiveEditorTab.IsDirty = false;
                if (ActiveEditorTab.IsUntitled)
                    ActiveEditorTab.IsUntitled = false;
            }
            AddRecentFile(_currentFilePath);
            RefreshCurrentFileSyntaxHighlighting();

            if (IsAutoSaveEnabled && HasFileOpen)
            {
                _autoSaveStatusMessage = AutoSaveSavedMessage;
                _autoSaveStatusTimer.Stop();
                _autoSaveStatusTimer.Start();
            }

            RefreshState();
            return true;
        }
        catch (Exception ex)
        {
            _autoSaveStatusTimer.Stop();
            _autoSaveStatusMessage = BuildAutoSaveFailureMessage(ex);
            OnPropertyChanged(nameof(FileSummaryText));
            OnPropertyChanged(nameof(AutoSaveStatusText));
            await ShowWarningDialogAsync("File save", ex);
            return false;
        }
        finally
        {
            _isSaving = false;
            OnPropertyChanged(nameof(FileSummaryText));
            OnPropertyChanged(nameof(AutoSaveStatusText));
        }
    }

    private async Task SaveAsync(bool allowPromptForPath = true) =>
        await SaveAsync(allowPromptForPath, forcePromptForPath: false);

    private async Task<bool> SaveAsAsync() =>
        await SaveAsync(allowPromptForPath: true, forcePromptForPath: true);

    // ── File tree ────────────────────────────────────────────────────────────

    private async Task PopulateFileTreeAsync(string folderPath)
    {
        var items = await CreateFileTreeItemsAsync(folderPath, depth: 0);
        // Swap the entire collection in one shot — avoids one CollectionChanged
        // notification (and ItemsControl re-render) per item.
        FileTreeItems.Clear();
        foreach (var item in items)
            FileTreeItems.Add(item);
    }

    private async Task AppendDirectoryContentsAsync(string dirPath, int depth, int insertAfterIndex = -1)
    {
        var items = await CreateFileTreeItemsAsync(dirPath, depth);
        if (items.Count == 0) return;

        var pos = insertAfterIndex + 1;
        foreach (var item in items)
        {
            if (insertAfterIndex < 0)
                FileTreeItems.Add(item);
            else
            {
                FileTreeItems.Insert(pos, item);
                pos++;
            }
        }
    }

    private static Task<List<FileTreeItem>> CreateFileTreeItemsAsync(string dirPath, int depth) =>
        Task.Run(() => GetSortedEntries(dirPath)
            .Select(entry => new FileTreeItem
            {
                Name        = Path.GetFileName(entry),
                FullPath    = entry,
                IsDirectory = Directory.Exists(entry),
                Depth       = depth,
            })
            .ToList());

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

    private async Task ToggleDirectoryExpansionAsync(FileTreeItem dirItem)
    {
        var index = FileTreeItems.IndexOf(dirItem);
        if (index < 0) return;

        if (dirItem.IsExpanded)
        {
            dirItem.IsExpanded = false;
            // Collect all descendants first, then remove in reverse index order
            // so each removal doesn't shift the indices of subsequent items.
            var toRemove = FileTreeItems
                .Skip(index + 1)
                .TakeWhile(i => i.Depth > dirItem.Depth)
                .ToList();
            for (var i = toRemove.Count - 1; i >= 0; i--)
                FileTreeItems.Remove(toRemove[i]);
        }
        else
        {
            dirItem.IsExpanded = true;
            await AppendDirectoryContentsAsync(dirItem.FullPath, dirItem.Depth + 1, insertAfterIndex: index);
        }
    }

    // ── Recent files ─────────────────────────────────────────────────────────

    private void LoadRecentFiles(IEnumerable<RecentFileEntry>? recentFiles)
    {
        RecentFiles.Clear();
        foreach (var entry in (recentFiles ?? []).Take(MaxRecentFiles))
        {
            if (entry.IsFolder ? Directory.Exists(entry.Path) : File.Exists(entry.Path))
                RecentFiles.Add(new RecentFileItem(entry.Path, entry.IsFolder, entry.LastOpened));
        }
    }

    private void AddRecentFile(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return;

        var existing = RecentFiles.FirstOrDefault(f => string.Equals(f.Path, path, StringComparison.OrdinalIgnoreCase));
        if (existing is not null) RecentFiles.Remove(existing);

        RecentFiles.Insert(0, new RecentFileItem(path, isFolder: false, DateTime.Now));
        while (RecentFiles.Count > MaxRecentFiles)
            RecentFiles.RemoveAt(RecentFiles.Count - 1);

        SaveSettings();
        OnPropertyChanged(nameof(HasRecentFiles));
    }

    private void AddRecentFolder(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return;

        var existing = RecentFiles.FirstOrDefault(f => string.Equals(f.Path, path, StringComparison.OrdinalIgnoreCase));
        if (existing is not null) RecentFiles.Remove(existing);

        RecentFiles.Insert(0, new RecentFileItem(path, isFolder: true, DateTime.Now));
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
        OnPropertyChanged(nameof(AutoSaveStatusText));
    }

    private static string BuildAutoSaveFailureMessage(Exception ex)
    {
        var message = string.IsNullOrWhiteSpace(ex.Message) ? "Unexpected error." : ex.Message.Trim();
        return $"{AutoSaveFailedMessagePrefix} {message}";
    }

    // ── Editor helpers ───────────────────────────────────────────────────────

    private string GetDocumentDisplayName() =>
        HasFileOpen ? Path.GetFileName(_currentFilePath!) : "untitled.txt";

    // Writes content into the TextEditor document without triggering dirty tracking
    private void SetEditorContent(string content)
    {
        _suppressDirtyTracking = true;
        EditorTextBox.Document.Text = content;
        // Clear the flag via a posted action so it stays true until after
        // AvaloniaEdit's async TextChanged event has fired and been handled.
        Dispatcher.UIThread.Post(
            () => _suppressDirtyTracking = false,
            DispatcherPriority.Background);
    }

    private void FocusEditor()
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (IsEditorPageVisible && IsTextEditorVisible)
                EditorTextBox.TextArea.Focus();
        }, DispatcherPriority.Background);
    }

    // ── Event handlers ───────────────────────────────────────────────────────

    // Switches the visible page in one pass — sets all backing fields before firing
    // any notifications, so the UI only re-renders once instead of once per property set.
    private enum Page { Home, Editor, Settings, Extensions, WhatsNew }

    private void NavigateTo(Page page)
    {
        var newHome       = page == Page.Home;
        var newSettings   = page == Page.Settings;
        var newExtensions = page == Page.Extensions;
        var newWhatsNew   = page == Page.WhatsNew;

        // Bail early if nothing actually changed
        if (_isHomePageVisible       == newHome       &&
            _isSettingsPageVisible   == newSettings   &&
            _isExtensionsPageVisible == newExtensions &&
            _isWhatsNewPageVisible   == newWhatsNew)
            return;

        _isHomePageVisible       = newHome;
        _isSettingsPageVisible   = newSettings;
        _isExtensionsPageVisible = newExtensions;
        _isWhatsNewPageVisible   = newWhatsNew;

        OnPropertyChanged(nameof(IsHomePageVisible));
        OnPropertyChanged(nameof(IsSettingsPageVisible));
        OnPropertyChanged(nameof(IsExtensionsPageVisible));
        OnPropertyChanged(nameof(IsWhatsNewPageVisible));
        OnPropertyChanged(nameof(IsEditorPageVisible));
        OnPropertyChanged(nameof(IsEditorTabsVisible));
        OnPropertyChanged(nameof(IsDocumentViewVisible));
        OnPropertyChanged(nameof(IsEmptyStateVisible));
        OnPropertyChanged(nameof(CanShowSaveActions));
        OnPropertyChanged(nameof(FileSummaryText));
        OnPropertyChanged(nameof(FilePathText));
        RefreshState();
    }

    private void EditorButton_OnClick(object? sender, RoutedEventArgs e)
    {
        NavigateTo(Page.Editor);
        FocusEditor();
    }

    private void HomeButton_OnClick(object? sender, RoutedEventArgs e) =>
        NavigateTo(Page.Home);

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

    private void CollapseExplorerButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (IsFolderOpen)
            CloseFolder();
        else
            IsFileExplorerVisible = !IsFileExplorerVisible;
    }

    private void SettingsButton_OnClick(object? sender, RoutedEventArgs e) =>
        NavigateTo(Page.Settings);

    private void ExtensionsButton_OnClick(object? sender, RoutedEventArgs e)
    {
        NavigateTo(Page.Extensions);
        _ = RefreshExtensionsDataAsync();
    }

    private async void RefreshExtensionsButton_OnClick(object? sender, RoutedEventArgs e) =>
        await RefreshExtensionsDataAsync(force: true);

    private void InstalledTabButton_OnClick(object? sender, RoutedEventArgs e) =>
        IsMarketplaceTabSelected = false;

    // Used by the tab strip inside the Extensions page — only switches the tab
    private void MarketplaceTabButton_OnClick(object? sender, RoutedEventArgs e) =>
        IsMarketplaceTabSelected = true;

    // Used by the "Visit Marketplace" button on the home screen —
    // opens the Extensions page AND switches to the Marketplace tab
    private void OpenMarketplaceButton_OnClick(object? sender, RoutedEventArgs e)
    {
        IsSettingsPageVisible = false;
        IsWhatsNewPageVisible = false;
        IsHomePageVisible = false;
        IsMarketplaceTabSelected = true;
        IsExtensionsPageVisible = true;
        _ = RefreshExtensionsDataAsync(force: true);
    }

    private void OpenWhatsNewButton_OnClick(object? sender, RoutedEventArgs e)
    {
        IsExtensionsPageVisible = false;
        IsSettingsPageVisible = false;
        IsHomePageVisible = false;
        IsWhatsNewPageVisible = true;
        IsWhatsNewExpanded = true;
        _ = RefreshLatestReleaseAsync();
    }

    private void BackToEditorButton_OnClick(object? sender, RoutedEventArgs e)
    {
        IsSettingsPageVisible = false;
        IsExtensionsPageVisible = false;
        IsWhatsNewPageVisible = false;
        IsHomePageVisible = false;
        FocusEditor();
    }

    private async void RefreshLatestReleaseButton_OnClick(object? sender, RoutedEventArgs e) =>
        await RefreshLatestReleaseAsync();

    private void ToggleWhatsNewExpandedButton_OnClick(object? sender, RoutedEventArgs e) =>
        IsWhatsNewExpanded = !IsWhatsNewExpanded;

    private void DismissUpdateBanner_OnClick(object? sender, RoutedEventArgs e)
    {
        _updateBannerDismissed = true;
        OnPropertyChanged(nameof(IsAppUpdateAvailable));
    }

    private void DismissExtensionUpdateBanner_OnClick(object? sender, RoutedEventArgs e)
    {
        _extensionUpdateBannerDismissed = true;
        OnPropertyChanged(nameof(IsExtensionUpdateBannerVisible));
    }

    private void OpenMarketplaceFromBannerButton_OnClick(object? sender, RoutedEventArgs e)
    {
        _extensionUpdateBannerDismissed = true;
        OnPropertyChanged(nameof(IsExtensionUpdateBannerVisible));
        IsSettingsPageVisible = false;
        IsWhatsNewPageVisible = false;
        IsHomePageVisible = false;
        IsMarketplaceTabSelected = true;
        IsExtensionsPageVisible = true;
    }

    private void OpenReleasesPageButton_OnClick(object? sender, RoutedEventArgs e) =>
        OpenUrl(ReleasesPageUrl);

    private void OpenLatestReleaseButton_OnClick(object? sender, RoutedEventArgs e) =>
        OpenUrl(LatestReleaseUrl);

    private void OpenReleaseLinkButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: string url })
            OpenUrl(url);
    }

    private static void OpenUrl(string url)
    {
        if (string.IsNullOrWhiteSpace(url))
            return;

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = url,
                UseShellExecute = true
            });
        }
        catch
        {
            // Ignore browser launch failures quietly; the release info remains visible in-app.
        }
    }

    private async void OpenExtensionsFolderButton_OnClick(object? sender, RoutedEventArgs e)
    {
        try
        {
            if (!Directory.Exists(ExtensionsFolderPath))
                Directory.CreateDirectory(ExtensionsFolderPath);

            Process.Start(new ProcessStartInfo
            {
                FileName        = ExtensionsFolderPath,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            ExtensionsStatusText = $"Could not open extensions folder: {ex.Message}";
            await ShowWarningDialogAsync("Open extensions folder", ex);
        }
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
                await ToggleDirectoryExpansionAsync(item);
            else
                await OpenFileFromPathAsync(item.FullPath);
        }
    }

    private async void RecentFileButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: RecentFileItem item }) return;

        if (item.IsFolder)
        {
            if (!Directory.Exists(item.Path)) { RemoveRecentFile(item.Path); return; }
            _currentFolderPath = item.Path;
            AddRecentFolder(item.Path);
            await PopulateFileTreeAsync(item.Path);
            IsFileExplorerVisible = true;
            RefreshState();
        }
        else
        {
            if (!File.Exists(item.Path)) { RemoveRecentFile(item.Path); return; }
            await OpenFileFromPathAsync(item.Path);
        }
    }

    private void OpenEditorTabButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: EditorTab tab })
            ActivateTab(tab);
    }

    private async void CloseEditorTabButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<EditorTab>(sender) is { } tab)
            await RequestCloseTabAsync(tab);
    }

    private async void CloseOtherTabsButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is not MenuItem { Tag: EditorTab pivotTab }) return;
        var others = OpenTabs.Where(t => !ReferenceEquals(t, pivotTab)).ToList();
        foreach (var tab in others)
        {
            if (!await RequestCloseTabAsync(tab))
                break;
        }
    }

    private async void CloseAllTabsButton_OnClick(object? sender, RoutedEventArgs e)
    {
        var all = OpenTabs.ToList();
        foreach (var tab in all)
        {
            if (!await RequestCloseTabAsync(tab))
                break;
        }
    }

    private static T? TryGetTaggedData<T>(object? sender) where T : class =>
        sender switch
        {
            MenuItem { Tag: T taggedItem } => taggedItem,
            Button { Tag: T taggedButton } => taggedButton,
            _ => null
        };

    private string GetRelativePathOrFullPath(string path)
    {
        if (string.IsNullOrWhiteSpace(_currentFolderPath))
            return path;

        try
        {
            return Path.GetRelativePath(_currentFolderPath, path);
        }
        catch
        {
            return path;
        }
    }

    private string GetExplorerTargetDirectory(FileTreeItem item) =>
        item.IsDirectory
            ? item.FullPath
            : Path.GetDirectoryName(item.FullPath) ?? _currentFolderPath ?? item.FullPath;

    private static string CreateUniqueSiblingPath(string path, bool isDirectory)
    {
        var directory = Path.GetDirectoryName(path) ?? string.Empty;
        var fileName = Path.GetFileName(path);
        string extension;
        string baseName;

        if (isDirectory || (fileName.StartsWith('.') && fileName.Count(ch => ch == '.') == 1))
        {
            extension = string.Empty;
            baseName = fileName;
        }
        else
        {
            extension = Path.GetExtension(fileName);
            baseName = Path.GetFileNameWithoutExtension(fileName);
        }

        for (var index = 1; ; index++)
        {
            var suffix = index == 1 ? " - Copy" : $" - Copy ({index})";
            var candidate = Path.Combine(directory, $"{baseName}{suffix}{extension}");
            if (!File.Exists(candidate) && !Directory.Exists(candidate))
                return candidate;
        }
    }

    private static string CreateUniqueChildPath(string directory, string baseName, string extension = "")
    {
        for (var index = 1; ; index++)
        {
            var candidateName = index == 1 ? $"{baseName}{extension}" : $"{baseName} ({index}){extension}";
            var candidate = Path.Combine(directory, candidateName);
            if (!File.Exists(candidate) && !Directory.Exists(candidate))
                return candidate;
        }
    }

    private static void CopyDirectoryRecursive(string sourceDirectory, string destinationDirectory)
    {
        Directory.CreateDirectory(destinationDirectory);

        foreach (var file in Directory.GetFiles(sourceDirectory))
        {
            var destinationFile = Path.Combine(destinationDirectory, Path.GetFileName(file));
            File.Copy(file, destinationFile, overwrite: false);
        }

        foreach (var directory in Directory.GetDirectories(sourceDirectory))
        {
            var destinationChild = Path.Combine(destinationDirectory, Path.GetFileName(directory));
            CopyDirectoryRecursive(directory, destinationChild);
        }
    }

    private async Task OpenPathInSystemExplorer(string path, bool selectItem)
    {
        try
        {
            if (OperatingSystem.IsWindows())
            {
                var startInfo = selectItem && File.Exists(path)
                    ? new ProcessStartInfo
                    {
                        FileName = "explorer.exe",
                        Arguments = $"/select,\"{path}\"",
                        UseShellExecute = true
                    }
                    : new ProcessStartInfo
                    {
                        FileName = Directory.Exists(path) ? path : Path.GetDirectoryName(path) ?? path,
                        UseShellExecute = true
                    };

                Process.Start(startInfo);
                return;
            }

            if (OperatingSystem.IsMacOS())
            {
                // `open -R <path>` reveals and selects the item in Finder.
                // Falls back to opening the parent directory if path doesn't exist.
                var target = selectItem && (File.Exists(path) || Directory.Exists(path)) ? path
                    : Directory.Exists(path) ? path
                    : Path.GetDirectoryName(path) ?? path;

                Process.Start(new ProcessStartInfo
                {
                    FileName = "open",
                    Arguments = $"-R \"{target}\"",
                    UseShellExecute = false
                });
                return;
            }

            // Linux: try common file managers that support --select / reveal flags,
            // then fall back to opening the parent directory via xdg-open.
            if (OperatingSystem.IsLinux() && selectItem && File.Exists(path))
            {
                var fileManagers = new[]
                {
                    ("nautilus", $"--select \"{path}\""),   // GNOME
                    ("dolphin",  $"--select \"{path}\""),   // KDE
                    ("nemo",     $"\"{path}\""),             // Cinnamon
                    ("thunar",   $"\"{Path.GetDirectoryName(path)}\""), // XFCE (no select)
                };

                foreach (var (binary, args) in fileManagers)
                {
                    // Check the binary exists before trying to launch it.
                    var which = Process.Start(new ProcessStartInfo
                    {
                        FileName = "which",
                        Arguments = binary,
                        UseShellExecute = false,
                        RedirectStandardOutput = true
                    });
                    which?.WaitForExit();
                    if (which?.ExitCode != 0) continue;

                    Process.Start(new ProcessStartInfo
                    {
                        FileName = binary,
                        Arguments = args,
                        UseShellExecute = false
                    });
                    return;
                }
            }

            // Universal fallback: open the containing directory.
            var fallbackDir = Directory.Exists(path) ? path : Path.GetDirectoryName(path) ?? path;
            Process.Start(new ProcessStartInfo { FileName = fallbackDir, UseShellExecute = true });
        }
        catch (Exception ex)
        {
            ExtensionsStatusText = $"Could not open path: {ex.Message}";
            await ShowWarningDialogAsync("Open in system explorer", ex);
        }
    }

    private async Task RefreshExplorerTreeAsync()
    {
        if (!string.IsNullOrWhiteSpace(_currentFolderPath) && Directory.Exists(_currentFolderPath))
            await PopulateFileTreeAsync(_currentFolderPath);
    }

    private async Task<bool> CloseTabsForPathAsync(string path, bool isDirectory)
    {
        var matchingTabs = OpenTabs
            .Where(tab => !tab.IsUntitled && (
                string.Equals(tab.Path, path, StringComparison.OrdinalIgnoreCase) ||
                (isDirectory && tab.Path.StartsWith(path + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))))
            .ToList();

        foreach (var tab in matchingTabs)
        {
            if (!await RequestCloseTabAsync(tab))
                return false;
        }

        return true;
    }

    private async Task<bool> EnsureTabsReadyForDeletionAsync(List<EditorTab> tabs)
    {
        var originalActiveTab = ActiveEditorTab;

        foreach (var tab in tabs.Where(t => t.IsDirty))
        {
            var action = await ShowUnsavedTabDialogAsync(tab);
            switch (action)
            {
                case UnsavedTabAction.Cancel:
                    if (originalActiveTab is not null && OpenTabs.Contains(originalActiveTab))
                        ActivateTab(originalActiveTab, focusEditor: false, preserveCurrentState: false);
                    return false;

                case UnsavedTabAction.Save:
                    if (!ReferenceEquals(tab, ActiveEditorTab))
                        ActivateTab(tab, focusEditor: false);

                    if (!await SaveAsync(allowPromptForPath: true, forcePromptForPath: false))
                    {
                        if (originalActiveTab is not null && OpenTabs.Contains(originalActiveTab))
                            ActivateTab(originalActiveTab, focusEditor: false, preserveCurrentState: false);
                        return false;
                    }
                    break;
            }
        }

        if (originalActiveTab is not null && OpenTabs.Contains(originalActiveTab))
            ActivateTab(originalActiveTab, focusEditor: false, preserveCurrentState: false);

        return true;
    }

    private void CloseTabsWithoutPrompt(IEnumerable<EditorTab> tabs)
    {
        foreach (var tab in tabs.ToList())
            CloseTab(tab);
    }

    /// <summary>
    /// After a rename or move, update every open tab whose path falls under
    /// <paramref name="oldPath"/> so it points to the new location.
    /// Also patches <c>_currentFilePath</c> when the active tab is affected.
    /// </summary>
    private void RetargetTabPaths(string oldPath, string newPath, bool wasDirectory)
    {
        foreach (var tab in OpenTabs.Where(t => !t.IsUntitled))
        {
            string? updated = null;

            if (wasDirectory)
            {
                var prefix = oldPath + Path.DirectorySeparatorChar;
                if (tab.Path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    updated = newPath + Path.DirectorySeparatorChar + tab.Path[prefix.Length..];
            }
            else if (string.Equals(tab.Path, oldPath, StringComparison.OrdinalIgnoreCase))
            {
                updated = newPath;
            }

            if (updated is null) continue;

            tab.Rename(updated, Path.GetFileName(updated));

            if (ReferenceEquals(tab, ActiveEditorTab))
            {
                _currentFilePath = updated;
                RefreshState();
            }
        }
    }

    /// <summary>
    /// Shows a small modal asking the user for a new name.
    /// Returns the trimmed input, or null if cancelled / empty.
    /// </summary>
    private async Task<string?> ShowRenameDialogAsync(string currentName)
    {
        string? result = null;
        Window? dialog = null;

        var inputBox = new TextBox
        {
            Text             = currentName,
            Background       = ButtonBrush,
            Foreground       = PrimaryTextBrush,
            BorderBrush      = SurfaceBorderBrush,
            Padding          = new Thickness(8, 6),
            FontSize         = 14,
            CaretBrush       = PrimaryTextBrush,
        };

        var confirmButton = CreateDialogButton("Rename", AccentBrush, AccentBrush, Brushes.White, () =>
        {
            result = inputBox.Text?.Trim();
            dialog!.Close();
        });

        inputBox.KeyDown += (_, e) =>
        {
            if (e.Key == Key.Enter)  { result = inputBox.Text?.Trim(); dialog!.Close(); }
            if (e.Key == Key.Escape) { dialog!.Close(); }
        };

        dialog = new Window
        {
            Width                   = 380,
            Height                  = 160,
            CanResize               = false,
            ShowInTaskbar           = false,
            WindowStartupLocation   = WindowStartupLocation.CenterOwner,
            Title                   = "Rename",
            Background              = CardBrush,
            Content = new Border
            {
                Padding = new Thickness(20),
                Child   = new StackPanel
                {
                    Spacing  = 14,
                    Children =
                    {
                        new TextBlock { Text = "Enter a new name:", FontSize = 15,
                            FontWeight = FontWeight.SemiBold, Foreground = PrimaryTextBrush },
                        inputBox,
                        new StackPanel
                        {
                            Orientation         = Orientation.Horizontal,
                            Spacing             = 10,
                            HorizontalAlignment = HorizontalAlignment.Right,
                            Children =
                            {
                                CreateDialogButton("Cancel", ButtonBrush, SurfaceBorderBrush, PrimaryTextBrush,
                                    () => dialog!.Close()),
                                confirmButton
                            }
                        }
                    }
                }
            }
        };

        dialog.Opened += (_, _) =>
        {
            inputBox.Focus();
            inputBox.SelectAll();
        };

        await dialog.ShowDialog(this);
        return string.IsNullOrWhiteSpace(result) ? null : result;
    }

    private async Task ActivateEditorTabForMenuActionAsync(EditorTab tab)
    {
        if (!ReferenceEquals(ActiveEditorTab, tab))
            ActivateTab(tab, focusEditor: false);

        await Task.CompletedTask;
    }

    private async void SaveEditorTabMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<EditorTab>(sender) is not { } tab) return;
        await ActivateEditorTabForMenuActionAsync(tab);
        await SaveAsync();
    }

    private async void SaveEditorTabAsMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<EditorTab>(sender) is not { } tab) return;
        await ActivateEditorTabForMenuActionAsync(tab);
        await SaveAsAsync();
        await RefreshExplorerTreeAsync();
    }

    private void CopyEditorTabPathMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<EditorTab>(sender) is not { IsUntitled: false } tab) return;
        TopLevel.GetTopLevel(this)?.Clipboard?.SetTextAsync(tab.Path);
    }

    private void CopyEditorTabRelativePathMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<EditorTab>(sender) is not { IsUntitled: false } tab) return;
        TopLevel.GetTopLevel(this)?.Clipboard?.SetTextAsync(GetRelativePathOrFullPath(tab.Path));
    }

    private async void RevealEditorTabInExplorerMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<EditorTab>(sender) is not { IsUntitled: false } tab) return;
        await OpenPathInSystemExplorer(tab.Path, selectItem: true);
    }

    private async void CollapseAllTreeButton_OnClick(object? sender, RoutedEventArgs e)
    {
        _isFileTreeExpanded = !_isFileTreeExpanded;
        if (!_isFileTreeExpanded)
        {
            // Collapse: repopulate from scratch — fastest correct collapse
            if (!string.IsNullOrWhiteSpace(_currentFolderPath))
                await PopulateFileTreeAsync(_currentFolderPath);
        }
        else
        {
            // Expand: toggle all top-level directories
            var rootDirs = FileTreeItems.Where(i => i.IsDirectory && i.Depth == 0 && !i.IsExpanded).ToList();
            foreach (var dir in rootDirs)
                await ToggleDirectoryExpansionAsync(dir);
        }
    }

    private async void NewFileInExplorerMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<FileTreeItem>(sender) is not { } item) return;

        try
        {
            var directory = GetExplorerTargetDirectory(item);
            var newFilePath = CreateUniqueChildPath(directory, "new-file", ".txt");
            await File.WriteAllTextAsync(newFilePath, string.Empty);
            await RefreshExplorerTreeAsync();
            await OpenFileFromPathAsync(newFilePath);
        }
        catch (Exception ex)
        {
            ExtensionsStatusText = $"New file failed: {ex.Message}";
            await ShowWarningDialogAsync("New file in explorer", ex);
        }
    }

    private async void NewFolderInExplorerMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<FileTreeItem>(sender) is not { } item) return;

        try
        {
            var directory = GetExplorerTargetDirectory(item);
            var newFolderPath = CreateUniqueChildPath(directory, "New Folder");
            Directory.CreateDirectory(newFolderPath);
            await RefreshExplorerTreeAsync();
        }
        catch (Exception ex)
        {
            ExtensionsStatusText = $"New folder failed: {ex.Message}";
            await ShowWarningDialogAsync("New folder in explorer", ex);
        }
    }

    private async void OpenExplorerItemMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<FileTreeItem>(sender) is not { } item) return;

        if (item.IsDirectory)
            await ToggleDirectoryExpansionAsync(item);
        else
            await OpenFileFromPathAsync(item.FullPath);
    }

    private async void RevealExplorerItemMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<FileTreeItem>(sender) is not { } item) return;
        await OpenPathInSystemExplorer(item.FullPath, selectItem: !item.IsDirectory);
    }

    private void CopyFileNameMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<FileTreeItem>(sender) is not { } item) return;
        TopLevel.GetTopLevel(this)?.Clipboard?.SetTextAsync(item.Name);
    }

    private void CopyFilePathMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<FileTreeItem>(sender) is { } item)
            TopLevel.GetTopLevel(this)?.Clipboard?.SetTextAsync(item.FullPath);
    }

    private void CopyRelativeFilePathMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<FileTreeItem>(sender) is not { } item) return;
        TopLevel.GetTopLevel(this)?.Clipboard?.SetTextAsync(GetRelativePathOrFullPath(item.FullPath));
    }

    private async void DuplicateExplorerItemMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<FileTreeItem>(sender) is not { } item) return;

        try
        {
            var duplicatePath = CreateUniqueSiblingPath(item.FullPath, item.IsDirectory);
            if (item.IsDirectory)
                CopyDirectoryRecursive(item.FullPath, duplicatePath);
            else
                File.Copy(item.FullPath, duplicatePath, overwrite: false);

            await RefreshExplorerTreeAsync();
        }
        catch (Exception ex)
        {
            ExtensionsStatusText = $"Duplicate failed: {ex.Message}";
            await ShowWarningDialogAsync("Duplicate file", ex);
        }
    }

    private async void DeleteFileMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<FileTreeItem>(sender) is not { } item) return;
        try
        {
            var matchingTabs = OpenTabs
                .Where(tab => !tab.IsUntitled && (
                    string.Equals(tab.Path, item.FullPath, StringComparison.OrdinalIgnoreCase) ||
                    (item.IsDirectory && tab.Path.StartsWith(item.FullPath + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))))
                .ToList();

            if (!await EnsureTabsReadyForDeletionAsync(matchingTabs))
                return;

            if (item.IsDirectory)
                Directory.Delete(item.FullPath, recursive: true);
            else
                File.Delete(item.FullPath);

            CloseTabsWithoutPrompt(matchingTabs);
            await RefreshExplorerTreeAsync();
        }
        catch (Exception ex)
        {
            ExtensionsStatusText = $"Delete failed: {ex.Message}";
            await ShowWarningDialogAsync("Delete file", ex);
        }
    }

    private async void RenameFileMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<FileTreeItem>(sender) is not { } item) return;

        var newName = await ShowRenameDialogAsync(item.Name);
        if (newName is null || string.Equals(newName, item.Name, StringComparison.Ordinal)) return;

        var newPath = Path.Combine(Path.GetDirectoryName(item.FullPath)!, newName);

        if ((item.IsDirectory ? Directory.Exists(newPath) : File.Exists(newPath)))
        {
            ExtensionsStatusText = $"Rename failed: '{newName}' already exists.";
            return;
        }

        // If dirty tabs are open under this path, ask to save first
        var affectedTabs = OpenTabs
            .Where(t => !t.IsUntitled && (
                string.Equals(t.Path, item.FullPath, StringComparison.OrdinalIgnoreCase) ||
                (item.IsDirectory && t.Path.StartsWith(item.FullPath + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))))
            .ToList();

        if (!await EnsureTabsReadyForDeletionAsync(affectedTabs)) return;

        try
        {
            if (item.IsDirectory)
                Directory.Move(item.FullPath, newPath);
            else
                File.Move(item.FullPath, newPath);

            RetargetTabPaths(item.FullPath, newPath, item.IsDirectory);
            await RefreshExplorerTreeAsync();
        }
        catch (Exception ex)
        {
            ExtensionsStatusText = $"Rename failed: {ex.Message}";
            await ShowWarningDialogAsync("Rename file", ex);
        }
    }

    private void CutFileMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<FileTreeItem>(sender) is not { } item) return;
        _clipboardItemPath        = item.FullPath;
        _clipboardItemIsDirectory = item.IsDirectory;
        _clipboardIsCut           = true;
        ExtensionsStatusText      = $"Cut: {item.Name}";
    }

    private void CopyFileMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<FileTreeItem>(sender) is not { } item) return;
        _clipboardItemPath        = item.FullPath;
        _clipboardItemIsDirectory = item.IsDirectory;
        _clipboardIsCut           = false;
        ExtensionsStatusText      = $"Copied: {item.Name}";
    }

    private async void PasteFileMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (_clipboardItemPath is null) return;
        if (TryGetTaggedData<FileTreeItem>(sender) is not { } target) return;

        // Paste destination is the target's directory (or the target itself if it's a folder)
        var destDir = target.IsDirectory ? target.FullPath : Path.GetDirectoryName(target.FullPath)!;
        var itemName = Path.GetFileName(_clipboardItemPath.TrimEnd(Path.DirectorySeparatorChar));
        var destPath = CreateUniqueSiblingPath(Path.Combine(destDir, itemName), _clipboardItemIsDirectory);

        try
        {
            if (_clipboardIsCut)
            {
                if (_clipboardItemIsDirectory)
                    Directory.Move(_clipboardItemPath, destPath);
                else
                    File.Move(_clipboardItemPath, destPath);

                RetargetTabPaths(_clipboardItemPath, destPath, _clipboardItemIsDirectory);
                _clipboardItemPath = null; // cut is consumed
            }
            else
            {
                if (_clipboardItemIsDirectory)
                    CopyDirectoryRecursive(_clipboardItemPath, destPath);
                else
                    File.Copy(_clipboardItemPath, destPath, overwrite: false);
            }

            await RefreshExplorerTreeAsync();
        }
        catch (Exception ex)
        {
            ExtensionsStatusText = $"Paste failed: {ex.Message}";
            await ShowWarningDialogAsync("Paste file", ex);
        }
    }

    private void ToggleFindPanel_OnClick(object? sender, RoutedEventArgs e)
    {
        if (!CanShowFindInFile)
        {
            IsFindPanelVisible = false;
            return;
        }

        IsFindPanelVisible = !IsFindPanelVisible;
    }

    private void FindNextButton_OnClick(object? sender, RoutedEventArgs e) =>
        FindInEditor(forward: true);

    private void FindPrevButton_OnClick(object? sender, RoutedEventArgs e) =>
        FindInEditor(forward: false);

    private void CloseFindPanel_OnClick(object? sender, RoutedEventArgs e) =>
        IsFindPanelVisible = false;

    private void FindInEditor(bool forward)
    {
        if (string.IsNullOrEmpty(FindText) || EditorTextBox?.Document is null) return;
        var doc = EditorTextBox.Document.Text;
        var caretOffset = EditorTextBox.TextArea.Caret.Offset;
        int index;
        if (forward)
        {
            index = doc.IndexOf(FindText, caretOffset + 1, StringComparison.OrdinalIgnoreCase);
            if (index < 0)
                index = doc.IndexOf(FindText, 0, StringComparison.OrdinalIgnoreCase);
        }
        else
        {
            var searchTo = Math.Max(0, caretOffset - 1);
            index = doc.LastIndexOf(FindText, searchTo, StringComparison.OrdinalIgnoreCase);
            if (index < 0)
                index = doc.LastIndexOf(FindText, doc.Length - 1, StringComparison.OrdinalIgnoreCase);
        }

        if (index < 0) return;
        EditorTextBox.TextArea.Caret.Offset = index;
        EditorTextBox.TextArea.Selection = AvaloniaEdit.Editing.Selection.Create(EditorTextBox.TextArea, index, index + FindText.Length);
        EditorTextBox.ScrollToLine(EditorTextBox.Document.GetLineByOffset(index).LineNumber);
    }

    private void IncreaseFontSizeButton_OnClick(object? sender, RoutedEventArgs e) =>
        EditorFontSize = Math.Min(32, EditorFontSize + 1);

    private void DecreaseFontSizeButton_OnClick(object? sender, RoutedEventArgs e) =>
        EditorFontSize = Math.Max(8, EditorFontSize - 1);

    private void ClearRecentFilesButton_OnClick(object? sender, RoutedEventArgs e)
    {
        RecentFiles.Clear();
        SaveSettings();
        OnPropertyChanged(nameof(HasRecentFiles));
    }

    private async void InstallMarketplaceExtensionButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: MarketplaceExtension marketplaceExtension })
            await InstallMarketplaceExtensionAsync(marketplaceExtension);
    }

    private async void UpdateInstalledExtensionButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: LoadedExtension extension })
            return;

        var marketplaceExtension = GetMarketplaceExtensionForInstalled(extension);
        if (marketplaceExtension is null)
        {
            ExtensionsStatusText = $"No marketplace update source found for {extension.Name}.";
            return;
        }

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
        _rainbowBracketColorizer.InvalidateCache();
        if (_suppressDirtyTracking) return;
        ClearAutoSaveStatus();
        _isDirty = true;
        if (ActiveEditorTab is not null)
        {
            ActiveEditorTab.Content = EditorTextBox.Document.Text;
            ActiveEditorTab.IsDirty = true;
        }
        QueueRefreshState();
        RestartAutoSaveTimerIfNeeded();
    }

	// Fires BEFORE the character is written into the document.
    // Used to skip-over an already-present auto-inserted closing character
    // instead of inserting a duplicate.
    private void EditorTextArea_OnTextEntering(object? sender, TextInputEventArgs e)
    {
        if (IsPlainTextMode()) return;
        if (string.IsNullOrEmpty(e.Text)) return;
        var ch     = e.Text[0];
        var caret  = EditorTextBox.TextArea.Caret;
        var doc    = EditorTextBox.Document;
        var offset = caret.Offset;

        if (!ClosingChars.Contains(ch)) return;
        if (offset >= doc.TextLength) return;
        if (doc.GetCharAt(offset) != ch) return;

        // Asymmetric pairs (closing char differs from opener): always safe to skip.
        // Symmetric pairs (" and '): only skip when the char immediately behind the
        // caret is the same quote, meaning we auto-inserted it and the caret is
        // sitting between the pair.
        bool skip = ch is ')' or ']' or '}' or '>';
        if (!skip && (ch == '"' || ch == '\''))
            skip = offset > 0 && doc.GetCharAt(offset - 1) == ch;

        if (skip)
        {
            caret.Offset = offset + 1;
            e.Handled = true;
        }
    }

    // Fires AFTER the character has been written into the document.
    // Used to insert the matching closing character right after the opener.
    private void EditorTextArea_OnTextEntered(object? sender, TextInputEventArgs e)
    {
        if (IsPlainTextMode()) return;
        if (string.IsNullOrEmpty(e.Text)) return;
        var ch = e.Text[0];

        if (!BracketPairs.TryGetValue(ch, out var closing)) return;

        var caret  = EditorTextBox.TextArea.Caret;
        var doc    = EditorTextBox.Document;
        var offset = caret.Offset;
        var selection = EditorTextBox.TextArea.Selection;

        if (!selection.IsEmpty)
        {
            var startOffset = selection.SurroundingSegment.Offset;
            var selectedText = selection.GetText();
            doc.Replace(selection.SurroundingSegment, $"{ch}{selectedText}{closing}");
            caret.Offset = startOffset + selectedText.Length + 2;
            return;
        }

        // For symmetric pairs, don't auto-close when the next char is alphanumeric
        // (avoids nuisance completions mid-word, e.g. typing " in  it's).
        if (ch == '"' || ch == '\'' || ch == '`')
        {
            if (offset < doc.TextLength)
            {
                var next = doc.GetCharAt(offset);
                if (char.IsLetterOrDigit(next) || next == ch) return;
            }
        }

        // Insert the closer and explicitly restore the caret to the space between the pair.
        doc.Insert(offset, closing.ToString());
        caret.Offset = offset;
    }

    private void MainWindow_EditorKeyIntercept_OnKeyDown(object? sender, KeyEventArgs e)
    {
        if (!IsEditorKeyEvent(e))
            return;

        if (EditorTextBox?.Document is null)
            return;

        var textArea = EditorTextBox.TextArea;
        var caret = textArea.Caret;
        var doc = EditorTextBox.Document;
        switch (e.Key)
        {
            case Key.Enter when !IsPlainTextMode() && (e.KeyModifiers & KeyModifiers.Shift) != KeyModifiers.Shift:
                HandleSmartEnter(doc, caret);
                e.Handled = true;
                return;

            case Key.Tab when e.KeyModifiers == KeyModifiers.Shift:
                HandleTabKey(() => HandleOutdent(doc, textArea.Selection, caret), doc, caret);
                e.Handled = true;
                return;

            case Key.Tab when e.KeyModifiers == KeyModifiers.None:
                HandleTabKey(() => HandleIndent(doc, textArea.Selection, caret), doc, caret);
                e.Handled = true;
                return;
        }
    }

    private bool IsEditorKeyEvent(KeyEventArgs e)
    {
        if (EditorTextBox is null || e.Source is not Visual visual)
            return false;

        if (ReferenceEquals(visual, EditorTextBox) || ReferenceEquals(visual, EditorTextBox.TextArea))
            return true;

        return visual.GetSelfAndVisualAncestors().Any(v =>
            ReferenceEquals(v, EditorTextBox) || ReferenceEquals(v, EditorTextBox.TextArea));
    }

    private void HandleTabKey(Action tabAction, AvaloniaEdit.Document.TextDocument doc, AvaloniaEdit.Editing.Caret caret)
    {
        try
        {
            tabAction();
        }
        catch
        {
            // Keep Tab editor-local even if AvaloniaEdit reports an invalid selection snapshot.
            var safeOffset = Math.Clamp(caret.Offset, 0, doc.TextLength);
            doc.Insert(safeOffset, GetIndentUnit());
            SetCaretOffsetSafely(caret, doc, safeOffset + GetIndentUnit().Length);
        }
    }

    private static void SetCaretOffsetSafely(
        AvaloniaEdit.Editing.Caret caret,
        AvaloniaEdit.Document.TextDocument doc,
        int desiredOffset)
    {
        caret.Offset = Math.Clamp(desiredOffset, 0, doc.TextLength);
    }

    private void HandleSmartEnter(AvaloniaEdit.Document.TextDocument doc, AvaloniaEdit.Editing.Caret caret)
    {
        var offset = caret.Offset;
        var line = doc.GetLineByOffset(offset);
        var lineText = doc.GetText(line);
        var caretColumnInLine = offset - line.Offset;
        var textBeforeCaret = lineText[..Math.Min(caretColumnInLine, lineText.Length)];
        var textAfterCaret = lineText[Math.Min(caretColumnInLine, lineText.Length)..];
        var indent = GetLeadingWhitespace(textBeforeCaret);
        var trimmedBeforeCaret = textBeforeCaret.TrimEnd();
        var extraIndent = ShouldIncreaseIndentAfter(trimmedBeforeCaret) ? GetIndentUnit() : string.Empty;

        if (ShouldInsertStructuredBlock(trimmedBeforeCaret, textAfterCaret))
        {
            var blockText = Environment.NewLine + indent + extraIndent + Environment.NewLine + indent;
            doc.Insert(offset, blockText);
            caret.Offset = offset + Environment.NewLine.Length + indent.Length + extraIndent.Length;
            return;
        }

        var adjustedIndent = StartsWithClosingDelimiter(textAfterCaret)
            ? RemoveOneIndentUnit(indent)
            : indent;

        var newLineText = Environment.NewLine + adjustedIndent + extraIndent;
        doc.Insert(offset, newLineText);
        caret.Offset = offset + newLineText.Length;
    }

    private void HandleIndent(AvaloniaEdit.Document.TextDocument doc, AvaloniaEdit.Editing.Selection? selection, AvaloniaEdit.Editing.Caret caret)
    {
        if (selection is null || selection.IsEmpty)
        {
            var safeOffset = Math.Clamp(caret.Offset, 0, doc.TextLength);
            doc.Insert(safeOffset, GetIndentUnit());
            SetCaretOffsetSafely(caret, doc, safeOffset + GetIndentUnit().Length);
            return;
        }

        var segment = selection.SurroundingSegment;
        if (segment is null)
        {
            var safeOffset = Math.Clamp(caret.Offset, 0, doc.TextLength);
            doc.Insert(safeOffset, GetIndentUnit());
            SetCaretOffsetSafely(caret, doc, safeOffset + GetIndentUnit().Length);
            return;
        }

        var lines = GetSelectedLines(doc, segment.Offset, segment.EndOffset);
        foreach (var line in lines.OrderByDescending(l => l.Offset))
            doc.Insert(line.Offset, GetIndentUnit());

        SetCaretOffsetSafely(caret, doc, segment.EndOffset + (GetIndentUnit().Length * lines.Count));
    }

    private void HandleOutdent(AvaloniaEdit.Document.TextDocument doc, AvaloniaEdit.Editing.Selection? selection, AvaloniaEdit.Editing.Caret caret)
    {
        if (selection is null || selection.IsEmpty)
        {
            var line = doc.GetLineByOffset(caret.Offset);
            var lineText = doc.GetText(line);
            var caretColumnInLine = caret.Offset - line.Offset;
            var removable = GetOutdentLength(lineText, caretColumnInLine);
            if (removable <= 0)
                return;

            doc.Remove(line.Offset, removable);
            SetCaretOffsetSafely(caret, doc, caret.Offset - removable);
            return;
        }

        var segment = selection.SurroundingSegment;
        if (segment is null)
            return;

        var lines = GetSelectedLines(doc, segment.Offset, segment.EndOffset);
        var removed = 0;
        foreach (var line in lines.OrderByDescending(l => l.Offset))
        {
            var lineText = doc.GetText(line);
            var removable = GetOutdentLength(lineText, lineText.Length);
            if (removable <= 0)
                continue;

            doc.Remove(line.Offset, removable);
            removed += removable;
        }

        SetCaretOffsetSafely(caret, doc, Math.Max(segment.Offset, segment.EndOffset - removed));
    }

    private static string GetLeadingWhitespace(string text)
    {
        var length = 0;
        while (length < text.Length && char.IsWhiteSpace(text[length]) && text[length] != '\r' && text[length] != '\n')
            length++;

        return text[..length];
    }

    private static bool ShouldIncreaseIndentAfter(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        return text.EndsWith(":", StringComparison.Ordinal) ||
               text.EndsWith("{", StringComparison.Ordinal) ||
               text.EndsWith("[", StringComparison.Ordinal) ||
               text.EndsWith("(", StringComparison.Ordinal) ||
               text.EndsWith("=>", StringComparison.Ordinal) ||
               text.EndsWith(" then", StringComparison.OrdinalIgnoreCase) ||
               text.EndsWith(" do", StringComparison.OrdinalIgnoreCase);
    }

    private string GetIndentUnit() => "\t";

    private static bool ShouldInsertStructuredBlock(string textBeforeCaret, string textAfterCaret)
    {
        var trimmedAfter = textAfterCaret.TrimStart();
        if (string.IsNullOrEmpty(trimmedAfter))
            return false;

        if (!BracketPairs.TryGetValue(textBeforeCaret.LastOrDefault(), out var closing))
            return false;

        return trimmedAfter.Length > 0 && trimmedAfter[0] == closing && closing is ')' or ']' or '}';
    }

    private static bool StartsWithClosingDelimiter(string text)
    {
        var trimmed = text.TrimStart();
        return trimmed.StartsWith("}", StringComparison.Ordinal) ||
               trimmed.StartsWith("]", StringComparison.Ordinal) ||
               trimmed.StartsWith(")", StringComparison.Ordinal);
    }

    private string RemoveOneIndentUnit(string indent)
    {
        var indentUnit = GetIndentUnit();
        if (indent.EndsWith(indentUnit, StringComparison.Ordinal))
            return indent[..^indentUnit.Length];

        return indent.Length > 0 ? indent[..^1] : indent;
    }

    private int GetOutdentLength(string lineText, int availableLength)
    {
        if (string.IsNullOrEmpty(lineText) || availableLength <= 0)
            return 0;

        var maxLength = Math.Min(availableLength, lineText.Length);
        var indentUnit = GetIndentUnit();
        if (maxLength >= indentUnit.Length &&
            lineText[..indentUnit.Length].Equals(indentUnit, StringComparison.Ordinal))
        {
            return indentUnit.Length;
        }

        var whitespaceCount = 0;
        while (whitespaceCount < maxLength && (lineText[whitespaceCount] == ' ' || lineText[whitespaceCount] == '\t'))
            whitespaceCount++;

        return whitespaceCount > 0 ? 1 : 0;
    }

    private static List<AvaloniaEdit.Document.DocumentLine> GetSelectedLines(
        AvaloniaEdit.Document.TextDocument doc,
        int startOffset,
        int endOffset)
    {
        if (endOffset > startOffset)
            endOffset--;

        var lines = new List<AvaloniaEdit.Document.DocumentLine>();
        var line = doc.GetLineByOffset(startOffset);
        while (line is not null)
        {
            lines.Add(line);
            if (line.EndOffset >= endOffset || line.NextLine is null)
                break;

            line = line.NextLine;
        }

        return lines;
    }

    private async void AutoSaveTimer_OnTick(object? sender, EventArgs e)
    {
        _autoSaveTimer.Stop();
        if (!IsAutoSaveEnabled || !HasFileOpen || !_isDirty) return;
        try
        {
            await SaveAsync(allowPromptForPath: false);
        }
        catch (Exception ex)
        {
            _autoSaveStatusMessage = BuildAutoSaveFailureMessage(ex);
            OnPropertyChanged(nameof(FileSummaryText));
            OnPropertyChanged(nameof(AutoSaveStatusText));
            await ShowWarningDialogAsync("Auto-save", ex);
        }
    }

    private void AutoSaveStatusTimer_OnTick(object? sender, EventArgs e)
    {
        _autoSaveStatusTimer.Stop();
        ClearAutoSaveStatus();
    }

    private bool _isConfirmedClose;

    private async void MainWindow_OnClosing(object? sender, WindowClosingEventArgs e)
    {
        // If we already confirmed through the dialog loop, let it through.
        if (_isConfirmedClose) return;

        var dirtyTabs = OpenTabs.Where(t => t.IsDirty).ToList();
        if (dirtyTabs.Count == 0 || !IsConfirmBeforeClosingUnsavedTabsEnabled) return;

        // Cancel the close and handle it ourselves asynchronously.
        e.Cancel = true;

        foreach (var tab in dirtyTabs)
        {
            var action = await ShowUnsavedTabDialogAsync(tab);
            switch (action)
            {
                case UnsavedTabAction.Cancel:
                    return; // User aborted — leave the window open.

                case UnsavedTabAction.Save:
                    ActivateTab(tab, focusEditor: false);
                    if (!await SaveAsync(allowPromptForPath: true, forcePromptForPath: false))
                        return; // Save was cancelled — leave the window open.
                    break;

                // UnsavedTabAction.Discard — just continue to the next tab.
            }
        }

        // All dirty tabs resolved — close for real.
        _isConfirmedClose = true;
        Close();
    }

    private void MainWindow_OnClosed(object? sender, EventArgs e)
    {
        _autoSaveTimer.Stop();
        _autoSaveStatusTimer.Stop();
        _discordReconnectTimer.Stop();
        _extensionsRefreshDebounceTimer.Stop();
        DisposeExtensionFolderWatchers();
        DisposeDiscordPresence();
        CurrentImagePreview = null;
    }

    private async void MainWindow_OnKeyDown(object? sender, KeyEventArgs e)
    {
        var hasControl = (e.KeyModifiers & KeyModifiers.Control) == KeyModifiers.Control;
        var hasShift   = (e.KeyModifiers & KeyModifiers.Shift)   == KeyModifiers.Shift;

        // Escape — dismiss Settings / Extensions / WhatsNew and return to editor
        if (e.Key == Key.Escape && !hasControl)
        {
            if (IsSettingsPageVisible || IsExtensionsPageVisible || IsWhatsNewPageVisible)
            {
                IsSettingsPageVisible  = false;
                IsExtensionsPageVisible = false;
                IsWhatsNewPageVisible  = false;
                FocusEditor();
                e.Handled = true;
            }
            return;
        }

        if (!hasControl) return;

        switch (e.Key)
        {
            case Key.N:
                NewFile();
                e.Handled = true;
                break;

            case Key.E:
                if (hasShift)
                {
                    // Ctrl+Shift+E — back to editor (original behaviour)
                    IsSettingsPageVisible  = false;
                    IsExtensionsPageVisible = false;
                    FocusEditor();
                    e.Handled = true;
                }
                else
                {
                    // Ctrl+E — open Extensions
                    IsSettingsPageVisible  = false;
                    IsWhatsNewPageVisible  = false;
                    IsHomePageVisible      = false;
                    IsExtensionsPageVisible = true;
                    _ = RefreshExtensionsDataAsync();
                    e.Handled = true;
                }
                break;

            case Key.OemComma:
                // Ctrl+, — open Settings
                IsExtensionsPageVisible = false;
                IsWhatsNewPageVisible   = false;
                IsHomePageVisible       = false;
                IsSettingsPageVisible   = true;
                e.Handled = true;
                break;

            case Key.H:
                // Ctrl+H — go to Home
                IsSettingsPageVisible   = false;
                IsExtensionsPageVisible = false;
                IsWhatsNewPageVisible   = false;
                IsHomePageVisible       = true;
                e.Handled = true;
                break;

            case Key.S:
                if (hasShift)
                {
                    // Ctrl+Shift+S — Save As (always prompt for path)
                    e.Handled = true;
                    await SaveAsAsync();
                }
                else
                {
                    // Ctrl+S — Save
                    e.Handled = true;
                    await SaveAsync();
                }
                break;

            case Key.O:
                e.Handled = true;
                await OpenFileAsync();
                break;

            case Key.K:
                // Ctrl+K — toggle folder open/close
                e.Handled = true;
                if (IsFolderOpen)
                    CloseFolder();
                else
                    await OpenFolderAsync();
                break;

            case Key.B:
                // Ctrl+B — toggle file explorer sidebar
                IsFileExplorerVisible = !IsFileExplorerVisible;
                e.Handled = true;
                break;

            case Key.W:
                // Ctrl+W — close current tab
                if (ActiveEditorTab is not null)
                    await RequestCloseTabAsync(ActiveEditorTab);
                e.Handled = true;
                break;

            case Key.F:
                // Ctrl+F — toggle find panel
                if (CanShowFindInFile)
                    IsFindPanelVisible = !IsFindPanelVisible;
                else
                    IsFindPanelVisible = false;
                e.Handled = true;
                break;

            case Key.X when hasShift:
                // Ctrl+Shift+X — open Extensions (secondary binding)
                IsSettingsPageVisible   = false;
                IsWhatsNewPageVisible   = false;
                IsHomePageVisible       = false;
                IsExtensionsPageVisible = true;
                _ = RefreshExtensionsDataAsync();
                e.Handled = true;
                break;
        }
    }

    // ── Nested types ─────────────────────────────────────────────────────────

    private static int NormalizeTabSize(int value) => value is 2 or 4 or 8 ? value : 4;

    private async Task<UnsavedTabAction> ShowUnsavedTabDialogAsync(EditorTab tab)
    {
        var result = UnsavedTabAction.Cancel;
        Window? dialog = null;
        dialog = new Window
        {
            Width = 420,
            Height = 190,
            CanResize = false,
            ShowInTaskbar = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Title = "Unsaved Changes",
            Background = CardBrush,
            Content = BuildUnsavedTabDialogContent(
                tab,
                () => { result = UnsavedTabAction.Save; dialog!.Close(); },
                () => { result = UnsavedTabAction.Discard; dialog!.Close(); },
                () => { result = UnsavedTabAction.Cancel; dialog!.Close(); })
        };

        await dialog.ShowDialog(this);
        return result;
    }

    private Control BuildUnsavedTabDialogContent(
        EditorTab tab,
        Action saveAction,
        Action discardAction,
        Action cancelAction)
    {
        var buttonRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing = 10,
            HorizontalAlignment = HorizontalAlignment.Center,
            Children =
            {
                CreateDialogButton("Cancel", ButtonBrush, SurfaceBorderBrush, PrimaryTextBrush, cancelAction),
                CreateDialogButton("Discard", ButtonHoverBrush, SurfaceBorderBrush, PrimaryTextBrush, discardAction),
                CreateDialogButton("Save", AccentBrush, AccentBrush, Brushes.White, saveAction)
            }
        };

        return new Border
        {
            Padding = new Thickness(20),
            Child = new StackPanel
            {
                Spacing = 14,
                Children =
                {
                    new TextBlock
                    {
                        Text = "Save changes before closing?",
                        FontSize = 18,
                        FontWeight = FontWeight.SemiBold,
                        Foreground = PrimaryTextBrush
                    },
                    new TextBlock
                    {
                        Text = $"{tab.DisplayName} has unsaved changes.",
                        Foreground = MutedTextBrush,
                        TextWrapping = TextWrapping.Wrap
                    },
                    new TextBlock
                    {
                        Text = "Choose Save to keep them, Discard to close without saving, or Cancel to keep editing.",
                        Foreground = MutedTextBrush,
                        TextWrapping = TextWrapping.Wrap
                    },
                    buttonRow
                }
            }
        };
    }

    private static Button CreateDialogButton(
        string text,
        IBrush background,
        IBrush borderBrush,
        IBrush foreground,
        Action clickAction)
    {
        var button = new Button
        {
            Content = text,
            Background = background,
            BorderBrush = borderBrush,
            BorderThickness = new Thickness(1),
            Foreground = foreground,
            Padding = new Thickness(14, 8),
            CornerRadius = new CornerRadius(8),
            MinWidth = 86,
            HorizontalContentAlignment = HorizontalAlignment.Center
        };

        button.Click += (_, _) => clickAction();
        return button;
    }

    // Shows a non-fatal warning dialog that mirrors the crash dialog in App.axaml.cs
    // but uses softer wording. Call this from any recoverable error path where the
    // user needs to know something went wrong.
    private async Task ShowWarningDialogAsync(string context, Exception exception)
    {
        try
        {
            var titleText = new TextBlock
            {
                Text = "Something went wrong",
                FontSize = 16,
                FontWeight = FontWeight.SemiBold,
                Foreground = PrimaryTextBrush,
                TextWrapping = TextWrapping.Wrap,
            };

            var subtitleText = new TextBlock
            {
                Text = "Kodo ran into a problem with this operation. No data was lost — you can try again.",
                FontSize = 13,
                Foreground = MutedTextBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 4, 0, 0),
            };

            // Context badge (e.g. "File save", "Extension install")
            var contextBadge = new Border
            {
                Background = ButtonBrush,
                BorderBrush = SurfaceBorderBrush,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(10, 5),
                HorizontalAlignment = HorizontalAlignment.Left,
                Child = new TextBlock
                {
                    Text = context,
                    FontSize = 12,
                    FontFamily = new FontFamily("Cascadia Code,Consolas,Menlo,monospace"),
                    Foreground = new SolidColorBrush(Color.Parse("#9CDCFE")),
                },
            };

            // Scrollable, selectable exception detail so the user can copy it if needed
            var exceptionText = new SelectableTextBlock
            {
                Text = exception.ToString(),
                FontSize = 12,
                FontFamily = new FontFamily("Cascadia Code,Consolas,Menlo,monospace"),
                Foreground = new SolidColorBrush(Color.Parse("#CE9178")),
                TextWrapping = TextWrapping.Wrap,
            };

            var exceptionScroll = new ScrollViewer
            {
                Content = exceptionText,
                MaxHeight = 220,
                VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            };

            var exceptionBorder = new Border
            {
                Background = new SolidColorBrush(Color.Parse("#1A1A1A")),
                BorderBrush = SurfaceBorderBrush,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(12),
                Child = exceptionScroll,
            };

            var dismissButton = new Button
            {
                Content = "Dismiss",
                HorizontalAlignment = HorizontalAlignment.Right,
                Padding = new Thickness(20, 8),
                Background = AccentBrush,
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                CornerRadius = new CornerRadius(8),
            };

            var content = new StackPanel
            {
                Spacing = 12,
                Margin = new Thickness(20),
                Children =
                {
                    titleText,
                    subtitleText,
                    contextBadge,
                    exceptionBorder,
                    dismissButton,
                },
            };

            Window? dialog = null;
            dialog = new Window
            {
                Title = "Kodo — Error",
                Width = 520,
                SizeToContent = SizeToContent.Height,
                MinWidth = 380,
                MinHeight = 160,
                MaxHeight = 640,
                CanResize = true,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Background = CardBrush,
                Content = content,
            };

            dismissButton.Click += (_, _) => dialog!.Close();
            await dialog.ShowDialog(this);
        }
        catch
        {
            // The warning dialog itself must never crash the app.
        }
    }

    private sealed class AppSettings
    {
        public string ThemeName { get; set; } = "Dark";
        public bool AutoSaveEnabled { get; set; }
        public bool DiscordRichPresenceEnabled { get; set; }
        public bool StatusBarFilePathVisible { get; set; } = true;
        public bool WordWrapEnabled { get; set; }
        public int TabSize { get; set; } = 4;
        public int EditorFontSize { get; set; } = 14;
        public bool ConfirmBeforeClosingUnsavedTabsEnabled { get; set; } = true;
        public bool RestoreOpenTabsOnLaunchEnabled { get; set; }
        public List<string> OpenTabPaths { get; set; } = [];
        public string? ActiveTabPath { get; set; }
        public List<RecentFileEntry> RecentFiles { get; set; } = [];
    }

    public sealed class RecentFileEntry
    {
        public string   Path       { get; set; } = string.Empty;
        public bool     IsFolder   { get; set; }
        public DateTime LastOpened { get; set; } = DateTime.Now;
    }

    private enum UnsavedTabAction
    {
        Save,
        Discard,
        Cancel
    }
}

public sealed class RecentFileItem
{
    public RecentFileItem(string path, bool isFolder, DateTime lastOpened)
    {
        Path        = path;
        IsFolder    = isFolder;
        LastOpened  = lastOpened;
    }

    public string   Path       { get; }
    public bool     IsFolder   { get; }
    public DateTime LastOpened { get; }

    public string DisplayName
    {
        get
        {
            if (IsFolder)
                return System.IO.Path.GetFileName(Path.TrimEnd(System.IO.Path.DirectorySeparatorChar));
            var name = System.IO.Path.GetFileName(Path);
            var dot = name.IndexOf('.');
            return dot > 0 ? name[..dot] : name;
        }
    }

    public string DirectoryPath => IsFolder
        ? System.IO.Path.GetDirectoryName(Path.TrimEnd(System.IO.Path.DirectorySeparatorChar)) ?? string.Empty
        : System.IO.Path.GetDirectoryName(Path) ?? string.Empty;

    public string FileTypeName
    {
        get
        {
            if (IsFolder) return "Folder";
            var ext = System.IO.Path.GetExtension(Path);
            if (string.IsNullOrEmpty(ext))
            {
                var name = System.IO.Path.GetFileName(Path);
                return string.IsNullOrWhiteSpace(name) ? "File" : $"{name} file";
            }
            return $"{ext.ToLowerInvariant()} file";
        }
    }

    public string LastOpenedText
    {
        get
        {
            var diff = DateTime.Now - LastOpened;
            if (diff.TotalMinutes < 1)  return "Just now";
            if (diff.TotalHours   < 1)  return $"{(int)diff.TotalMinutes}m ago";
            if (diff.TotalDays    < 1)  return $"{(int)diff.TotalHours}h ago";
            if (diff.TotalDays    < 30) return $"{(int)diff.TotalDays}d ago";
            return LastOpened.ToString("MMM d");
        }
    }
}

public sealed class RainbowBracketColorizer : DocumentColorizingTransformer
{
    private static readonly Dictionary<char, char> OpeningToClosing = new()
    {
        ['('] = ')',
        ['['] = ']',
        ['{'] = '}'
    };

    private static readonly Dictionary<char, char> ClosingToOpening = new()
    {
        [')'] = '(',
        [']'] = '[',
        ['}'] = '{'
    };

    private static readonly IBrush[] RainbowBrushes =
    [
        Brush.Parse("#FFD700"),
        Brush.Parse("#DA70D6"),
        Brush.Parse("#4FC1FF"),
        Brush.Parse("#C586C0"),
        Brush.Parse("#9CDCFE"),
        Brush.Parse("#D7BA7D")
    ];
    private static readonly MethodInfo? SetTextRunPropertiesMethod =
        typeof(VisualLineElement).GetMethod("SetTextRunProperties", BindingFlags.Instance | BindingFlags.NonPublic);

    private ParseSnapshot? _snapshot;
    private string _commentLine = "//";
    private string _commentBlockStart = "/*";
    private string _commentBlockEnd = "*/";
    private string[] _stringDelimiters = ["\"", "'"];
    private string[] _multiLineStringDelimiters = [];

    public void UpdateSyntax(LoadedExtension? extension)
    {
        if (extension is null)
        {
            _commentLine = "//";
            _commentBlockStart = "/*";
            _commentBlockEnd = "*/";
            _stringDelimiters = ["\"", "'"];
            _multiLineStringDelimiters = [];
        }
        else
        {
            _commentLine = extension.CommentLine;
            _commentBlockStart = extension.CommentBlockStart;
            _commentBlockEnd = extension.CommentBlockEnd;
            _stringDelimiters = extension.StringDelimiters
                .Where(d => !string.IsNullOrEmpty(d))
                .Distinct()
                .ToArray();
            _multiLineStringDelimiters = extension.MultiLineStringDelimiters
                .Where(d => !string.IsNullOrEmpty(d))
                .Distinct()
                .ToArray();
        }

        InvalidateCache();
    }

    public void InvalidateCache() => _snapshot = null;

    protected override void ColorizeLine(AvaloniaEdit.Document.DocumentLine line)
    {
        var document = CurrentContext.Document;
        if (document is null || line.Length <= 0)
            return;

        var snapshot = EnsureSnapshot(document.Text ?? string.Empty);
        var lineState = snapshot.GetLineState(line.LineNumber);
        var text = document.GetText(line.Offset, line.Length);
        var stack = new Stack<char>(lineState.BracketStack.Reverse());
        var mode = lineState.Mode;
        var activeDelimiter = lineState.Delimiter;

        for (var index = 0; index < text.Length; index++)
        {
            var nextMode = mode;
            var nextDelimiter = activeDelimiter;

            if (mode == ScanMode.LineComment)
                break;

            if (TryConsumeLineBreak(text, ref index, ref nextMode, ref nextDelimiter))
            {
                mode = nextMode;
                activeDelimiter = nextDelimiter;
                continue;
            }

            if (mode == ScanMode.BlockComment)
            {
                if (!string.IsNullOrEmpty(activeDelimiter) && MatchesAt(text, index, activeDelimiter))
                {
                    index += activeDelimiter.Length - 1;
                    mode = ScanMode.Normal;
                    activeDelimiter = null;
                }

                continue;
            }

            if (mode == ScanMode.String)
            {
                if (!string.IsNullOrEmpty(activeDelimiter) &&
                    MatchesAt(text, index, activeDelimiter) &&
                    !IsEscaped(text, index))
                {
                    index += activeDelimiter.Length - 1;
                    mode = ScanMode.Normal;
                    activeDelimiter = null;
                }

                continue;
            }

            if (mode == ScanMode.MultiLineString)
            {
                if (!string.IsNullOrEmpty(activeDelimiter) &&
                    MatchesAt(text, index, activeDelimiter) &&
                    !IsEscaped(text, index))
                {
                    index += activeDelimiter.Length - 1;
                    mode = ScanMode.Normal;
                    activeDelimiter = null;
                }

                continue;
            }

            if (!string.IsNullOrEmpty(_commentLine) && MatchesAt(text, index, _commentLine))
            {
                mode = ScanMode.LineComment;
                index += _commentLine.Length - 1;
                continue;
            }

            if (!string.IsNullOrEmpty(_commentBlockStart) && MatchesAt(text, index, _commentBlockStart))
            {
                mode = ScanMode.BlockComment;
                activeDelimiter = _commentBlockEnd;
                index += _commentBlockStart.Length - 1;
                continue;
            }

            var multiLineDelimiter = MatchDelimiter(text, index, _multiLineStringDelimiters);
            if (multiLineDelimiter is not null)
            {
                mode = ScanMode.MultiLineString;
                activeDelimiter = multiLineDelimiter;
                index += multiLineDelimiter.Length - 1;
                continue;
            }

            var stringDelimiter = MatchDelimiter(text, index, _stringDelimiters);
            if (stringDelimiter is not null)
            {
                mode = ScanMode.String;
                activeDelimiter = stringDelimiter;
                index += stringDelimiter.Length - 1;
                continue;
            }

            var ch = text[index];
            var offset = line.Offset + index;

            if (OpeningToClosing.ContainsKey(ch))
            {
                ApplyBracketColor(offset, GetBrushForDepth(stack.Count));
                stack.Push(ch);
                continue;
            }

            if (ClosingToOpening.TryGetValue(ch, out var opening) &&
                stack.Count > 0 &&
                stack.Peek() == opening)
            {
                ApplyBracketColor(offset, GetBrushForDepth(stack.Count - 1));
                stack.Pop();
            }
        }
    }

    private void ApplyBracketColor(int offset, IBrush brush)
    {
        ChangeLinePart(offset, offset + 1, element =>
        {
            var properties = element.TextRunProperties.Clone();
            properties.SetForegroundBrush(brush);
            SetTextRunPropertiesMethod?.Invoke(element, [properties]);
        });
    }

    private IBrush GetBrushForDepth(int depth) => RainbowBrushes[Math.Abs(depth) % RainbowBrushes.Length];

    private ParseSnapshot EnsureSnapshot(string text)
    {
        if (_snapshot is { } snapshot && string.Equals(snapshot.Text, text, StringComparison.Ordinal))
            return snapshot;

        _snapshot = BuildSnapshot(text);
        return _snapshot;
    }

    private ParseSnapshot BuildSnapshot(string text)
    {
        var lineStates = new List<LineState> { new(string.Empty, ScanMode.Normal, null) };
        var stack = new Stack<char>();
        var mode = ScanMode.Normal;
        string? activeDelimiter = null;

        for (var index = 0; index < text.Length; index++)
        {
            if (TryConsumeLineBreak(text, ref index, ref mode, ref activeDelimiter))
            {
                lineStates.Add(new(new string(stack.Reverse().ToArray()), mode, activeDelimiter));
                continue;
            }

            if (mode == ScanMode.LineComment)
                continue;

            if (mode == ScanMode.BlockComment)
            {
                if (!string.IsNullOrEmpty(activeDelimiter) && MatchesAt(text, index, activeDelimiter))
                    ExitDelimitedState(ref index, ref mode, ref activeDelimiter);

                continue;
            }

            if (mode == ScanMode.String)
            {
                if (!string.IsNullOrEmpty(activeDelimiter) && MatchesAt(text, index, activeDelimiter) && !IsEscaped(text, index))
                    ExitDelimitedState(ref index, ref mode, ref activeDelimiter);

                continue;
            }

            if (mode == ScanMode.MultiLineString)
            {
                if (!string.IsNullOrEmpty(activeDelimiter) && MatchesAt(text, index, activeDelimiter) && !IsEscaped(text, index))
                    ExitDelimitedState(ref index, ref mode, ref activeDelimiter);

                continue;
            }

            if (!string.IsNullOrEmpty(_commentLine) && MatchesAt(text, index, _commentLine))
            {
                mode = ScanMode.LineComment;
                index += _commentLine.Length - 1;
                continue;
            }

            if (!string.IsNullOrEmpty(_commentBlockStart) && MatchesAt(text, index, _commentBlockStart))
            {
                mode = ScanMode.BlockComment;
                activeDelimiter = _commentBlockEnd;
                index += _commentBlockStart.Length - 1;
                continue;
            }

            var multiLineDelimiter = MatchDelimiter(text, index, _multiLineStringDelimiters);
            if (multiLineDelimiter is not null)
            {
                mode = ScanMode.MultiLineString;
                activeDelimiter = multiLineDelimiter;
                index += multiLineDelimiter.Length - 1;
                continue;
            }

            var stringDelimiter = MatchDelimiter(text, index, _stringDelimiters);
            if (stringDelimiter is not null)
            {
                mode = ScanMode.String;
                activeDelimiter = stringDelimiter;
                index += stringDelimiter.Length - 1;
                continue;
            }

            var ch = text[index];
            if (OpeningToClosing.ContainsKey(ch))
            {
                stack.Push(ch);
            }
            else if (ClosingToOpening.TryGetValue(ch, out var opening) &&
                     stack.Count > 0 &&
                     stack.Peek() == opening)
            {
                stack.Pop();
            }
        }

        return new ParseSnapshot(text, lineStates);
    }

    private static void ExitDelimitedState(ref int index, ref ScanMode mode, ref string? activeDelimiter)
    {
        index += activeDelimiter!.Length - 1;
        mode = ScanMode.Normal;
        activeDelimiter = null;
    }

    private static bool TryConsumeLineBreak(string text, ref int index, ref ScanMode mode, ref string? activeDelimiter)
    {
        if (text[index] == '\r')
        {
            if (index + 1 < text.Length && text[index + 1] == '\n')
                index++;

            if (mode is ScanMode.LineComment or ScanMode.String)
            {
                mode = ScanMode.Normal;
                activeDelimiter = null;
            }

            return true;
        }

        if (text[index] == '\n')
        {
            if (mode is ScanMode.LineComment or ScanMode.String)
            {
                mode = ScanMode.Normal;
                activeDelimiter = null;
            }

            return true;
        }

        return false;
    }

    private static bool MatchesAt(string text, int index, string token) =>
        index + token.Length <= text.Length &&
        string.CompareOrdinal(text, index, token, 0, token.Length) == 0;

    private static string? MatchDelimiter(string text, int index, IEnumerable<string> delimiters) =>
        delimiters.FirstOrDefault(delimiter => MatchesAt(text, index, delimiter));

    private static bool IsEscaped(string text, int index)
    {
        var slashCount = 0;
        for (var i = index - 1; i >= 0 && text[i] == '\\'; i--)
            slashCount++;

        return slashCount % 2 == 1;
    }

    private sealed record ParseSnapshot(string Text, List<LineState> LineStates)
    {
        public LineState GetLineState(int lineNumber) =>
            lineNumber > 0 && lineNumber <= LineStates.Count
                ? LineStates[lineNumber - 1]
                : new LineState(string.Empty, ScanMode.Normal, null);
    }

    private readonly record struct LineState(string BracketStack, ScanMode Mode, string? Delimiter);

    private enum ScanMode
    {
        Normal,
        LineComment,
        BlockComment,
        String,
        MultiLineString
    }
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
        var commentColor     = ColorFor(ext, "comment",      "#6A9955");
        var stringColor      = ColorFor(ext, "string",       "#CE9178");
        var charLiteralColor = ColorFor(ext, "charLiteral",  "#CE9178");
        var keywordColor     = ColorFor(ext, "keyword",      "#569CD6");
        var typeColor        = ColorFor(ext, "type",         "#4EC9B0");
        var numberColor      = ColorFor(ext, "number",       "#B5CEA8");
        var functionColor    = ColorFor(ext, "function",     "#DCDCAA");
        var namespaceColor   = ColorFor(ext, "namespace",    "#4FC1FF");
        var propertyColor    = ColorFor(ext, "property",     "#9CDCFE");
        var attributeColor   = ColorFor(ext, "attribute",    "#C586C0");
        var operatorColor    = ColorFor(ext, "operator",     "#D4D4D4");
        var punctuationColor = ColorFor(ext, "punctuation",  "#D4D4D4");
        var preprocessorColor = ColorFor(ext, "preprocessor", "#C586C0");
        var variableColor    = ColorFor(ext, "variable",      "#A0DBFD");

        // ── Inner rulesets ────────────────────────────────────────────────────────────
        // codeRuleSet  — holds keyword/type/number rules; used as the inner ruleset of
        //                spans that should still syntax-colour their contents.
        //                (Currently unused as inner ruleset, but kept for clarity.)
        // emptyRuleSet — no rules; used as the inner ruleset of comment and string spans
        //                so that keyword/number rules cannot fire inside them.
        //                AvaloniaEdit applies SpanColor to the entire span body, so the
        //                correct span colour is provided by the outer SpanColor property.
        //                The inner ruleset only needs to be empty — no DefaultColor needed.
        var codeRuleSet  = new HighlightingRuleSet();
        var emptyRuleSet = new HighlightingRuleSet();

        if (ext.Keywords.Length > 0)
        {
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = BuildTokenRegex(ext.Keywords),
                Color = keywordColor
            });
        }

        if (ext.Types.Length > 0)
        {
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = BuildTokenRegex(ext.Types),
                Color = typeColor
            });
        }

        codeRuleSet.Rules.Add(new HighlightingRule
        {
            Regex = new Regex(@"(?<![\p{L}\p{Nd}_])(?:0[xX][0-9A-Fa-f]+|0[bB][01]+|0[oO][0-7]+|\d+(?:\.\d+)?(?:[eE][+\-]?\d+)?)(?![\p{L}\p{Nd}_])", RegexOptions.Compiled),
            Color = numberColor
        });

        codeRuleSet.Rules.Add(new HighlightingRule
        {
            Regex = new Regex(@"(?<=\.|->|::)[\p{L}_][\p{L}\p{Nd}_:-]*", RegexOptions.Compiled),
            Color = propertyColor
        });

        codeRuleSet.Rules.Add(new HighlightingRule
        {
            Regex = new Regex(@"(?<![\p{L}\p{Nd}_])[\p{L}_][\p{L}\p{Nd}_]*(?=\.)", RegexOptions.Compiled),
            Color = namespaceColor
        });

        codeRuleSet.Rules.Add(new HighlightingRule
        {
            Regex = new Regex(@"(?<![\p{L}\p{Nd}_])[@#][\p{L}_][\p{L}\p{Nd}_-]*", RegexOptions.Compiled),
            Color = preprocessorColor
        });

        codeRuleSet.Rules.Add(new HighlightingRule
        {
            Regex = new Regex(@"(?<=\[)[\p{L}_][\p{L}\p{Nd}_:.]*(?=[,\]\(])|(?<=<)[\p{L}_][\p{L}\p{Nd}_:-]*(?=[^>]*>)", RegexOptions.Compiled),
            Color = attributeColor
        });

        codeRuleSet.Rules.Add(new HighlightingRule
        {
            Regex = new Regex(@"(?<![\p{L}\p{Nd}_])[\p{L}_][\p{L}\p{Nd}_]*(?=\s*\()", RegexOptions.Compiled),
            Color = functionColor
        });

        codeRuleSet.Rules.Add(new HighlightingRule
        {
            Regex = new Regex(@"=>|->|::|\+\+|--|\+=|-=|\*=|/=|%=|&&|\|\||<<|>>|<=|>=|==|!=|=|\+|-|\*|/|%|!|\?|:|<|>|&|\||\^|~", RegexOptions.Compiled),
            Color = operatorColor
        });

        codeRuleSet.Rules.Add(new HighlightingRule
        {
            Regex = new Regex(@"[{}\[\]();,.]", RegexOptions.Compiled),
            Color = punctuationColor
        });

        // Last segment of an import/using directive for any language.
        // Covers: "using System;", "import os", "import numpy as np",
        //         "#include <vector>", "require 'json'", "use std::io", etc.
        // The final identifier before ; or end-of-line — after a chain of
        // dotted/slashed segments — is coloured as a namespace.
        // This fires for C#, Python, Java, Rust, Ruby, Go, JS/TS, C/C++, and more.
        codeRuleSet.Rules.Add(new HighlightingRule
        {
            Regex = new Regex(
                @"(?<=\b(?:using|import|include|require|use|from)\b\s+(?:[\p{L}_][\p{L}\p{Nd}_./\\]*\s*[./\\]\s*)?)[\p{L}_][\p{L}\p{Nd}_]*(?=\s*(?:;|$))",
                RegexOptions.Compiled),
            Color = namespaceColor
        });

        // User-defined variables — bare identifiers that are not preceded by a dot
        // (property), not followed by '(' (function call) or '.' (namespace segment),
        // and were not already claimed by keyword/type/number rules above.
        codeRuleSet.Rules.Add(new HighlightingRule
        {
            Regex = new Regex(@"(?<![.\p{L}\p{Nd}_])[\p{L}_][\p{L}\p{Nd}_]*(?!\s*[\.(]|[\p{L}\p{Nd}_])", RegexOptions.Compiled),
            Color = variableColor
        });

        // Char literal rule (C#-style languages with disableSingleQuoteStrings).
        // Lives in codeRuleSet so that its rules can be copied to mainRuleSet;
        // it will not fire inside comment or string spans (which use emptyRuleSet).
        if (ext.DisableSingleQuoteStrings)
        {
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(
                    @"'(?:\\(?:u[0-9A-Fa-f]{4}|U[0-9A-Fa-f]{8}|x[0-9A-Fa-f]{1,4}|[0-7]{1,3}|[abfnrtv\\""'0])|[^\\'])'",
                    RegexOptions.Compiled),
                Color = charLiteralColor
            });
        }

        // ── Main ruleset ──────────────────────────────────────────────────────────────
        // Spans are checked in order — first match wins. SpanColor is applied by
        // AvaloniaEdit to the entire span (start delimiter + body + end delimiter),
        // modulated by SpanColorIncludesStart / SpanColorIncludesEnd for the delimiters.
        // The inner emptyRuleSet ensures no keyword/number rules fire inside the span.
        var mainRuleSet = new HighlightingRuleSet();

        // Block comment /* … */ — added first so it takes priority over // on the same line.
        if (!string.IsNullOrEmpty(ext.CommentBlockStart) && !string.IsNullOrEmpty(ext.CommentBlockEnd))
        {
            mainRuleSet.Spans.Add(new HighlightingSpan
            {
                StartExpression        = new Regex(Regex.Escape(ext.CommentBlockStart), RegexOptions.Compiled),
                EndExpression          = new Regex(Regex.Escape(ext.CommentBlockEnd),   RegexOptions.Compiled),
                SpanColor              = commentColor,
                SpanColorIncludesStart = true,
                SpanColorIncludesEnd   = true,
                RuleSet                = emptyRuleSet
            });
        }

        // Single-line comment // … end-of-line.
        // Use $ as the explicit end-of-line anchor so the whole remainder of the
        // line is coloured with the extension's configured comment colour.
        if (!string.IsNullOrEmpty(ext.CommentLine))
        {
            mainRuleSet.Spans.Add(new HighlightingSpan
            {
                StartExpression        = new Regex(Regex.Escape(ext.CommentLine), RegexOptions.Compiled),
                EndExpression          = new Regex("$", RegexOptions.Compiled),
                SpanColor              = commentColor,
                SpanColorIncludesStart = true,
                SpanColorIncludesEnd   = false,
                RuleSet                = emptyRuleSet
            });
        }

        foreach (var delimiter in ext.MultiLineStringDelimiters.Where(d => !string.IsNullOrEmpty(d)).Distinct())
        {
            mainRuleSet.Spans.Add(CreateStringSpan(delimiter, stringColor, emptyRuleSet, allowEndOfLineFallback: false));
        }

        var stringDelimiters = ext.StringDelimiters
            .Where(d => !string.IsNullOrEmpty(d))
            .Distinct()
            .ToList();

        if (ext.DisableSingleQuoteStrings)
            stringDelimiters.RemoveAll(d => d == "'");

        foreach (var delimiter in stringDelimiters)
            mainRuleSet.Spans.Add(CreateStringSpan(delimiter, stringColor, emptyRuleSet, allowEndOfLineFallback: true));

        // Copy keyword/type/number/char-literal rules from codeRuleSet into mainRuleSet.
        // When the engine enters a comment or string span it switches to emptyRuleSet,
        // so these rules are automatically suppressed inside those regions.
        foreach (var rule in codeRuleSet.Rules)
            mainRuleSet.Rules.Add(rule);

        return mainRuleSet;
    }

    private static Regex BuildTokenRegex(IEnumerable<string> tokens)
    {
        var distinctTokens = tokens
            .Where(token => !string.IsNullOrWhiteSpace(token))
            .Distinct()
            .OrderByDescending(token => token.Length)
            .Select(Regex.Escape)
            .ToArray();

        return new Regex(
            @"(?<![\p{L}\p{Nd}_])(" + string.Join("|", distinctTokens) + @")(?![\p{L}\p{Nd}_])",
            RegexOptions.Compiled);
    }

    private static HighlightingSpan CreateStringSpan(
        string delimiter,
        HighlightingColor stringColor,
        HighlightingRuleSet emptyRuleSet,
        bool allowEndOfLineFallback)
    {
        var escapedDelimiter = Regex.Escape(delimiter);
        var endPattern = allowEndOfLineFallback
            ? $@"(?<!\\){escapedDelimiter}|$"
            : $@"(?<!\\){escapedDelimiter}";

        return new HighlightingSpan
        {
            StartExpression = new Regex(escapedDelimiter, RegexOptions.Compiled),
            EndExpression = new Regex(endPattern, RegexOptions.Compiled),
            SpanColor = stringColor,
            SpanColorIncludesStart = true,
            SpanColorIncludesEnd = true,
            RuleSet = emptyRuleSet
        };
    }

    public HighlightingColor GetNamedColor(string name) => new();
    // Named ruleset lookups are only used when span definitions reference an inner
    // ruleset by name. This definition uses only anonymous inline rulesets, so no
    // name will ever be looked up — return an empty set for all names.
    public HighlightingRuleSet GetNamedRuleSet(string name) => new HighlightingRuleSet();
}

public sealed class IndentGuideBackgroundRenderer : IBackgroundRenderer
{
    public KnownLayer Layer => KnownLayer.Background;

    public int TabSize { get; set; } = 4;

    public IBrush GuideBrush { get; set; } = new SolidColorBrush(Color.Parse("#808080"), 0.4);

    public void Draw(TextView textView, DrawingContext drawingContext)
    {
        if (!textView.VisualLinesValid || TabSize <= 0)
            return;

        var spaceWidth = textView.WideSpaceWidth;
        if (spaceWidth <= 0)
            return;

        var document = textView.Document;
        if (document is null || document.LineCount == 0)
            return;

        if (!textView.VisualLines.Any())
            return;

        // Pre-compute indent depth (in tab-stop levels) for every document line.
        var totalLines = document.LineCount;
        var lineDepths = new int[totalLines + 1]; // 1-based index

        for (var i = 1; i <= totalLines; i++)
        {
            var docLine = document.GetLineByNumber(i);
            var text    = document.GetText(docLine);
            lineDepths[i] = string.IsNullOrWhiteSpace(text) ? -1 : GetIndentColumns(text) / TabSize;
        }

        // Fill blank lines from surrounding context so guides are continuous
        for (var i = 1; i <= totalLines; i++)
        {
            if (lineDepths[i] != -1) continue;
            var above = 0;
            for (var a = i - 1; a >= 1; a--)
                if (lineDepths[a] >= 0) { above = lineDepths[a]; break; }
            var below = 0;
            for (var b = i + 1; b <= totalLines; b++)
                if (lineDepths[b] >= 0) { below = lineDepths[b]; break; }
            lineDepths[i] = Math.Min(above, below);
        }

        // Measure true text-start X from first visible line.
        // AvaloniaEdit's DrawingContext for background renderers is NOT pre-scrolled,
        // so we subtract ScrollOffset for both X and Y.
        var scrollX = textView.ScrollOffset.X;
        var scrollY = textView.ScrollOffset.Y;

        var refLine = textView.VisualLines[0].FirstDocumentLine;
        var originX = textView.GetVisualPosition(
            new AvaloniaEdit.TextViewPosition(refLine.LineNumber, 1),
            VisualYPosition.LineTop).X - scrollX;

        // Dashed pen matching VS Code-style indent guides
        var dashStyle = new DashStyle([2, 2], 0);
        var pen = new Pen(GuideBrush, 1, dashStyle);

        foreach (var visualLine in textView.VisualLines)
        {
            var lineNumber = visualLine.FirstDocumentLine.LineNumber;
            var depth = lineDepths[lineNumber];
            if (depth <= 0) continue;

            // VisualTop is document-absolute; subtract scrollY for screen coords
            var top    = visualLine.VisualTop - scrollY;
            var bottom = top + visualLine.Height;

            for (var level = 1; level <= depth; level++)
            {
                // Each guide sits at the first character of that indent level:
                // level 1 → column TabSize, level 2 → column 2*TabSize, etc.
                // originX is the pixel X of column 1, so offset by (level*TabSize - 1) chars.
                var x = originX + (level * TabSize - 1) * spaceWidth;
                if (x < 0 || x > textView.Bounds.Width) continue;

                drawingContext.DrawLine(pen, new Point(x, top), new Point(x, bottom));
            }
        }
    }

    private int GetIndentColumns(string lineText)
    {
        var columns = 0;
        foreach (var ch in lineText)
        {
            if      (ch == ' ')  columns++;
            else if (ch == '\t') columns += TabSize - (columns % TabSize);
            else break;
        }
        return columns;
    }
}