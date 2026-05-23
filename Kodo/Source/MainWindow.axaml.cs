// Licensed under the Kodo Public License v1.1
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
using System.Net.NetworkInformation;
using System.Runtime.CompilerServices;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.Win32;
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

// Compares strings the way humans expect: "1.9" < "1.10" < "1.11"
// by splitting each string into alternating non-digit / digit chunks and
// comparing digit chunks numerically instead of lexicographically.
public sealed class NaturalSortComparer : IComparer<string>
{
    public static readonly NaturalSortComparer OrdinalIgnoreCase = new();

    public int Compare(string? x, string? y)
    {
        if (ReferenceEquals(x, y)) return 0;
        if (x is null) return -1;
        if (y is null) return  1;

        var xi = 0;
        var yi = 0;

        while (xi < x.Length && yi < y.Length)
        {
            var xIsDigit = char.IsAsciiDigit(x[xi]);
            var yIsDigit = char.IsAsciiDigit(y[yi]);

            if (xIsDigit && yIsDigit)
            {
                // Skip leading zeros so "007" == "7" numerically
                while (xi < x.Length && x[xi] == '0') xi++;
                while (yi < y.Length && y[yi] == '0') yi++;

                // Find end of digit run in both strings
                var xStart = xi;
                var yStart = yi;
                while (xi < x.Length && char.IsAsciiDigit(x[xi])) xi++;
                while (yi < y.Length && char.IsAsciiDigit(y[yi])) yi++;

                var xLen = xi - xStart;
                var yLen = yi - yStart;

                // Longer digit sequence is numerically larger
                if (xLen != yLen) return xLen.CompareTo(yLen);

                // Same length: compare digit-by-digit
                var cmp = string.Compare(x, xStart, y, yStart, xLen, StringComparison.Ordinal);
                if (cmp != 0) return cmp;
            }
            else
            {
                // Non-digit chunk: plain case-insensitive char comparison
                var cmp = char.ToUpperInvariant(x[xi])
                              .CompareTo(char.ToUpperInvariant(y[yi]));
                if (cmp != 0) return cmp;
                xi++;
                yi++;
            }
        }

        return (x.Length - xi).CompareTo(y.Length - yi);
    }
}

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
			    ".txt" or ".rst" or ".log" => "TXT",
                ".md" or ".markdown" => "MD",
			    ".png" => "PNG",
			    ".jpg" or ".jpeg" => "JPG",
			    ".gif" => "GIF",
			    ".svg" => "SVG",
			    ".ico" => "ICO",
			    ".webp" => "WBP",
			    ".bmp" => "BMP",
			    ".py" => "PY",
			    ".js" or ".jsx" => "JS",
			    ".ts" or ".tsx" => "TS",
			    ".vue" or ".svelte" => "UI",
			    ".css" or ".scss" or ".less" => "CSS",
			    ".sh" or ".bat" or ".ps1" => ">_",
			    ".zip" or ".tar" or ".gz" or ".rar" => "ZIP",
			    ".cpp" or ".cc" or ".cxx" => "C++",
			    ".c" => "C",
			    ".h" or ".hpp" or ".hxx" => "C++",
			    ".rs" => "RS",
			    ".go" => "GO",
			    ".rb" => "RB",
			    ".java" => "JAVA",
			    ".kt" or ".kts" => "KT",
			    ".swift" => "SW",
			    ".fs" or ".fsi" or ".fsx" => "F#",
			    ".sql" => "DB",
			    ".lua" => "LUA",
			    ".r" => "R",
			    ".lock" => "Lk",
				".csv" or ".tsv" => "CSV",
				".nova" => "NOVA",
				".kox" => "KOX",
                ".sln" => "SLN",
                ".manifest" => "MF",
                ".exe" => "EXE",
                ".dll" => "DLL",
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
    public string[] Functions { get; set; } = [];
    public string[] Properties { get; set; } = [];
    public string[] Namespaces { get; set; } = [];
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
    public DateTime? InstalledOnUtc { get; set; }
    public ExtensionThemeDefinition? ThemeDefinition { get; set; }
    public string ThemeCardThemeId => ThemeDefinition?.ThemeId ?? string.Empty;
    public string ThemeCardDisplayName => ThemeDefinition?.DisplayName ?? Name;
    public string ThemeCardPreviewBackground => ThemeDefinition?.PreviewBackground ?? "#000000";
    public string ThemeCardPreviewBorder => ThemeDefinition?.PreviewBorder ?? "#4A4A4A";
    public string ThemeCardAccent => ThemeDefinition?.Accent ?? "#8C00FF";
    // True for the 2nd, 3rd, etc. entries split out of a multi-theme array -
    // they appear in ThemeExtensions but are hidden from the Installed list.
    public bool IsThemeSubEntry { get; init; }
    // Raw PNG bytes read from icon.png on the background scan thread.
    // Decoded into IconImage on the UI thread by ApplyLoadedExtensionsResult.
    public byte[]? IconBytes { get; set; }
    // Optional icon loaded from icon.png inside the .kox / folder
    public Bitmap? IconImage { get; set; }
    // SVG text for icons sourced from the marketplace index (set on UI thread).
    public string? SvgData { get; set; }
    // Fallback: first two letters of the name, shown when no icon is present
    public string NameAbbreviation => Name.Length >= 2 ? Name[..2] : Name;
    public bool HasIcon => IconImage is not null || SvgData is not null;
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
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SvgData)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasIcon)));
    }
}

public sealed class LanguageSyntaxProfile
{
    public string[] Extensions { get; init; } = [];
    public string[] Keywords { get; init; } = [];
    public string[] Types { get; init; } = [];
    public string[] Functions { get; init; } = [];
    public string[] Properties { get; init; } = [];
    public string[] Namespaces { get; init; } = [];
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

public sealed class TerminalShellOption
{
    public string Id { get; init; } = string.Empty;
    public string DisplayName { get; init; } = string.Empty;
    public string FileName { get; init; } = string.Empty;
    public string Arguments { get; init; } = string.Empty;
}

public static class ExtensionSortModes
{
    public const string Alphabetical = "A-Z";
    public const string ReverseAlphabetical = "Z-A";
    public const string RecentlyInstalled = "Recently Installed";
    public const string UpdatesAvailable = "Updates Available";
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
    private DateTime? _installedOnUtc;

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

    private string? _svgData;
    public string? SvgData
    {
        get => _svgData;
        set
        {
            if (_svgData == value) return;
            _svgData = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(HasIcon));
        }
    }

    public bool HasIcon => IconImage is not null || SvgData is not null;
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

    public DateTime? InstalledOnUtc
    {
        get => _installedOnUtc;
        private set
        {
            if (_installedOnUtc == value) return;
            _installedOnUtc = value;
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
        InstalledOnUtc = installedExtension?.InstalledOnUtc;
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
    private const string SettingsFileName = "kodosettings.json";
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
    private const string DiscordServerUrl = "https://discord.gg/cUQ6C88Z9C";

    private string? _currentFilePath;
    // Encoding detected (or chosen) for the currently open file. Defaults to UTF-8.
    private System.Text.Encoding _currentFileEncoding = System.Text.Encoding.UTF8;
    private string? _currentFolderPath;
    private DiscordRpcClient? _discordRpcClient;
    private readonly DispatcherTimer _autoSaveTimer = new() { Interval = TimeSpan.FromSeconds(2) };
    private readonly DispatcherTimer _autoSaveStatusTimer = new() { Interval = TimeSpan.FromSeconds(3) };
    private readonly DispatcherTimer _discordReconnectTimer = new() { Interval = TimeSpan.FromSeconds(10) };
    private readonly DispatcherTimer _editorStateRefreshTimer = new() { Interval = TimeSpan.FromMilliseconds(75) };
    private readonly DispatcherTimer _wordCountRefreshTimer = new() { Interval = TimeSpan.FromMilliseconds(175) };
    private readonly DispatcherTimer _settingsSaveDebounceTimer = new() { Interval = TimeSpan.FromMilliseconds(400) };
    // Polls the Windows accent registry key so the blob and active accent stay
    // live without requiring the Microsoft.Win32.SystemEvents NuGet package.
    private readonly DispatcherTimer _windowsAccentPollTimer = new() { Interval = TimeSpan.FromSeconds(2) };
    private string _lastSeenWindowsAccentHex = string.Empty;
    private readonly RainbowBracketColorizer _rainbowBracketColorizer = new();
    private readonly InterpolatedStringColorizer _interpolatedStringColorizer = new();
    private readonly MarkdownColorizer _markdownColorizer = new();
    private EditorTab? _activeEditorTab;
    private int _nextUntitledTabNumber = 1;
    private string? _autoSaveStatusMessage;
    private bool _isAutoSaveEnabled;
    private bool _isDirty;
    private bool _isSaving;
    private bool _isDiscordRichPresenceEnabled;
    private bool _hasUntitledDocument;
    private bool _isRefreshingExtensions;
    private bool _isUpdatingAllExtensions;
    private bool _isRefreshingLatestRelease;
    private bool _isSettingsPageVisible;
    private bool _isExtensionsPageVisible;
    private bool _isTutorialPageVisible;
    private bool _tutorialOpenedFromSettings;
    private bool _isWhatsNewPageVisible;
    private bool _isWhatsNewExpanded;
    private bool _isHomePageVisible;
    private bool _isFileExplorerVisible;
    private bool _isFileTreeExpanded;
    private bool _isStatusBarFilePathVisible = true;
    private bool _isWordWrapEnabled;
    private bool _suppressExplorerWidthRefresh;
    private bool _isConfirmBeforeClosingUnsavedTabsEnabled = true;
    private bool _isRestoreOpenTabsOnLaunchEnabled;
    private bool _isMarketplaceTabSelected;
    private bool _suppressDirtyTracking;
    // True during the constructor + OnOpened startup sequence so that incidental
    // SaveSettings() calls (tab restore, CollectionChanged, ApplyTheme) cannot
    // overwrite the just-loaded settings before the window is fully initialised.
    private bool _suppressSettingsSave;
    private int _tabSize = 4;
    private int _editorFontSize = 14;
    private string _accentColorMode = "kodo";   // "kodo" | "windows" | "custom"
    private string _customAccentHex = "#8C00FF";
    // The accent colour supplied by the active theme; restored when switching back to "kodo" mode.
    private string _themeAccentHex = "#8C00FF";
    private string _currentThemeName = "Dark";
    private string _requestedThemeName = "Dark";
    private string _editorStatsText = "0 lines";
    private string _wordCountText = string.Empty;
    private bool _pendingFullStateRefresh = true;
    private string _lastDiscordPresenceDetails = string.Empty;
    private string _lastDiscordPresenceState = string.Empty;
    private readonly DateTime _sessionStart = DateTime.UtcNow;
    // True when settings.json did not exist on this launch - used to show the tutorial once.
    private bool _isFirstLaunch;
    private bool _hasCompletedTutorial;
    private string _extensionsStatusText = "Drop .kox extension files into the Extensions folder to install them.";
    private string _latestReleaseStatusText = "Loading latest release...";
    private string _marketplaceConnectivityMessage = string.Empty;
    private bool _isMarketplaceConnectivityWarningVisible;
    private ReleaseInfo? _latestRelease;
    private LoadedExtension? _currentLanguageExtension;
    private Bitmap? _currentImagePreview;
    private FileSystemWatcher? _extensionsFolderWatcher;
    private FileSystemWatcher? _projectExtensionsFolderWatcher;
    private readonly DispatcherTimer _extensionsRefreshDebounceTimer = new() { Interval = TimeSpan.FromMilliseconds(250) };
    private readonly IndentGuideBackgroundRenderer _indentGuideRenderer = new();
    private readonly List<string> _startupOpenTabPaths = [];
    private readonly Dictionary<string, IBrush> _brushCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, byte[]> _marketplaceIconBytesCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, DateTime> _warningDialogCooldowns = new(StringComparer.OrdinalIgnoreCase);
    private readonly SemaphoreSlim _iconFetchSemaphore = new(4, 4);
    private static readonly TimeSpan ExtensionsRefreshCooldown = TimeSpan.FromSeconds(8);
    private static readonly TimeSpan WarningDialogCooldown = TimeSpan.FromSeconds(10);
    private static readonly HashSet<string> ImagePreviewExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".png", ".apng", ".jpg", ".jpeg", ".jpe", ".jfif", ".bmp", ".dib", ".gif",
        ".webp", ".ico", ".cur", ".tif", ".tiff"
    };
    private DateTime _lastExtensionsRefreshUtc = DateTime.MinValue;
    private string? _startupActiveTabPath;
    private string? _startupFilePath;
    private string _extensionSearchText = string.Empty;
    private string _selectedInstalledExtensionSort = ExtensionSortModes.Alphabetical;
    private string _selectedMarketplaceExtensionSort = ExtensionSortModes.Alphabetical;
    // Personalization - set in Settings, persisted in settings.json.
    // _userCountry:        ISO 3166-1 alpha-2 country code (e.g. "CA", "US", "GB").
    //                      Empty → auto-detected from the OS regional settings.
    // _userHemisphere:     0 = auto-detect from country, 1 = northern, 2 = southern.
    // _userTimezoneOffset: UTC offset string (e.g. "-5", "+1"). Empty → auto-detect.
    private string _userCountry = string.Empty;
    private int    _userHemisphere = 0;
    private string _userTimezoneOffset = string.Empty;
    private string _userName = string.Empty;
    private bool _isFindPanelVisible;
    private bool _isFileCorrupted;
    private readonly HashSet<EditorTab> _corruptedTabs = new(ReferenceEqualityComparer.Instance);
    private TerminalSession? _activeTerminalSession;
    private TerminalShellOption? _selectedTerminalShell;
    private int _nextTerminalNumber = 1;
    private bool _isTerminalVisible;
    private bool _isTerminalSupported = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);

    // Caches compiled KodoHighlightingDefinition instances by LoadedExtension identity.
    // Building one involves compiling multiple Regex objects (RegexOptions.Compiled), which
    // is expensive enough to cause a noticeable delay on every tab switch. The cache lives
    // for the session and is cleared whenever extensions are reloaded so it never goes stale.
    private readonly Dictionary<LoadedExtension, KodoHighlightingDefinition> _highlightingCache =
        new(ReferenceEqualityComparer.Instance);
    private string _findText = string.Empty;
    private int _tutorialStepIndex;

    // File tree clipboard state (cut/copy/paste)
    private string? _clipboardItemPath;
    private bool _clipboardItemIsDirectory;
    private bool _clipboardIsCut;
    private event PropertyChangedEventHandler? ViewModelPropertyChanged;
    private static readonly HttpClient MarketplaceHttpClient = CreateHttpClient();
    private static readonly TutorialStep[] TutorialSteps =
    [
        new(
            "",
            "Welcome to Kodo!",
            "A fast, focused code editor built to stay out of your way. This short tutorial will walk you through the essentials - it only takes a minute.",
            "",
            "",
            "",
            "",
            ""
        ),
        new(
            "Welcome",
            "Meet your workspace",
            "Kodo starts on Home so you can jump straight into a recent project, create a new file, or open an existing folder without hunting through menus.",
            "Ctrl+H",
            "Home is your launchpad",
            "Open recent files and folders in one click.",
            "Start fresh work quickly with New File or Open Folder.",
            "Jump back here anytime from the activity bar."
        ),
        new(
            "Editing",
            "Create and work fast",
            "Use the editor for scratch files or full projects, then lean on autosave, tab restore, and the file explorer when you are moving across a bigger codebase.",
            "Ctrl+N / Ctrl+O / Ctrl+K",
            "Core editing flow",
            "Create a file instantly with Ctrl+N.",
            "Open a file with Ctrl+O or a folder with Ctrl+K.",
            "Keep momentum with tabs, autosave, and explorer tools."
        ),
        new(
            "Marketplace",
            "Install language support and themes",
            "The Marketplace is where Kodo pulls syntax highlighting packs, language definitions, and theme extensions from the web so you can tailor the app to your stack.",
            "Ctrl+E",
            "Extensions, themes, updates",
            "Browse installable extensions from inside the app.",
            "Update installed packages when newer versions appear.",
            "Watch for connectivity warnings if downloads cannot reach the internet."
        ),
        new(
            "Settings",
            "Tune Kodo to your workflow",
            "Adjust themes, font size, tab width, autosave, and other quality-of-life settings so the editor feels right for the way you work.",
            "Ctrl+,",
            "Personalize the experience",
            "Switch between built-in and extension themes.",
            "Set editor font size and tab behavior.",
            "Control autosave and launch preferences."
        ),
        new(
            "Set up",
            "Make Kodo yours",
            "Pick a theme and accent colour, then tell Kodo a little about yourself so the welcome message on Home feels personal. You can change any of these later in Settings.",
            "Ctrl+,  ·  Settings",
            "Why personalise?",
            "Your name makes greetings feel like they're written for you.",
            "Country and hemisphere keep seasonal messages accurate.",
            "Theme and accent colour apply instantly across the whole app."
        )
    ];

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

    // Closing characters - when typed over an existing auto-inserted closer, skip past it
    private static readonly HashSet<char> ClosingChars = new() { ')', ']', '}', '>', '"', '\'', '`' };
    private static readonly Dictionary<string, string> FenceLanguageAliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["c#"] = "cs",
        ["csharp"] = "cs",
        ["f#"] = "fs",
        ["fsharp"] = "fs",
        ["js"] = "js",
        ["javascript"] = "js",
        ["ts"] = "ts",
        ["typescript"] = "ts",
        ["py"] = "py",
        ["python"] = "py",
        ["rb"] = "rb",
        ["ruby"] = "rb",
        ["rs"] = "rs",
        ["rust"] = "rs",
        ["ps"] = "ps1",
        ["powershell"] = "ps1",
        ["shell"] = "sh",
        ["bash"] = "sh",
        ["zsh"] = "sh",
        ["yml"] = "yml",
        ["yaml"] = "yml",
        ["jsonc"] = "json",
        ["md"] = "md",
        ["markdown"] = "md"
    };

    private static HttpClient CreateHttpClient()
    {
        var client = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
        client.DefaultRequestHeaders.UserAgent.ParseAdd("Kodo/1.0.0-DEV (https://github.com/KerbalMissile/Kodo)");
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
    public ObservableCollection<TerminalSession> TerminalSessions { get; } = new();
    public ObservableCollection<LoadedExtension> LoadedExtensions { get; } = new();
    public ObservableCollection<MarketplaceExtension> MarketplaceExtensions { get; } = new();
    public ObservableCollection<string> ExtensionLoadErrors { get; } = new();
    public ObservableCollection<TerminalShellOption> AvailableTerminalShells { get; } = new();

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
        EditorTextBox.TextArea.TextView.LineTransformers.Add(_interpolatedStringColorizer);
        EditorTextBox.TextArea.TextView.LineTransformers.Add(_markdownColorizer);
        EditorTextBox.TextArea.TextView.LinkTextForegroundBrush = Brush.Parse("#5BA3D9");
        EditorTextBox.TextArea.TextView.LinkTextBackgroundBrush = Brushes.Transparent;
        OpenTabs.CollectionChanged += OpenTabs_CollectionChanged;
        TerminalSessions.CollectionChanged += TerminalSessions_CollectionChanged;
        FileTreeItems.CollectionChanged += FileTreeItems_CollectionChanged;
        // TextEditor uses EventHandler (not RoutedEventHandler), so hook up in code-behind
        EditorTextBox.TextChanged += EditorTextBox_OnTextChanged;
        EditorTextBox.TextArea.Caret.PositionChanged += (_, _) => QueueRefreshState();
		// Auto-completion: insert closing bracket/quote after opener, skip-over when typing a closer
        EditorTextBox.TextArea.TextEntering += EditorTextArea_OnTextEntering;
        EditorTextBox.TextArea.TextEntered  += EditorTextArea_OnTextEntered;
        AddHandler(InputElement.KeyDownEvent, MainWindow_EditorKeyIntercept_OnKeyDown, RoutingStrategies.Tunnel, handledEventsToo: true);
        _isFirstLaunch = !File.Exists(SettingsFilePath);
        var settings = LoadSettings();
        _requestedThemeName = string.IsNullOrWhiteSpace(settings.ThemeName) ? "Dark" : settings.ThemeName;
        _isAutoSaveEnabled = settings.AutoSaveEnabled;
        _isDiscordRichPresenceEnabled = settings.DiscordRichPresenceEnabled;
        _isStatusBarFilePathVisible = settings.StatusBarFilePathVisible;
        _isWordWrapEnabled = settings.WordWrapEnabled;
        _isConfirmBeforeClosingUnsavedTabsEnabled = settings.ConfirmBeforeClosingUnsavedTabsEnabled;
        _isRestoreOpenTabsOnLaunchEnabled = settings.RestoreOpenTabsOnLaunchEnabled;
        _hasCompletedTutorial = settings.HasCompletedTutorial;
        _accentColorMode = settings.AccentColorMode is "kodo" or "windows" or "custom"
            ? settings.AccentColorMode : "kodo";
        _customAccentHex = string.IsNullOrWhiteSpace(settings.CustomAccentHex)
            ? "#8C00FF" : settings.CustomAccentHex;
        _tabSize = NormalizeTabSize(settings.TabSize);
        _editorFontSize = settings.EditorFontSize is >= 8 and <= 32 ? settings.EditorFontSize : 14;
        _userCountry = string.IsNullOrWhiteSpace(settings.UserCountry)
            ? DetectCountryCode()
            : settings.UserCountry.ToUpperInvariant();
        _userHemisphere     = settings.UserHemisphere is >= 0 and <= 2 ? settings.UserHemisphere : 0;
        _userTimezoneOffset = settings.UserTimezoneOffset ?? string.Empty;
        _userName           = settings.UserName ?? string.Empty;
        _isTerminalVisible = false; // always start hidden; user opens it manually
        _startupOpenTabPaths.AddRange(settings.OpenTabPaths
            .Where(path => File.Exists(path))
            .Distinct(StringComparer.OrdinalIgnoreCase));
        _startupActiveTabPath = settings.ActiveTabPath;
        LoadRecentFiles(settings.RecentFiles);
        RefreshAvailableTerminalShells(settings.PreferredTerminalShellId);
        _autoSaveTimer.Tick += AutoSaveTimer_OnTick;
        _autoSaveStatusTimer.Tick += AutoSaveStatusTimer_OnTick;
        _discordReconnectTimer.Tick += DiscordReconnectTimer_OnTick;
        _editorStateRefreshTimer.Tick += EditorStateRefreshTimer_OnTick;
        _wordCountRefreshTimer.Tick += WordCountRefreshTimer_OnTick;
        _settingsSaveDebounceTimer.Tick += SettingsSaveDebounceTimer_OnTick;
        _extensionsRefreshDebounceTimer.Tick += ExtensionsRefreshDebounceTimer_OnTick;

        // ── Pre-DataContext theme bootstrap ────────────────────────────────────
        // Extensions must be loaded before ApplyTheme so custom themes are available.
        // Both calls happen before DataContext = this so that when bindings first
        // evaluate, all brush properties already carry the correct values.
        // This eliminates the one-frame flash / re-paint that was visible on startup
        // when using the Light theme or any high-contrast extension theme.
        EnsureExtensionsFolder();
        SetupExtensionFolderWatchers();
        LoadExtensions();
        ApplyThemeBrushes(_requestedThemeName);

        DataContext = this;
        IsHomePageVisible = true;

        // Kick off the full async refresh (marketplace fetch, icon loading, etc.)
        // now that the UI is live. The synchronous LoadExtensions() above already
        // populated the theme; the async path below will update everything else
        // (marketplace data, update badges) without touching the brush values again
        // unless the user has changed settings while offline.
        UpdateDiscordRichPresenceLifecycle();
        ApplyEditorSettings();
        NetworkChange.NetworkAvailabilityChanged += NetworkChange_OnNetworkAvailabilityChanged;
        NetworkChange.NetworkAddressChanged += NetworkChange_OnNetworkAddressChanged;
        RefreshMarketplaceConnectivityState();
        // Poll the Windows accent registry every 2 s so the blob preview and
        // active accent stay live without the Microsoft.Win32.SystemEvents package.
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            _lastSeenWindowsAccentHex = GetWindowsAccentColor() ?? string.Empty;
            _windowsAccentPollTimer.Tick += WindowsAccentPollTimer_OnTick;
            _windowsAccentPollTimer.Start();
        }
        Opened += MainWindow_OnOpened;
        Closing += MainWindow_OnClosing;
        Closed += MainWindow_OnClosed;
        RefreshState(fullRefresh: true);
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
            // ScanInstalledExtensions is pure I/O — run it off the UI thread.
            // Everything after the scan (collection mutations, PropertyChanged
            // notifications, marketplace fetch) MUST run on the UI thread because
            // Avalonia's binding engine requires it.  We marshal explicitly with
            // InvokeAsync rather than relying on the SynchronizationContext, which
            // is not guaranteed to be the UI context after a Task.Run await.
            var extensionScan = await Task.Run(ScanInstalledExtensions);
            await Dispatcher.UIThread.InvokeAsync(() => ApplyLoadedExtensionsResult(extensionScan));
            await LoadMarketplaceExtensionsAsync();
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
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
            });
        }
        catch (Exception ex)
        {
            await Dispatcher.UIThread.InvokeAsync(() => ExtensionsStatusText = $"Refresh failed: {ex.Message}");
            await ShowWarningDialogAsync("Extension refresh", ex);
        }
        finally
        {
            await Dispatcher.UIThread.InvokeAsync(() => IsRefreshingExtensions = false);
        }
    }

    private void LoadExtensions()
    {
        ApplyLoadedExtensionsResult(ScanInstalledExtensions());
    }

    private ExtensionScanResult ScanInstalledExtensions()
    {
        // Compiled highlighting definitions are keyed by LoadedExtension reference.
        // Extension reload creates new instances, so the old entries are now orphaned;
        // drop them here to avoid a memory leak and ensure fresh definitions are built.
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

        return new ExtensionScanResult(loadedExtensions, extensionLoadErrors);
    }

    private void ApplyLoadedExtensionsResult(ExtensionScanResult result)
    {
        _highlightingCache.Clear();
        SyncObservableCollection(LoadedExtensions, result.Extensions, ext => ext.Id);
        SyncObservableCollection(ExtensionLoadErrors, result.LoadErrors, error => error);

        // Decode icon bitmaps here on the UI thread. The background scan stored raw
        // PNG bytes in IconBytes to avoid creating Avalonia Bitmaps off-thread (which
        // is unsafe and causes silent failures). Now that we're on the UI thread we
        // can safely decode them and clear the staging bytes to free the memory.
        foreach (var ext in LoadedExtensions)
        {
            if (ext.IconImage is null && ext.IconBytes is not null)
            {
                ext.IconImage = DecodeBitmapOnUiThread(ext.IconBytes);
                ext.IconBytes = null;
                ext.NotifyIconChanged();
            }
        }

        // Re-stamp IsActiveTheme on every theme extension now that the collection
        // may contain brand-new LoadedExtension instances (SyncObservableCollection
        // adds new objects with IsActiveTheme = false). The CurrentThemeName setter
        // only ran once during ApplyThemeBrushes, before these objects existed.
        foreach (var ext in ThemeExtensions)
            ext.IsActiveTheme = string.Equals(ext.ThemeCardThemeId, _currentThemeName, StringComparison.OrdinalIgnoreCase);

        OnPropertyChanged(nameof(ExtensionLoadErrors));
        OnPropertyChanged(nameof(VisibleLoadedExtensions));
        NotifyExtensionFiltersChanged();
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

        await Dispatcher.UIThread.InvokeAsync(() => RefreshMarketplaceConnectivityState());

        try
        {
            // Request the index with no-cache so GitHub's CDN always serves the
            // latest committed version rather than a stale cached copy.
            using var indexRequest = new HttpRequestMessage(HttpMethod.Get, DefaultMarketplaceIndexUrl);
            indexRequest.Headers.CacheControl = new System.Net.Http.Headers.CacheControlHeaderValue { NoCache = true, NoStore = true };
            using var indexResponse = await MarketplaceHttpClient.SendAsync(indexRequest);
            indexResponse.EnsureSuccessStatusCode();
            var remoteJson = await indexResponse.Content.ReadAsStringAsync();
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
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                RefreshMarketplaceConnectivityState("Marketplace fetch", ex);
            });
            await ShowWarningDialogAsync("Marketplace fetch", ex);
        }

        // All ObservableCollection mutations and PropertyChanged notifications must
        // run on the UI thread — Avalonia's binding engine requires it.
        Dictionary<string, string> marketplaceIconMap = [];
        await Dispatcher.UIThread.InvokeAsync(() =>
        {
            SyncMarketplaceExtensionCollection(MarketplaceExtensions, marketplaceExtensions);
            SyncObservableCollection(
                ExtensionLoadErrors,
                ExtensionLoadErrors.Concat(extensionLoadErrors).Distinct().ToList(),
                error => error);

            SyncMarketplaceInstallStates();
            OnPropertyChanged(nameof(ExtensionLoadErrors));
            OnPropertyChanged(nameof(IsMarketplaceEmptyVisible));
            NotifyExtensionFiltersChanged();

            marketplaceIconMap = MarketplaceExtensions
                .Where(entry => !string.IsNullOrWhiteSpace(entry.IconUrl))
                .ToDictionary(entry => entry.Id, entry => entry.IconUrl, StringComparer.OrdinalIgnoreCase);
        });

        await FetchMarketplaceIconsAsync(marketplaceIconMap);
        await FetchInstalledExtensionIconsAsync(marketplaceIconMap);
    }

    private async Task FetchInstalledExtensionIconsAsync(IReadOnlyDictionary<string, string> marketplaceIconMap)
    {
        var tasks = LoadedExtensions
            .Select(ext => (ext, iconUrl: marketplaceIconMap.TryGetValue(ext.Id, out var iconUrl) ? iconUrl : string.Empty))
            .Where(pair => !string.IsNullOrWhiteSpace(pair.iconUrl))
            .Select(async pair =>
            {
                try
                {
                    var icon = await GetCachedIconAsync(pair.iconUrl);

                    await Dispatcher.UIThread.InvokeAsync(() =>
                    {
                        if (icon.HasValue)
                        {
                            // Index icon fetched successfully — use it, replacing any kox icon.
                            ReplaceLoadedExtensionIcon(pair.ext, icon);
                        }
                        // else: fetch returned nothing (bad URL, corrupt bytes, etc.) —
                        // leave whatever the kox provided in place.
                    });
                }
                catch
                {
                    // Network failure — leave the kox icon (or abbreviation) in place.
                }
            });

        await Task.WhenAll(tasks);
    }
    private async Task FetchMarketplaceIconsAsync(IReadOnlyDictionary<string, string> marketplaceIconMap)
    {
        // Apply icons whose bytes are already cached synchronously on the UI thread —
        // this covers entries that came in via SyncMarketplaceExtensionCollection without
        // a bitmap (e.g. a brand-new item that happened to share a URL with a previously
        // fetched icon) and avoids an unnecessary async round-trip for them.
        await Dispatcher.UIThread.InvokeAsync(() =>
        {
            foreach (var entry in MarketplaceExtensions)
            {
                if (entry.IconImage is not null || entry.SvgData is not null)
                    continue;

                if (!marketplaceIconMap.TryGetValue(entry.Id, out var cachedUrl))
                    continue;

                if (!_marketplaceIconBytesCache.TryGetValue(cachedUrl, out var cachedBytes))
                    continue;

                var icon = DecodeCachedIconBytes(cachedBytes);
                if (icon.HasValue)
                    ReplaceMarketplaceIcon(entry, icon);
            }
        });

        var tasks = MarketplaceExtensions
            .Where(entry => entry.IconImage is null && entry.SvgData is null && marketplaceIconMap.TryGetValue(entry.Id, out _))
            .Select(async entry =>
            {
                try
                {
                    var icon = await GetCachedIconAsync(marketplaceIconMap[entry.Id]);
                    if (!icon.HasValue)
                        return;

                    await Dispatcher.UIThread.InvokeAsync(() =>
                    {
                        ReplaceMarketplaceIcon(entry, icon);
                    });
                }
                catch { /* Network failure - fallback to local icon or abbreviation. */ }
            });

        await Task.WhenAll(tasks);
    }

    // Discriminated result from GetCachedIconAsync.
    private readonly record struct IconResult(Bitmap? Bitmap, string? SvgData)
    {
        public bool HasValue => Bitmap is not null || SvgData is not null;
    }

    private static bool IsSvgContent(byte[] bytes)
    {
        // SVG files start with either a UTF-8 BOM + '<' or directly with '<'.
        // Check for the <?xml or <svg opening tag in the first 512 bytes.
        var header = System.Text.Encoding.UTF8.GetString(bytes, 0, Math.Min(bytes.Length, 512));
        return header.TrimStart().StartsWith("<?xml", StringComparison.OrdinalIgnoreCase)
            || header.TrimStart().StartsWith("<svg", StringComparison.OrdinalIgnoreCase);
    }

    private async Task<IconResult> GetCachedIconAsync(string iconUrl)
    {
        if (string.IsNullOrWhiteSpace(iconUrl))
            return default;

        // Fast path: bytes already cached.
        if (_marketplaceIconBytesCache.TryGetValue(iconUrl, out var bytes))
            return DecodeCachedIconBytes(bytes);

        // Cache miss — fetch under semaphore to avoid duplicate requests.
        await _iconFetchSemaphore.WaitAsync();
        try
        {
            if (!_marketplaceIconBytesCache.TryGetValue(iconUrl, out bytes))
            {
                bytes = await MarketplaceHttpClient.GetByteArrayAsync(iconUrl);
                _marketplaceIconBytesCache[iconUrl] = bytes;
            }
        }
        finally
        {
            _iconFetchSemaphore.Release();
        }

        return DecodeCachedIconBytes(bytes);
    }

    private static IconResult DecodeCachedIconBytes(byte[] bytes)
    {
        if (IsSvgContent(bytes))
        {
            try
            {
                return new IconResult(null, System.Text.Encoding.UTF8.GetString(bytes));
            }
            catch { return default; }
        }

        try
        {
            using var ms = new MemoryStream(bytes);
            return new IconResult(new Bitmap(ms), null);
        }
        catch { return default; }
    }

    private static void ReplaceLoadedExtensionIcon(LoadedExtension extension, IconResult icon)
    {
        if (icon.Bitmap is not null)
        {
            if (ReferenceEquals(extension.IconImage, icon.Bitmap)) return;
            extension.IconImage?.Dispose();
            extension.IconImage = icon.Bitmap;
            extension.SvgData = null;
        }
        else if (icon.SvgData is not null)
        {
            extension.IconImage?.Dispose();
            extension.IconImage = null;
            extension.SvgData = icon.SvgData;
        }
        extension.NotifyIconChanged();
    }

    private static void ReplaceMarketplaceIcon(MarketplaceExtension extension, IconResult icon)
    {
        if (icon.Bitmap is not null)
        {
            if (ReferenceEquals(extension.IconImage, icon.Bitmap)) return;
            extension.IconImage?.Dispose();
            extension.IconImage = icon.Bitmap;
            extension.SvgData = null;
        }
        else if (icon.SvgData is not null)
        {
            extension.IconImage?.Dispose();
            extension.IconImage = null;
            extension.SvgData = icon.SvgData;
        }
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

    private static MarketplaceExtension ParseMarketplaceExtension(JsonElement item)
    {
        var id = item.TryGetProperty("id", out var idElement) ? idElement.GetString() ?? string.Empty : string.Empty;
        var declaredVersion = item.TryGetProperty("version", out var versionElement) ? versionElement.GetString() ?? string.Empty : string.Empty;
        var name = item.TryGetProperty("name", out var nameElement) ? nameElement.GetString() ?? string.Empty : string.Empty;
        var type = item.TryGetProperty("type", out var typeElement) ? typeElement.GetString() ?? string.Empty : string.Empty;
        var author = item.TryGetProperty("author", out var authorElement) ? authorElement.GetString() ?? string.Empty : string.Empty;
        var description = item.TryGetProperty("description", out var descriptionElement) ? descriptionElement.GetString() ?? string.Empty : string.Empty;
        var rawDownloadUrl = item.TryGetProperty("downloadUrl", out var downloadUrlElement) ? downloadUrlElement.GetString() ?? string.Empty : string.Empty;
        var declaredFileName = item.TryGetProperty("fileName", out var fileNameElement) ? fileNameElement.GetString() ?? string.Empty : string.Empty;
        var iconUrl = NormalizeIconUrl(item.TryGetProperty("iconUrl", out var iconUrlElement) ? iconUrlElement.GetString() ?? string.Empty : string.Empty);
        var urlFileName = TryGetFileNameFromUrl(rawDownloadUrl);
        var bestKnownVersion = GetHighestKnownExtensionVersion(declaredVersion, declaredFileName, urlFileName);
        var canonicalFileName = GetCanonicalMarketplaceFileName(declaredFileName, urlFileName, bestKnownVersion);
        var canonicalDownloadUrl = NormalizeMarketplaceDownloadUrl(rawDownloadUrl, canonicalFileName);

        return new MarketplaceExtension
        {
            Id = id,
            Version = bestKnownVersion,
            Name = name,
            Type = type,
            Author = author,
            Description = description,
            DownloadUrl = canonicalDownloadUrl,
            FileName = canonicalFileName,
            IconUrl = iconUrl
        };
    }

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
        }

        OnPropertyChanged(nameof(AvailableExtensionUpdatesCount));
        OnPropertyChanged(nameof(IsExtensionUpdateBannerVisible));
        OnPropertyChanged(nameof(ExtensionUpdatesBannerText));
        NotifyExtensionActionStateChanged();
        NotifyExtensionFiltersChanged();
    }

    private void NotifyExtensionFiltersChanged()
    {
        OnPropertyChanged(nameof(FilteredInstalledExtensions));
        OnPropertyChanged(nameof(FilteredMarketplaceExtensions));
    }

    private void NotifyExtensionActionStateChanged()
    {
        OnPropertyChanged(nameof(CanUpdateAllExtensions));
        OnPropertyChanged(nameof(UpdateAllExtensionsButtonText));
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
        var sourceByKey = source.ToDictionary(keySelector);

        for (var i = target.Count - 1; i >= 0; i--)
        {
            var key = keySelector(target[i]);
            if (!sourceByKey.ContainsKey(key))
                target.RemoveAt(i);
        }

        var targetIndexByKey = new Dictionary<TKey, int>();
        for (var i = 0; i < target.Count; i++)
            targetIndexByKey[keySelector(target[i])] = i;

        for (var i = 0; i < source.Count; i++)
        {
            var item = source[i];
            var key = keySelector(item);
            var existingIndex = targetIndexByKey.TryGetValue(key, out var foundIndex) ? foundIndex : -1;

            if (existingIndex == -1)
            {
                target.Insert(Math.Min(i, target.Count), item);
                for (var j = i; j < target.Count; j++)
                    targetIndexByKey[keySelector(target[j])] = j;
                continue;
            }

            if (existingIndex != i)
            {
                target.Move(existingIndex, i);
                var start = Math.Min(existingIndex, i);
                for (var j = start; j < target.Count; j++)
                    targetIndexByKey[keySelector(target[j])] = j;
            }

            if (!ReferenceEquals(target[i], item))
            {
                target[i] = item;
                targetIndexByKey[key] = i;
            }
        }
    }

    // Like SyncObservableCollection<MarketplaceExtension>, but carries the already-fetched
    // IconImage bitmap across to the replacement object instead of discarding it.
    // Without this, every LoadMarketplaceExtensionsAsync call (including the one triggered
    // after an install) replaces items with freshly-parsed objects whose IconImage is null,
    // causing a flicker back to the abbreviation placeholder and an unnecessary re-fetch.
    private static void SyncMarketplaceExtensionCollection(
        ObservableCollection<MarketplaceExtension> target,
        IList<MarketplaceExtension> source)
    {
        var sourceByKey = source.ToDictionary(e => e.Id, StringComparer.OrdinalIgnoreCase);

        for (var i = target.Count - 1; i >= 0; i--)
        {
            if (!sourceByKey.ContainsKey(target[i].Id))
            {
                target[i].IconImage?.Dispose();
                target.RemoveAt(i);
            }
        }

        var targetIndexByKey = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < target.Count; i++)
            targetIndexByKey[target[i].Id] = i;

        for (var i = 0; i < source.Count; i++)
        {
            var incoming = source[i];
            var key = incoming.Id;
            var existingIndex = targetIndexByKey.TryGetValue(key, out var foundIndex) ? foundIndex : -1;

            if (existingIndex == -1)
            {
                target.Insert(Math.Min(i, target.Count), incoming);
                for (var j = i; j < target.Count; j++)
                    targetIndexByKey[target[j].Id] = j;
                continue;
            }

            if (existingIndex != i)
            {
                target.Move(existingIndex, i);
                var start = Math.Min(existingIndex, i);
                for (var j = start; j < target.Count; j++)
                    targetIndexByKey[target[j].Id] = j;
            }

            var existing = target[i];
            if (!ReferenceEquals(existing, incoming))
            {
                // Transfer the already-decoded bitmap/SVG so the UI keeps showing the icon
                // while the rest of the object is refreshed with updated metadata.
                if (existing.IconImage is not null && incoming.IconImage is null)
                    incoming.IconImage = existing.IconImage;
                else
                    existing.IconImage?.Dispose();

                if (existing.SvgData is not null && incoming.SvgData is null)
                    incoming.SvgData = existing.SvgData;

                target[i] = incoming;
                targetIndexByKey[key] = i;
            }
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

        RefreshMarketplaceConnectivityState();
        marketplaceExtension.IsInstalling = true;
        NotifyExtensionActionStateChanged();
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
            ValidateDownloadedExtensionPackage(marketplaceExtension, bytes);
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
            RefreshMarketplaceConnectivityState($"Extension install - {marketplaceExtension.Name}", ex);
            ExtensionsStatusText = $"Failed to install {marketplaceExtension.Name}: {ex.Message}";
            await ShowWarningDialogAsync($"Extension install - {marketplaceExtension.Name}", ex);
        }
        finally
        {
            marketplaceExtension.IsInstalling = false;
            NotifyExtensionActionStateChanged();
            SyncMarketplaceInstallStates();
        }
    }

    private static void ValidateDownloadedExtensionPackage(MarketplaceExtension marketplaceExtension, byte[] packageBytes)
    {
        using var ms = new MemoryStream(packageBytes, writable: false);
        using var archive = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: false);
        var manifestEntry = archive.GetEntry("manifest.json")
            ?? throw new InvalidDataException($"Downloaded package for {marketplaceExtension.Name} is missing manifest.json.");

        using var manifestStream = manifestEntry.Open();
        using var manifestDoc = JsonDocument.Parse(manifestStream);
        var manifest = manifestDoc.RootElement;
        var manifestId = manifest.TryGetProperty("id", out var idElement) ? idElement.GetString() ?? string.Empty : string.Empty;
        var manifestVersion = manifest.TryGetProperty("version", out var versionElement) ? versionElement.GetString() ?? string.Empty : string.Empty;

        if (!string.Equals(manifestId, marketplaceExtension.Id, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidDataException(
                $"Downloaded package id '{manifestId}' does not match expected id '{marketplaceExtension.Id}'.");
        }

        if (CompareExtensionVersions(manifestVersion, marketplaceExtension.Version) < 0)
        {
            throw new InvalidDataException(
                $"Downloaded package version '{manifestVersion}' is older than the marketplace version '{marketplaceExtension.Version}'.");
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
            await ShowWarningDialogAsync($"Extension uninstall - {extension.Name}", ex);
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

    private static string GetHighestKnownExtensionVersion(params string[] candidates)
    {
        var bestVersion = string.Empty;
        foreach (var candidate in candidates)
        {
            var normalizedCandidate = ExtractVersionFromName(candidate);
            if (string.IsNullOrWhiteSpace(normalizedCandidate))
                normalizedCandidate = candidate;

            if (CompareExtensionVersions(normalizedCandidate, bestVersion) > 0)
                bestVersion = normalizedCandidate;
        }

        return string.IsNullOrWhiteSpace(bestVersion) ? string.Empty : bestVersion;
    }

    private static string GetCanonicalMarketplaceFileName(string declaredFileName, string urlFileName, string bestKnownVersion)
    {
        var baseFileName = !string.IsNullOrWhiteSpace(declaredFileName)
            ? declaredFileName
            : !string.IsNullOrWhiteSpace(urlFileName) && !string.Equals(urlFileName, "extension.kox", StringComparison.OrdinalIgnoreCase)
                ? urlFileName
                : string.Empty;

        if (string.IsNullOrWhiteSpace(baseFileName))
            return string.Empty;

        if (string.IsNullOrWhiteSpace(bestKnownVersion))
            return baseFileName;

        var fileVersion = ExtractVersionFromName(baseFileName);
        if (string.IsNullOrWhiteSpace(fileVersion))
            return baseFileName;

        return string.Equals(fileVersion, bestKnownVersion, StringComparison.OrdinalIgnoreCase)
            ? baseFileName
            : ReplaceVersionInValue(baseFileName, fileVersion, bestKnownVersion);
    }

    private static string NormalizeMarketplaceDownloadUrl(string rawDownloadUrl, string canonicalFileName)
    {
        if (string.IsNullOrWhiteSpace(rawDownloadUrl) || string.IsNullOrWhiteSpace(canonicalFileName))
            return rawDownloadUrl;

        if (!Uri.TryCreate(rawDownloadUrl, UriKind.Absolute, out var uri))
            return rawDownloadUrl;

        var absolutePath = uri.AbsolutePath;
        var lastSlashIndex = absolutePath.LastIndexOf('/');
        if (lastSlashIndex < 0)
            return rawDownloadUrl;

        var pathPrefix = absolutePath[..(lastSlashIndex + 1)];
        var normalizedPath = pathPrefix + Uri.EscapeDataString(canonicalFileName);
        var builder = new UriBuilder(uri) { Path = normalizedPath };
        return builder.Uri.ToString();
    }

    /// <summary>
    /// Converts a GitHub "blob" viewer URL to its raw.githubusercontent.com equivalent
    /// so the image bytes can be fetched directly instead of receiving an HTML page.
    /// e.g. https://github.com/owner/repo/blob/main/icon.png
    ///   -> https://raw.githubusercontent.com/owner/repo/main/icon.png
    /// Non-GitHub and already-raw URLs are returned unchanged.
    /// </summary>
    private static string NormalizeIconUrl(string iconUrl)
    {
        if (string.IsNullOrWhiteSpace(iconUrl))
            return iconUrl;

        if (!Uri.TryCreate(iconUrl, UriKind.Absolute, out var uri))
            return iconUrl;

        if (uri.Host.Equals("raw.githubusercontent.com", StringComparison.OrdinalIgnoreCase))
            return iconUrl;

        if (uri.Host.Equals("github.com", StringComparison.OrdinalIgnoreCase))
        {
            // Expect: /{owner}/{repo}/blob/{branch}/{...path}
            var segments = uri.AbsolutePath.TrimStart('/').Split('/');
            if (segments.Length >= 5 &&
                segments[2].Equals("blob", StringComparison.OrdinalIgnoreCase))
            {
                var owner  = segments[0];
                var repo   = segments[1];
                var branch = segments[3];
                var path   = string.Join("/", segments, 4, segments.Length - 4);
                return $"https://raw.githubusercontent.com/{owner}/{repo}/{branch}/{path}";
            }
        }

        return iconUrl;
    }

    private static string ReplaceVersionInValue(string value, string oldVersion, string newVersion)
    {
        if (string.IsNullOrWhiteSpace(value) ||
            string.IsNullOrWhiteSpace(oldVersion) ||
            string.IsNullOrWhiteSpace(newVersion))
        {
            return value;
        }

        return Regex.Replace(
            value,
            Regex.Escape(oldVersion),
            newVersion,
            RegexOptions.IgnoreCase,
            TimeSpan.FromMilliseconds(250));
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

    private static DateTime? GetExtensionSourceActivityUtc(string path, bool isDirectory)
    {
        try
        {
            var createdUtc = isDirectory ? Directory.GetCreationTimeUtc(path) : File.GetCreationTimeUtc(path);
            var modifiedUtc = isDirectory ? Directory.GetLastWriteTimeUtc(path) : File.GetLastWriteTimeUtc(path);
            var activityUtc = createdUtc > modifiedUtc ? createdUtc : modifiedUtc;
            return activityUtc == DateTime.MinValue ? null : activityUtc;
        }
        catch
        {
            return null;
        }
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
        baseExt.InstalledOnUtc = GetExtensionSourceActivityUtc(folderPath, isDirectory: true);

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
            baseExt.IconBytes = ReadIconBytesFromStream(iconStream);
        }

        var themePath = Path.Combine(folderPath, "theme.json");
        if (!File.Exists(themePath))
        {
            // No theme file - yield the extension as-is (language extension, etc.)
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
        Functions         = src.Functions,
        Properties        = src.Properties,
        Namespaces        = src.Namespaces,
        CommentLine       = src.CommentLine,
        CommentBlockStart = src.CommentBlockStart,
        CommentBlockEnd   = src.CommentBlockEnd,
        StringDelimiters  = src.StringDelimiters.ToArray(),
        MultiLineStringDelimiters = src.MultiLineStringDelimiters.ToArray(),
        DisableSingleQuoteStrings = src.DisableSingleQuoteStrings,
        ColorTokens       = new Dictionary<string, string>(src.ColorTokens),
        SourcePath        = src.SourcePath,
        IsDirectorySource = src.IsDirectorySource,
        InstalledOnUtc    = src.InstalledOnUtc,
        IconImage         = src.IconImage,
        IconBytes         = src.IconBytes,
    };

    // Loads a PNG from a stream and scales it to 48x48 if it is square,
    // otherwise returns null so the text fallback is used.
    private static byte[]? ReadIconBytesFromStream(Stream stream)
    {
        try
        {
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            return ms.ToArray();
        }
        catch { return null; }
    }

    // Must be called on the UI thread. Decodes raw PNG bytes into an Avalonia Bitmap
    // and validates that the image is square (non-square icons are rejected).
    private static Bitmap? DecodeBitmapOnUiThread(byte[]? iconBytes)
    {
        if (iconBytes is null) return null;
        try
        {
            using var ms = new MemoryStream(iconBytes);
            var bmp = new Bitmap(ms);
            if (bmp.PixelSize.Width != bmp.PixelSize.Height) return null;
            return bmp;
        }
        catch { return null; }
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
        baseExt.InstalledOnUtc = GetExtensionSourceActivityUtc(koxPath, isDirectory: false);

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
            baseExt.IconBytes = ReadIconBytesFromStream(iconStream);
        }

        var themeEntry = archive.GetEntry("theme.json");
        if (themeEntry is null)
        {
            yield return baseExt;
            yield break;
        }

        // ZipArchiveEntry streams are forward-only - read to memory first so we can
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
            Functions = ReadStringArray(lang, "functions"),
            Properties = ReadStringArray(lang, "properties"),
            Namespaces = ReadStringArray(lang, "namespaces"),
            CommentLine = lang.TryGetProperty("commentLine", out var cl) ? NormalizeSyntaxToken(cl.GetString()) : null,
            CommentBlockStart = lang.TryGetProperty("commentBlockStart", out var cbs) ? NormalizeSyntaxToken(cbs.GetString()) : null,
            CommentBlockEnd = lang.TryGetProperty("commentBlockEnd", out var cbe) ? NormalizeSyntaxToken(cbe.GetString()) : null,
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
        ext.Functions = ext.Functions.Union(profile.Functions).ToArray();
        ext.Properties = ext.Properties.Union(profile.Properties).ToArray();
        ext.Namespaces = ext.Namespaces.Union(profile.Namespaces).ToArray();

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

    private static string? NormalizeSyntaxToken(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;

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

        // Keep the active-theme dot in sync whenever brushes are refreshed.
        // This covers the case where a new LoadedExtension instance was added
        // after CurrentThemeName was last set (e.g. post-startup extension load).
        foreach (var ext in ThemeExtensions)
            ext.IsActiveTheme = string.Equals(ext.ThemeCardThemeId, _currentThemeName, StringComparison.OrdinalIgnoreCase);
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
        {
            // No extension matched - try to detect the language from file content.
            // We read only the first non-empty line so this stays cheap even for large files.
            extension = TryDetectLanguageFromContent(filePath);
            if (extension is null)
                return null;

            // Content-sniffed match: use the base extension as-is (no profile narrowing).
            return extension;
        }

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

    // Peeks at the first non-empty line of a file and tries to match it against known
    // XML/MSBuild root elements so that extensionless or ambiguous files (e.g. a bare
    // "Makefile" that is actually an MSBuild .proj) get syntax highlighting anyway.
    private LoadedExtension? TryDetectLanguageFromContent(string filePath)
    {
        try
        {
            string? firstLine = null;
            using (var reader = new StreamReader(filePath, detectEncodingFromByteOrderMarks: true))
            {
                string? line;
                while ((line = reader.ReadLine()) is not null)
                {
                    var trimmed = line.Trim();
                    if (trimmed.Length > 0)
                    {
                        firstLine = trimmed;
                        break;
                    }
                }
            }

            if (firstLine is null)
                return null;

            // Map root-element signatures to a representative extension that an installed
            // language extension already claims, so we reuse the full profile lookup.
            string? syntheticExt = null;

            if (firstLine.StartsWith("<Project", StringComparison.OrdinalIgnoreCase))
                syntheticExt = ".csproj";
            else if (firstLine.StartsWith("<?xml", StringComparison.OrdinalIgnoreCase) ||
                     firstLine.StartsWith("<", StringComparison.OrdinalIgnoreCase))
                syntheticExt = ".xml";

            if (syntheticExt is null)
                return null;

            return LoadedExtensions.FirstOrDefault(e =>
                e.Type == "language" &&
                e.Extensions.Any(ex => ex.Equals(syntheticExt, StringComparison.OrdinalIgnoreCase)));
        }
        catch
        {
            return null;
        }
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

    // Returns true when the byte content of a file indicates it is binary (non-text).
    // We sample up to 8 KB and treat the file as binary if it contains any null bytes,
    // which is the standard heuristic used by git and most editors.
    private static bool IsBinaryContent(string path)
    {
        try
        {
            const int sampleSize = 8192;
            Span<byte> buffer = stackalloc byte[sampleSize];
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var read = fs.Read(buffer);
            for (var i = 0; i < read; i++)
            {
                if (buffer[i] == 0x00)
                    return true;
            }
            return false;
        }
        catch
        {
            // If we can't read the file at all, treat it as corrupted.
            return true;
        }
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

    private bool IsSmartSyntaxEnabled() =>
        !IsPlainTextMode() &&
        HasFileOpen &&
        ActiveEditorTab is { IsUntitled: false };

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
        if (!_highlightingCache.TryGetValue(ext, out var definition))
        {
            definition = new KodoHighlightingDefinition(ext);
            _highlightingCache[ext] = definition;
        }
        EditorTextBox.SyntaxHighlighting = definition;
        ConfigureRainbowBrackets(ext);
        ConfigureInterpolatedStrings(ext);
        ConfigureMarkdownHighlighting(ext);
    }

    private void RefreshCurrentFileSyntaxHighlighting()
    {
        if (EditorTextBox is null)
            return;

        if (string.IsNullOrWhiteSpace(_currentFilePath))
        {
            CurrentLanguageExtension = null;
            ClearEditorSyntaxState();
            _indentGuideRenderer.IsEnabled = false;
            EditorTextBox.TextArea.TextView.InvalidateLayer(KnownLayer.Background);
            return;
        }

        var langExt = GetLanguageExtension(_currentFilePath);
        CurrentLanguageExtension = langExt;

        // Indent guides are only meaningful when a language extension is active.
        // Plain-text files (.txt / .log / .text) return null from GetLanguageExtension.
        _indentGuideRenderer.IsEnabled = langExt is not null;
        EditorTextBox.TextArea.TextView.InvalidateLayer(KnownLayer.Background);

        if (langExt is null)
        {
            ClearEditorSyntaxState();
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
            SetFileCorrupted(false);
            CurrentLanguageExtension = null;
            ClearEditorSyntaxState();
            return;
        }

        RefreshCurrentFileSyntaxHighlighting();
    }

    private void ClearEditorSyntaxState()
    {
        if (EditorTextBox is null)
            return;

        EditorTextBox.SyntaxHighlighting = null;
        ConfigureRainbowBrackets(null);
        ConfigureInterpolatedStrings(null);
        ConfigureMarkdownHighlighting(null);
    }

    private void ConfigureMarkdownHighlighting(LoadedExtension? extension)
    {
        _markdownColorizer.UpdateSyntax(extension, ResolveFenceLanguageExtension, ResolveInlineCodeLanguageExtension);
        EditorTextBox?.TextArea.TextView.InvalidateLayer(KnownLayer.Text);
    }

    private LoadedExtension? ResolveFenceLanguageExtension(string fenceLanguage)
    {
        if (string.IsNullOrWhiteSpace(fenceLanguage))
            return null;

        var token = fenceLanguage.Trim();
        if (token.StartsWith("{", StringComparison.Ordinal) && token.EndsWith("}", StringComparison.Ordinal) && token.Length > 2)
            token = token[1..^1];

        token = token.Split([' ', '\t', ',', ';', ':'], StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? token;
        if (string.IsNullOrWhiteSpace(token))
            return null;

        var normalized = token.Trim().TrimStart('.').ToLowerInvariant();
        if (FenceLanguageAliases.TryGetValue(normalized, out var alias))
            normalized = alias;

        var bestMatch = LoadedExtensions
            .Where(extension =>
                extension.Type == "language" &&
                !string.Equals(extension.Id, "markdown-kodo-extension", StringComparison.OrdinalIgnoreCase))
            .Select(extension => new
            {
                Extension = extension,
                Score = ScoreFenceLanguageMatch(extension, normalized, token)
            })
            .Where(result => result.Score > 0)
            .OrderByDescending(result => result.Score)
            .ThenBy(result => result.Extension.Name, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault();

        return bestMatch?.Extension;
    }

    private static int ScoreFenceLanguageMatch(LoadedExtension extension, string normalizedFenceLanguage, string rawFenceLanguage)
    {
        if (string.IsNullOrWhiteSpace(normalizedFenceLanguage))
            return 0;

        var score = 0;
        var normalizedCompact = normalizedFenceLanguage.Replace("-", string.Empty, StringComparison.OrdinalIgnoreCase);
        var rawCompact = rawFenceLanguage.Trim().TrimStart('.').Replace("-", string.Empty, StringComparison.OrdinalIgnoreCase);

        if (extension.Extensions.Any(ext =>
                ext.TrimStart('.').Equals(normalizedFenceLanguage, StringComparison.OrdinalIgnoreCase)))
        {
            score = Math.Max(score, 100);
        }

        var normalizedId = extension.Id
            .Replace("-kodo-extension", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("-language-support", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("-", string.Empty, StringComparison.OrdinalIgnoreCase);
        if (normalizedId.Equals(normalizedCompact, StringComparison.OrdinalIgnoreCase) ||
            normalizedId.Equals(rawCompact, StringComparison.OrdinalIgnoreCase))
        {
            score = Math.Max(score, 90);
        }

        var compactName = extension.Name
            .Replace("Language Support", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Support", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace(" ", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Trim();
        if (compactName.Equals(normalizedCompact, StringComparison.OrdinalIgnoreCase) ||
            compactName.Equals(rawCompact, StringComparison.OrdinalIgnoreCase))
        {
            score = Math.Max(score, 80);
        }

        if (extension.SyntaxProfiles.Any(profile =>
                profile.Extensions.Any(ext =>
                    ext.TrimStart('.').Equals(normalizedFenceLanguage, StringComparison.OrdinalIgnoreCase))))
        {
            score = Math.Max(score, 70);
        }

        if (extension.Types.Any(type =>
                type.Equals(rawFenceLanguage, StringComparison.OrdinalIgnoreCase) ||
                type.Equals(normalizedFenceLanguage, StringComparison.OrdinalIgnoreCase)))
        {
            score = Math.Max(score, 60);
        }

        if (extension.Keywords.Any(keyword =>
                keyword.Equals(rawFenceLanguage, StringComparison.OrdinalIgnoreCase) ||
                keyword.Equals(normalizedFenceLanguage, StringComparison.OrdinalIgnoreCase)))
        {
            score = Math.Max(score, 40);
        }

        return score;
    }

    private LoadedExtension? ResolveInlineCodeLanguageExtension(string codeSnippet)
    {
        if (string.IsNullOrWhiteSpace(codeSnippet))
            return null;

        var snippet = codeSnippet.Trim();
        if (snippet.Length < 2)
            return null;

        var bestMatch = LoadedExtensions
            .Where(extension =>
                extension.Type == "language" &&
                !string.Equals(extension.Id, "markdown-kodo-extension", StringComparison.OrdinalIgnoreCase))
            .Select(extension => new { Extension = extension, Score = ScoreInlineCodeLanguage(extension, snippet) })
            .Where(result => result.Score > 0)
            .OrderByDescending(result => result.Score)
            .ThenBy(result => result.Extension.Name, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault();

        return bestMatch?.Extension;
    }

    private static int ScoreInlineCodeLanguage(LoadedExtension extension, string snippet)
    {
        var score = 0;
        score += ScoreTokenMatches(snippet, extension.Keywords, 5);
        score += ScoreTokenMatches(snippet, extension.Types, 4);
        score += ScoreTokenMatches(snippet, extension.Functions, 4);
        score += ScoreTokenMatches(snippet, extension.Properties, 3);
        score += ScoreTokenMatches(snippet, extension.Namespaces, 3);

        if (!string.IsNullOrWhiteSpace(extension.CommentLine) &&
            snippet.Contains(extension.CommentLine, StringComparison.Ordinal))
        {
            score += 2;
        }

        if (extension.DisableSingleQuoteStrings && snippet.Contains("=>", StringComparison.Ordinal))
            score += 2;

        return score;
    }

    private static int ScoreTokenMatches(string snippet, IEnumerable<string> tokens, int weight)
    {
        var total = 0;

        foreach (var token in tokens.Where(token => !string.IsNullOrWhiteSpace(token)).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            var escaped = Regex.Escape(token);
            var regex = new Regex($@"(?<![\p{{L}}\p{{Nd}}_]){escaped}(?![\p{{L}}\p{{Nd}}_])", RegexOptions.IgnoreCase);
            if (regex.IsMatch(snippet))
                total += weight;
        }

        return total;
    }

    // Sets the corrupted/unsupported state and fires all dependent property notifications.
    private void SetFileCorrupted(bool corrupted)
    {
        if (_isFileCorrupted == corrupted) return;
        _isFileCorrupted = corrupted;
        OnPropertyChanged(nameof(IsCorruptedFileViewVisible));
        OnPropertyChanged(nameof(IsTextEditorVisible));
        OnPropertyChanged(nameof(CanShowFindInFile));
        OnPropertyChanged(nameof(IsFindPanelActive));
        OnPropertyChanged(nameof(CanShowSaveActions));
    }

    private void ConfigureRainbowBrackets(LoadedExtension? ext)
    {
        _rainbowBracketColorizer.UpdateSyntax(ext);
        EditorTextBox?.TextArea.TextView.InvalidateLayer(KnownLayer.Text);
    }

    private void ConfigureInterpolatedStrings(LoadedExtension? ext)
    {
        _interpolatedStringColorizer.UpdateSyntax(ext);
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

    public bool IsTutorialPageVisible
    {
        get => _isTutorialPageVisible;
        set
        {
            if (_isTutorialPageVisible == value) return;
            _isTutorialPageVisible = value;
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

    public bool IsEditorPageVisible => !IsSettingsPageVisible && !IsExtensionsPageVisible && !IsTutorialPageVisible && !IsWhatsNewPageVisible;

    public bool HasDocumentOpen => _currentFilePath is not null || _hasUntitledDocument;

    public bool HasOpenEditors => OpenTabs.Count > 0;

    public bool IsDocumentViewVisible => HasDocumentOpen && IsEditorPageVisible && !IsHomePageVisible;

    public bool HasImagePreview => CurrentImagePreview is not null;

    public bool IsImagePreviewVisible => IsDocumentViewVisible && HasImagePreview;

    public bool IsCorruptedFileViewVisible => IsDocumentViewVisible && !HasImagePreview && _isFileCorrupted;

    public bool IsTextEditorVisible => IsDocumentViewVisible && !HasImagePreview && !_isFileCorrupted;

    public bool CanShowFindInFile => IsTextEditorVisible;

    public bool IsFindPanelActive => IsFindPanelVisible && CanShowFindInFile;

    public bool IsEditorTabsVisible => OpenTabs.Count >= 1 && IsEditorPageVisible && !IsHomePageVisible;

    public bool CanShowSaveActions => IsTextEditorVisible;

    public string WhatsNewToggleText => IsWhatsNewExpanded ? "Hide release notes" : "Show release notes";

    public string WhatsNewToggleGlyph => IsWhatsNewExpanded ? "▾" : "▸";

    public bool HasFileOpen => _currentFilePath is not null;

    public bool IsFolderOpen => _currentFolderPath is not null;

    public bool IsEmptyStateVisible => IsHomePageVisible || (IsEditorPageVisible && !HasDocumentOpen);

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
            OnPropertyChanged(nameof(CanUpdateAllExtensions));
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
            OnPropertyChanged(nameof(SelectedExtensionSort));
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
            OnPropertyChanged(nameof(SelectedExtensionSort));
        }
    }

    public IReadOnlyList<string> ExtensionSortOptions { get; } =
    [
        ExtensionSortModes.Alphabetical,
        ExtensionSortModes.ReverseAlphabetical,
        ExtensionSortModes.RecentlyInstalled,
        ExtensionSortModes.UpdatesAvailable
    ];

    public string SelectedExtensionSort
    {
        get => IsMarketplaceTabSelected ? _selectedMarketplaceExtensionSort : _selectedInstalledExtensionSort;
        set
        {
            var normalized = string.IsNullOrWhiteSpace(value) ? ExtensionSortModes.Alphabetical : value;
            if (IsMarketplaceTabSelected)
            {
                if (string.Equals(_selectedMarketplaceExtensionSort, normalized, StringComparison.Ordinal))
                    return;

                _selectedMarketplaceExtensionSort = normalized;
            }
            else
            {
                if (string.Equals(_selectedInstalledExtensionSort, normalized, StringComparison.Ordinal))
                    return;

                _selectedInstalledExtensionSort = normalized;
            }

            OnPropertyChanged();
            NotifyExtensionFiltersChanged();
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

    // Core version check - dismissal has no effect here.
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

            // Same numeric version - compare by suffix priority
            return VersionPriority(LatestReleaseTag) > VersionPriority(CurrentAppVersion);
        }
    }

    // Banner visibility - collapses when dismissed, reappears if the app restarts.
    public bool IsAppUpdateAvailable => IsNewerVersionAvailable && !_updateBannerDismissed;
    public int AvailableExtensionUpdatesCount => MarketplaceExtensions.Count(e => e.IsUpdateAvailable);
    public bool IsExtensionUpdateBannerVisible => AvailableExtensionUpdatesCount > 0 && !_extensionUpdateBannerDismissed;
    public string ExtensionUpdatesBannerText =>
        $"{AvailableExtensionUpdatesCount} extension{(AvailableExtensionUpdatesCount == 1 ? string.Empty : "s")} {(AvailableExtensionUpdatesCount == 1 ? "has" : "have")} updates available";
    public bool CanUpdateAllExtensions =>
        !IsUpdatingAllExtensions &&
        !IsRefreshingExtensions &&
        !MarketplaceExtensions.Any(e => e.IsInstalling) &&
        MarketplaceExtensions.Any(e => e.IsUpdateAvailable && !e.IsInstalling);
    public string UpdateAllExtensionsButtonText =>
        IsUpdatingAllExtensions
            ? "Updating..."
            : AvailableExtensionUpdatesCount > 0
                ? $"Update All ({AvailableExtensionUpdatesCount})"
                : "Update All";

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

    // Pre-compiled regex patterns used by ConvertMarkdownToDisplayText - compiled
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

    public bool IsMarketplaceConnectivityWarningVisible
    {
        get => _isMarketplaceConnectivityWarningVisible;
        private set
        {
            if (_isMarketplaceConnectivityWarningVisible == value) return;
            _isMarketplaceConnectivityWarningVisible = value;
            OnPropertyChanged();
        }
    }

    // The welcome step (index 0) is not counted in "Step X of Y" - only the content steps are.
    public string TutorialStepLabel => $"Step {TutorialStepIndex} of {TutorialSteps.Length - 1}";

    public string TutorialProgressDotsText =>
        string.Join(" ", Enumerable.Range(0, TutorialSteps.Length).Select(index => index <= TutorialStepIndex ? "●" : "○"));

    public string TutorialSectionTitle => CurrentTutorialStep.SectionTitle;

    public string TutorialTitle => CurrentTutorialStep.Title;

    // Step 5 ("Set up") body contains "accent colour" - swap for US users.
    public string TutorialBody => IsAmericanEnglish
        ? CurrentTutorialStep.Body.Replace("accent colour", "accent color")
        : CurrentTutorialStep.Body;

    public string TutorialShortcutText => CurrentTutorialStep.Shortcut;

    // Step 4 ("Settings") SpotlightTitle is "Personalize the experience" in the static array.
    // Step 5 ("Set up") SpotlightTitle is "Why personalise?" in the static array.
    // Swap to the correct regional form for each.
    public string TutorialSpotlightTitle => TutorialStepIndex switch
    {
        4 => IsAmericanEnglish ? "Personalize the experience"  : "Personalise the experience",
        5 => IsAmericanEnglish ? "Why personalize?"            : "Why personalise?",
        _ => CurrentTutorialStep.SpotlightTitle,
    };

    public string TutorialHighlightOne => CurrentTutorialStep.HighlightOne;

    // Step 4 ("Settings") HighlightTwo contains "tab behavior" - swap for non-US users.
    public string TutorialHighlightTwo => (!IsAmericanEnglish && TutorialStepIndex == 4)
        ? CurrentTutorialStep.HighlightTwo.Replace("tab behavior", "tab behaviour")
        : CurrentTutorialStep.HighlightTwo;

    // Step 5 ("Set up") HighlightThree contains "accent colour" - swap for US users.
    public string TutorialHighlightThree => IsAmericanEnglish
        ? CurrentTutorialStep.HighlightThree.Replace("accent colour", "accent color")
        : CurrentTutorialStep.HighlightThree;

    public bool CanGoToPreviousTutorialStep => TutorialStepIndex > 0;

    // True only on the final "Set up Kodo" step so the AXAML can show
    // interactive personalisation controls instead of the spotlight text panel.
    public bool IsTutorialSetupStep => TutorialStepIndex == TutorialSteps.Length - 1;
    public bool IsNotTutorialSetupStep => !IsTutorialSetupStep;

    // True only on the very first "Welcome to Kodo!" splash step.
    public bool IsTutorialWelcomeStep => TutorialStepIndex == 0;
    public bool IsNotTutorialWelcomeStep => !IsTutorialWelcomeStep;

    // Show the "Tutorial" page header only when opened deliberately from Settings,
    // not on first-launch where the welcome splash is the first thing seen.
    public bool IsTutorialHeaderVisible => _tutorialOpenedFromSettings;

    public string TutorialPrimaryButtonText => TutorialStepIndex >= TutorialSteps.Length - 1 ? "Finish tutorial" : "Next";

    public int TutorialStepIndex
    {
        get => _tutorialStepIndex;
        private set
        {
            var clamped = Math.Clamp(value, 0, TutorialSteps.Length - 1);
            if (_tutorialStepIndex == clamped) return;
            _tutorialStepIndex = clamped;
            OnTutorialStepChanged();
        }
    }

    public string MarketplaceConnectivityMessage
    {
        get => _marketplaceConnectivityMessage;
        private set
        {
            if (string.Equals(_marketplaceConnectivityMessage, value, StringComparison.Ordinal)) return;
            _marketplaceConnectivityMessage = value;
            OnPropertyChanged();
        }
    }

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
            OnPropertyChanged(nameof(ActiveTerminalWorkingDirectory));
            OnPropertyChanged(nameof(ActiveTerminalFooterText));
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
            NotifyExtensionFiltersChanged();
        }
    }

    public bool IsUpdatingAllExtensions
    {
        get => _isUpdatingAllExtensions;
        private set
        {
            if (_isUpdatingAllExtensions == value) return;
            _isUpdatingAllExtensions = value;
            OnPropertyChanged();
            NotifyExtensionActionStateChanged();
        }
    }

    public IEnumerable<LoadedExtension> FilteredInstalledExtensions
    {
        get
        {
            var source = string.IsNullOrWhiteSpace(_extensionSearchText)
                ? VisibleLoadedExtensions
                : VisibleLoadedExtensions.Where(e =>
                    e.Name.Contains(_extensionSearchText, StringComparison.OrdinalIgnoreCase) ||
                    e.Description.Contains(_extensionSearchText, StringComparison.OrdinalIgnoreCase));
            return SortInstalledExtensions(source);
        }
    }

    public IEnumerable<MarketplaceExtension> FilteredMarketplaceExtensions
    {
        get
        {
            var source = string.IsNullOrWhiteSpace(_extensionSearchText)
                ? MarketplaceExtensions
                : (IEnumerable<MarketplaceExtension>)MarketplaceExtensions.Where(e =>
                    e.Name.Contains(_extensionSearchText, StringComparison.OrdinalIgnoreCase) ||
                    e.Description.Contains(_extensionSearchText, StringComparison.OrdinalIgnoreCase) ||
                    e.Author.Contains(_extensionSearchText, StringComparison.OrdinalIgnoreCase));
            return SortMarketplaceExtensions(source);
        }
    }

    private IEnumerable<LoadedExtension> SortInstalledExtensions(IEnumerable<LoadedExtension> source) =>
        SelectedExtensionSort switch
        {
            ExtensionSortModes.ReverseAlphabetical => source.OrderByDescending(e => e.Name, StringComparer.OrdinalIgnoreCase),
            ExtensionSortModes.RecentlyInstalled => source
                .OrderByDescending(e => e.InstalledOnUtc ?? DateTime.MinValue)
                .ThenBy(e => e.Name, StringComparer.OrdinalIgnoreCase),
            ExtensionSortModes.UpdatesAvailable => source
                .OrderByDescending(e => e.IsUpdateAvailable)
                .ThenBy(e => e.Name, StringComparer.OrdinalIgnoreCase),
            _ => source.OrderBy(e => e.Name, StringComparer.OrdinalIgnoreCase)
        };

    private IEnumerable<MarketplaceExtension> SortMarketplaceExtensions(IEnumerable<MarketplaceExtension> source) =>
        SelectedExtensionSort switch
        {
            ExtensionSortModes.ReverseAlphabetical => source.OrderByDescending(e => e.Name, StringComparer.OrdinalIgnoreCase),
            ExtensionSortModes.RecentlyInstalled => source
                .OrderByDescending(e => e.InstalledOnUtc.HasValue)
                .ThenByDescending(e => e.InstalledOnUtc ?? DateTime.MinValue)
                .ThenBy(e => e.Name, StringComparer.OrdinalIgnoreCase),
            ExtensionSortModes.UpdatesAvailable => source
                .OrderBy(GetMarketplaceUpdatePriority)
                .ThenBy(e => e.Name, StringComparer.OrdinalIgnoreCase),
            _ => source.OrderBy(e => e.Name, StringComparer.OrdinalIgnoreCase)
        };

    private static int GetMarketplaceUpdatePriority(MarketplaceExtension extension)
    {
        if (extension.IsUpdateAvailable)
            return 0;
        if (!extension.IsInstalled)
            return 1;
        return 2;
    }

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

    public bool IsTerminalVisible
    {
        get => _isTerminalVisible;
        set
        {
            if (_isTerminalVisible == value) return;
            _isTerminalVisible = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(TerminalStatusBarText));
            OnPropertyChanged(nameof(ActiveTerminalFooterText));
            SaveSettings();
            if (_isTerminalVisible)
                FocusActiveTerminal();
            else
                RefreshTerminalWindows();
        }
    }

    public double TerminalPanelHeight => 250;


    public TerminalSession? ActiveTerminalSession
    {
        get => _activeTerminalSession;
        private set
        {
            if (ReferenceEquals(_activeTerminalSession, value))
                return;

            // ── Save the outgoing session's screen buffer ─────────────────────
            // PseudoConsoleTerminal hosts exactly one ConPTY at a time. Switching
            // sessions always kills the outgoing process (Start() calls Stop()
            // internally), so we save the snapshot and mark it not-running now,
            // before that happens, so the session knows it needs a cold-start
            // when switched back to.
            if (_activeTerminalSession is not null)
            {
                if (TerminalHostControl.HasLiveProcess)
                    _activeTerminalSession.Snapshot = TerminalHostControl.SaveSnapshot();
                _activeTerminalSession.IsRunning = false;
                _activeTerminalSession.IsSelected = false;
            }

            _activeTerminalSession = value;

            if (_activeTerminalSession is not null)
                _activeTerminalSession.IsSelected = true;

            OnPropertyChanged();
            OnPropertyChanged(nameof(HasActiveTerminal));
            OnPropertyChanged(nameof(ActiveTerminalStatusText));
            OnPropertyChanged(nameof(ActiveTerminalWorkingDirectory));
            OnPropertyChanged(nameof(ActiveTerminalShellDisplayName));
            OnPropertyChanged(nameof(ActiveTerminalFooterText));
            OnPropertyChanged(nameof(TerminalStatusBarText));
            if (_activeTerminalSession is not null)
            {
                // Cold-start the incoming session's process. Start() internally calls
                // ResizeCells which allocates a fresh empty cell grid, so we must
                // restore the snapshot AFTER Start() — not before — otherwise the
                // newly allocated grid would immediately overwrite the restored buffer.
                var shell = AvailableTerminalShells.FirstOrDefault(s =>
                    string.Equals(s.Id, _activeTerminalSession.ShellId, StringComparison.OrdinalIgnoreCase))
                    ?? GetSelectedTerminalShellOrFallback();
                if (shell is not null)
                {
                    try
                    {
                        var hasSnapshot = _activeTerminalSession.Snapshot is not null;
                        TerminalHostControl.Start(shell.FileName, shell.Arguments,
                            _activeTerminalSession.WorkingDirectory,
                            suppressOutputUntilRestored: hasSnapshot);
                        _activeTerminalSession.IsRunning = true;
                        _activeTerminalSession.StatusText = "Ready";
                    }
                    catch (Exception ex)
                    {
                        _activeTerminalSession.IsRunning = false;
                        _activeTerminalSession.StatusText = $"Failed to start: {ex.Message}";
                    }
                }

                // Restore the saved screen buffer now that Start() has finished
                // initialising the cell grid. This makes the previous session output
                // visible immediately while the new shell process is still starting up.
                if (_activeTerminalSession.Snapshot is not null)
                    TerminalHostControl.RestoreSnapshot(_activeTerminalSession.Snapshot);
            }
            else
            {
                TerminalHostControl.Stop();
            }
            RefreshTerminalWindows();
            SaveSettings();
        }
    }

    public TerminalShellOption? SelectedTerminalShell
    {
        get => _selectedTerminalShell;
        set
        {
            if (ReferenceEquals(_selectedTerminalShell, value))
                return;

            _selectedTerminalShell = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(ActiveTerminalShellDisplayName));
            SaveSettings();
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
    public string FileSummaryText => IsTutorialPageVisible
        ? "Tutorial"
        : IsHomePageVisible
        ? "Home"
        : HasDocumentOpen
        ? $"{GetDocumentDisplayName()}{GetDocumentStatusSuffix()}"
        : "Editor";
    public bool IsTerminalSupported => _isTerminalSupported;
    public bool HasActiveTerminal => ActiveTerminalSession is not null;
    public int TerminalSessionCount => TerminalSessions.Count;
    public string ActiveTerminalStatusText => ActiveTerminalSession?.StatusText ?? (IsTerminalSupported ? "No active terminal" : "Windows only");
    public string ActiveTerminalWorkingDirectory
    {
        get
        {
            var fullPath = ActiveTerminalSession?.WorkingDirectory ?? ResolveTerminalWorkingDirectory();
            if (IsStatusBarFilePathVisible) return fullPath;
            var trimmed = fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return Path.GetFileName(trimmed) is { Length: > 0 } name ? name : fullPath;
        }
    }
    public string ActiveTerminalShellDisplayName => ActiveTerminalSession?.ShellDisplayName ?? SelectedTerminalShell?.DisplayName ?? "Terminal";
    public string ActiveTerminalFooterText => HasActiveTerminal
        ? $"{ActiveTerminalWorkingDirectory}  |  {ActiveTerminalStatusText}"
        : IsTerminalSupported ? "Choose a shell and open a terminal session." : "Embedded terminal is currently supported on Windows only.";
    public string TerminalStatusBarText => IsTerminalVisible
        ? $"Terminal ({TerminalSessionCount})"
        : TerminalSessionCount > 0 ? $"Show terminal ({TerminalSessionCount})" : "Terminal";

    public string FilePathText => IsTutorialPageVisible
        ? "Getting started with Kodo"
        : IsHomePageVisible
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

    // ── Personalization settings ─────────────────────────────────────────────

    /// <summary>
    /// ISO 3166-1 alpha-2 country code entered (or auto-detected) for the user.
    /// Used to pick region-appropriate holiday / long-weekend messages.
    /// </summary>
    public string UserCountry
    {
        get => _userCountry;
        set
        {
            if (_userCountry == value) return;
            _userCountry = value.ToUpperInvariant();
            _welcomeMessagesCache = null;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsAmericanEnglish));
            OnPropertyChanged(nameof(LabelAccentColour));
            OnPropertyChanged(nameof(LabelPersonalization));
            OnPropertyChanged(nameof(LabelPersonalizationDescription));
            OnPropertyChanged(nameof(TutorialSpotlightTitle));
            OnPropertyChanged(nameof(TutorialBody));
            OnPropertyChanged(nameof(TutorialHighlightOne));
            OnPropertyChanged(nameof(TutorialHighlightThree));
            SaveSettings();
        }
    }

    /// <summary>
    /// True when the user's country is US, driving American English spelling
    /// (e.g. "Color" / "Personalization") instead of the default British/Canadian spelling.
    /// </summary>
    public bool IsAmericanEnglish => _userCountry == "US";

    // Regional spelling labels - swap between American and British/Canadian English
    // based on the selected country. US gets "Color" / "Personalization" / "personalize";
    // everyone else gets "Colour" / "Personalization" - wait, "Personalisation" is the
    // British form but "Personalization" is widely accepted internationally; only
    // "colour" vs "color" and "personalise" vs "personalize" are the real visible splits.
    public string LabelAccentColour        => IsAmericanEnglish ? "Accent Color"      : "Accent Colour";
    public string LabelPersonalization     => IsAmericanEnglish ? "Personalization"   : "Personalisation";
    public string LabelPersonalizationDescription => IsAmericanEnglish
        ? "These settings personalize the welcome message on the Home screen. Your name is used in greetings when set. Country is auto-detected from your system if left blank. Hemisphere and time zone are also auto-detected when possible."
        : "These settings personalise the welcome message on the Home screen. Your name is used in greetings when set. Country is auto-detected from your system if left blank. Hemisphere and time zone are also auto-detected when possible.";

    /// <summary>
    /// Hemisphere override: 0 = auto-detect from country, 1 = northern, 2 = southern.
    /// Affects which season is inferred for welcome-message flavour text.
    /// </summary>
    public int UserHemisphereIndex
    {
        get => _userHemisphere;
        set
        {
            if (_userHemisphere == value) return;
            _userHemisphere = value;
            _welcomeMessagesCache = null;
            OnPropertyChanged();
            SaveSettings();
        }
    }

    /// <summary>
    /// UTC offset string entered by the user (e.g. "-5", "+1").
    /// When non-empty, overrides the system clock for time-of-day greetings.
    /// </summary>
    public string UserTimezoneOffset
    {
        get => _userTimezoneOffset;
        set
        {
            if (_userTimezoneOffset == value) return;
            _userTimezoneOffset = value;
            _welcomeMessagesCache = null;
            OnPropertyChanged();
            SaveSettings();
        }
    }

    /// <summary>
    /// Optional display name used to personalise greetings (e.g. "Good morning, Alex!").
    /// Empty means greetings won't include a name.
    /// </summary>
    public string UserName
    {
        get => _userName;
        set
        {
            var trimmed = value?.Trim() ?? string.Empty;
            if (_userName == trimmed) return;
            _userName = trimmed;
            _welcomeMessagesCache = null;
            OnPropertyChanged();
            SaveSettings();
        }
    }

    // ── Welcome message ──────────────────────────────────────────────────────

    /// <summary>
    /// Tries to infer the user's country from the OS regional settings.
    /// Returns an uppercase ISO 3166-1 alpha-2 code (e.g. "CA", "US") or an
    /// empty string when no reliable mapping is available.
    /// </summary>
    private static string DetectCountryCode()
    {
        try
        {
            var region = System.Globalization.RegionInfo.CurrentRegion;
            return region.TwoLetterISORegionName.ToUpperInvariant();
        }
        catch { return string.Empty; }
    }

    // ── Holiday / calendar helpers ────────────────────────────────────────────

    private record HolidayEntry(string Name, string? Greeting);

    /// <summary>
    /// Returns a <see cref="HolidayEntry"/> when <paramref name="date"/> is a
    /// public holiday (or its eve) that is relevant for <paramref name="country"/>.
    /// Returns <c>null</c> when no match is found.
    /// </summary>
    private static HolidayEntry? GetHolidayEntry(DateTime date, string country)
    {
        var m = date.Month;
        var d = date.Day;
        var dow = date.DayOfWeek;
        var y = date.Year;

        // ── Universal / very widely observed ──────────────────────────────────
        if (m == 1  && d == 1)  return new("New Year's Day", "Happy New Year!");
        if (m == 12 && d == 31) return new("New Year's Eve", "Happy New Year's Eve!");
        if (m == 12 && d == 25) return new("Christmas Day", "Merry Christmas!");
        if (m == 12 && d == 24) return new("Christmas Eve", "Happy Christmas Eve!");
        if (m == 12 && d == 26 && country is "CA" or "GB" or "AU" or "NZ" or "ZA")
            return new("Boxing Day", "Happy Boxing Day!");
        if (m == 12 && d == 26) return new("Kwanzaa", "Happy Kwanzaa!");
        if (m == 10 && d == 31) return new("Halloween", "Happy Halloween!");
        if (m == 2  && d == 14) return new("Valentine's Day", "Happy Valentine's Day!");
        if (m == 4  && d == 1)  return new("April Fools' Day", "Happy April Fools'! (Or is it?)");
        if (m == 3  && d == 8)  return new("International Women's Day", "Happy International Women's Day!");
        if (m == 4  && d == 22) return new("Earth Day", "Happy Earth Day!");
        if (m == 5  && d == 5)  return new("Cinco de Mayo", "¡Feliz Cinco de Mayo!");
        if (m == 6  && d == 5)  return new("World Environment Day", "Happy World Environment Day!");
        if (m == 9  && d == 21) return new("International Day of Peace", "Happy International Day of Peace.");
        if (m == 12 && d == 10) return new("International Human Rights Day", "Happy Human Rights Day.");

        // ── Mother's Day: second Sunday of May ────────────────────────────────
        if (m == 5 && dow == DayOfWeek.Sunday && d >= 8 && d <= 14)
            return new("Mother's Day", "Happy Mother's Day!");

        // ── Father's Day: third Sunday of June ────────────────────────────────
        if (m == 6 && dow == DayOfWeek.Sunday && d >= 15 && d <= 21)
            return new("Father's Day", "Happy Father's Day!");

        // ── Easter (Anonymous Gregorian algorithm) ────────────────────────────
        var easter = ComputeEaster(y);
        if (m == easter.Month && d == easter.Day)
            return new("Easter Sunday", "Happy Easter!");
        if (date == easter.AddDays(-2))
            return new("Good Friday", "Good Friday - enjoy the long weekend!");
        if (date == easter.AddDays(1) && country is "CA" or "GB" or "AU" or "NZ")
            return new("Easter Monday", "Happy Easter Monday!");

        // ── Lunar New Year (Chinese/Vietnamese/Korean) ────────────────────────
        if (LunarNewYear(y) is { } lny && m == lny.Month && d == lny.Day)
            return new("Lunar New Year", "Happy Lunar New Year!");

        // ── Holi (full moon of Phalguna) ──────────────────────────────────────
        if (HoliDate(y) is { } holi && m == holi.Month && d == holi.Day)
            return new("Holi", "Happy Holi!");

        // ── Vesak / Buddha Day (full moon of Vaisakha) ───────────────────────
        if (VesakDate(y) is { } vesak && m == vesak.Month && d == vesak.Day)
            return new("Vesak", "Happy Vesak!");

        // ── Eid al-Fitr (1 Shawwal) ───────────────────────────────────────────
        if (EidAlFitr(y) is { } eidFitr && m == eidFitr.Month && d == eidFitr.Day)
            return new("Eid al-Fitr", "Eid Mubarak!");

        // ── Eid al-Adha (10 Dhu al-Hijjah) ───────────────────────────────────
        if (EidAlAdha(y) is { } eidAdha && m == eidAdha.Month && d == eidAdha.Day)
            return new("Eid al-Adha", "Eid Mubarak!");

        // ── Rosh Hashanah (1 Tishrei) ─────────────────────────────────────────
        if (RoshHashanah(y) is { } rosh && m == rosh.Month && d == rosh.Day)
            return new("Rosh Hashanah", "Shana Tova! Happy New Year!");

        // ── Yom Kippur (10 Tishrei) ───────────────────────────────────────────
        if (YomKippur(y) is { } yk && m == yk.Month && d == yk.Day)
            return new("Yom Kippur", "G'mar Chatima Tova. Easy fast.");

        // ── Navratri / Sharad Navratri (day after new moon of Ashwin) ────────
        if (NavratriDate(y) is { } nav && m == nav.Month && d == nav.Day)
            return new("Navratri", "Happy Navratri!");

        // ── Diwali (new moon of Kartika) ──────────────────────────────────────
        if (DiwaliDate(y) is { } diwali && m == diwali.Month && d == diwali.Day)
            return new("Diwali", "Happy Diwali!");

        // ── Hanukkah (25 Kislev) ──────────────────────────────────────────────
        if (HanukkahDate(y) is { } hanukkah && m == hanukkah.Month && d == hanukkah.Day)
            return new("Hanukkah", "Happy Hanukkah!");

        // ── Canada ────────────────────────────────────────────────────────────
        if (country == "CA")
        {
            if (m == 7  && d == 1)  return new("Canada Day", "Happy Canada Day!");
            if (m == 11 && d == 11) return new("Remembrance Day", "Lest we forget. Happy coding.");
            // Victoria Day: last Monday before May 25
            if (m == 5 && dow == DayOfWeek.Monday && d >= 18 && d <= 24)
                return new("Victoria Day", "Happy Victoria Day! Enjoy the long weekend.");
            // Labour Day: first Monday of September
            if (m == 9 && dow == DayOfWeek.Monday && d <= 7)
                return new("Labour Day", "Happy Labour Day! Enjoy the long weekend.");
            // Thanksgiving: second Monday of October
            if (m == 10 && dow == DayOfWeek.Monday && d >= 8 && d <= 14)
                return new("Thanksgiving", "Happy Thanksgiving!");
            // Family Day: third Monday of February (most provinces)
            if (m == 2 && dow == DayOfWeek.Monday && d >= 15 && d <= 21)
                return new("Family Day", "Happy Family Day! Enjoy the long weekend.");
        }

        // ── United States ─────────────────────────────────────────────────────
        if (country == "US")
        {
            if (m == 7  && d == 4)  return new("Independence Day", "Happy Fourth of July!");
            if (m == 11 && d == 11) return new("Veterans Day", "Thank you to all veterans. Happy coding.");
            // Thanksgiving: fourth Thursday of November
            if (m == 11 && dow == DayOfWeek.Thursday && d >= 22 && d <= 28)
                return new("Thanksgiving", "Happy Thanksgiving! (and happy coding after dinner)");
            // Memorial Day: last Monday of May
            if (m == 5 && dow == DayOfWeek.Monday && d >= 25)
                return new("Memorial Day", "Happy Memorial Day! Enjoy the long weekend.");
            // Labor Day: first Monday of September
            if (m == 9 && dow == DayOfWeek.Monday && d <= 7)
                return new("Labor Day", "Happy Labor Day! Enjoy the long weekend.");
            // MLK Day: third Monday of January
            if (m == 1 && dow == DayOfWeek.Monday && d >= 15 && d <= 21)
                return new("MLK Day", "Happy Martin Luther King Jr. Day!");
            // Presidents' Day: third Monday of February
            if (m == 2 && dow == DayOfWeek.Monday && d >= 15 && d <= 21)
                return new("Presidents' Day", "Happy Presidents' Day! Enjoy the long weekend!");
        }

        // ── United Kingdom ────────────────────────────────────────────────────
        if (country == "GB")
        {
            if (m == 8 && dow == DayOfWeek.Monday && d >= 25)
                return new("August Bank Holiday", "Happy Bank Holiday! Enjoy the long weekend!");
            if (m == 5 && dow == DayOfWeek.Monday && d >= 1 && d <= 7)
                return new("Early May Bank Holiday", "Happy May Bank Holiday! Enjoy the long weekend!");
            if (m == 5 && dow == DayOfWeek.Monday && d >= 25)
                return new("Spring Bank Holiday", "Happy Spring Bank Holiday! Enjoy the long weekend!");
            if (m == 11 && d == 5)
                return new("Bonfire Night", "Remember, remember the 5th of November!");
        }

        // ── Australia ─────────────────────────────────────────────────────────
        if (country == "AU")
        {
            if (m == 1  && d == 26) return new("Australia Day", "Happy Australia Day!");
            if (m == 4  && d == 25) return new("ANZAC Day", "Lest we forget. Happy ANZAC Day.");
            if (m == 6  && dow == DayOfWeek.Monday && d >= 8 && d <= 14)
                return new("King's Birthday (AU)", "Happy King's Birthday long weekend!");
        }

        // ── New Zealand ───────────────────────────────────────────────────────
        if (country == "NZ")
        {
            if (m == 2  && d == 6)  return new("Waitangi Day", "Happy Waitangi Day!");
            if (m == 4  && d == 25) return new("ANZAC Day", "Lest we forget. Happy ANZAC Day.");
        }

        // ── Germany ───────────────────────────────────────────────────────────
        if (country == "DE")
        {
            if (m == 10 && d == 3) return new("German Unity Day", "Happy German Unity Day!");
            if (m == 5  && d == 1) return new("Labour Day", "Happy Labour Day!");
        }

        // ── France ────────────────────────────────────────────────────────────
        if (country == "FR")
        {
            if (m == 7  && d == 14) return new("Bastille Day", "Bonne fête nationale!");
            if (m == 5  && d == 1)  return new("Fête du Travail", "Bonne Fête du Travail!");
        }

        // ── Japan ─────────────────────────────────────────────────────────────
        if (country == "JP")
        {
            if (m == 1  && d == 1) return new("Shōgatsu", "あけましておめでとうございます！Happy New Year!");
            if (m == 11 && d == 3) return new("Culture Day", "Happy Culture Day!");
        }

        return null;
    }

    /// <summary>
    /// Computes Easter Sunday for a given year using the Anonymous Gregorian algorithm.
    /// </summary>
    private static DateTime ComputeEaster(int year)
    {
        int a = year % 19, b = year / 100, c = year % 100;
        int d2 = b / 4, e = b % 4, f = (b + 8) / 25;
        int g = (b - f + 1) / 3, h = (19 * a + b - d2 - g + 15) % 30;
        int i = c / 4, k = c % 4;
        int l = (32 + 2 * e + 2 * i - h - k) % 7;
        int m2 = (a + 11 * h + 22 * l) / 451;
        int month = (h + l - 7 * m2 + 114) / 31;
        int day = ((h + l - 7 * m2 + 114) % 31) + 1;
        return new DateTime(year, month, day);
    }

    // ── Astronomical calendar helpers ─────────────────────────────────────────
    // These compute floating-date holidays algorithmically so they remain correct
    // through 2100 without needing lookup tables.

    /// <summary>
    /// Julian Day Number of the kth new moon since J2000 (Meeus ch.49).
    /// Pass k+0.5 for the corresponding full moon.
    /// </summary>
    private static double MoonPhaseJdn(double k)
    {
        double T   = k / 1236.85;
        double jde = 2451550.09766
                   + 29.530588861 * k
                   + 0.00015437   * T * T
                   - 0.000000150  * T * T * T
                   + 0.00000000073 * T * T * T * T;
        double M  = Rad(2.5534   + 29.10535670  * k - 0.0000014 * T * T);
        double Mp = Rad(201.5643 + 385.81693528 * k + 0.0107582 * T * T);
        double F  = Rad(160.7108 + 390.67050284 * k - 0.0016118 * T * T);
        double Om = Rad(124.7746 -  1.56375588  * k + 0.0020672 * T * T);
        double E  = 1 - 0.002516 * T - 0.0000074 * T * T;
        return jde
            + (-0.40720 * Math.Sin(Mp))
            + ( 0.17241 * E * Math.Sin(M))
            + ( 0.01608 * Math.Sin(2 * Mp))
            + ( 0.01039 * Math.Sin(2 * F))
            + ( 0.00739 * E * Math.Sin(Mp - M))
            + (-0.00514 * E * Math.Sin(Mp + M))
            + ( 0.00208 * E * E * Math.Sin(2 * M))
            + (-0.00111 * Math.Sin(Mp - 2 * F))
            + (-0.00057 * Math.Sin(Mp + 2 * F))
            + ( 0.00056 * E * Math.Sin(2 * Mp + M))
            + (-0.00042 * Math.Sin(3 * Mp))
            + ( 0.00042 * E * Math.Sin(M + 2 * F))
            + ( 0.00038 * E * Math.Sin(M - 2 * F))
            + (-0.00024 * E * Math.Sin(2 * Mp - M))
            + (-0.00017 * Math.Sin(Om))
            + (-0.00007 * Math.Sin(Mp + 2 * M))
            + ( 0.00004 * Math.Sin(2 * Mp - 2 * F))
            + ( 0.00004 * Math.Sin(3 * M))
            + ( 0.00003 * Math.Sin(Mp + M - 2 * F))
            + ( 0.00003 * Math.Sin(2 * Mp + 2 * F))
            + (-0.00003 * Math.Sin(Mp + M + 2 * F))
            + ( 0.00003 * Math.Sin(Mp - M + 2 * F))
            + (-0.00002 * Math.Sin(Mp - M - 2 * F))
            + (-0.00002 * Math.Sin(3 * Mp + M))
            + ( 0.00002 * Math.Sin(4 * Mp));
    }

    private static double Rad(double deg) => deg * Math.PI / 180.0;

    /// <summary>Converts a Julian Day Number to a Gregorian DateTime (UTC noon).</summary>
    private static DateTime JdnToDateTime(double jdn)
    {
        int j = (int)(jdn + 0.5);
        int a = j + 32044;
        int b = (4 * a + 3) / 146097;
        int c = a - 146097 * b / 4;
        int d = (4 * c + 3) / 1461;
        int e = c - 1461 * d / 4;
        int mo = (5 * e + 2) / 153;
        int day   = e - (153 * mo + 2) / 5 + 1;
        int month = mo + 3 - 12 * (mo / 10);
        int year  = 100 * b + d - 4800 + mo / 10;
        return new DateTime(year, month, day);
    }

    /// <summary>
    /// Finds the new moon (or full moon when fullMoon=true) that falls in the
    /// given Gregorian year and month. Returns null if none falls in that month.
    /// </summary>
    private static DateTime? MoonInMonth(int year, int month, bool fullMoon = false)
    {
        double kApprox = (year - 2000) * 12.3685 + month - 1;
        for (int offset = -2; offset <= 3; offset++)
        {
            double k   = Math.Floor(kApprox) + offset + (fullMoon ? 0.5 : 0.0);
            double jdn = MoonPhaseJdn(k);
            var    dt  = JdnToDateTime(jdn);
            if (dt.Year == year && dt.Month == month)
                return dt;
        }
        return null;
    }

    /// <summary>
    /// Lunar New Year: the new moon that falls between Jan 20 and Feb 20
    /// (in China Standard Time, UTC+8). This is the first new moon after the
    /// Sun enters Aquarius (~Jan 20), which is the standard Chinese calendar rule.
    /// </summary>
    private static DateTime? LunarNewYear(int year)
    {
        double kApprox = (year - 2000) * 12.3685;
        for (int offset = -2; offset <= 3; offset++)
        {
            double k      = Math.Floor(kApprox) + offset;
            double jdn    = MoonPhaseJdn(k) + 8.0 / 24.0; // shift to UTC+8
            var    dt     = JdnToDateTime(jdn);
            if (dt.Year == year && ((dt.Month == 1 && dt.Day >= 20) || (dt.Month == 2 && dt.Day <= 20)))
                return dt;
        }
        return null;
    }

    /// <summary>
    /// Holi: the full moon of the Hindu month Phalguna, which falls in March
    /// (or occasionally very late February).
    /// </summary>
    private static DateTime? HoliDate(int year)
    {
        var march = MoonInMonth(year, 3, fullMoon: true);
        if (march != null) return march;
        var feb = MoonInMonth(year, 2, fullMoon: true);
        return feb?.Day >= 20 ? feb : null;
    }

    /// <summary>
    /// Vesak (Buddha Day): the full moon of the month of Vaisakha, observed
    /// on the full moon in May by Theravada countries.
    /// </summary>
    private static DateTime? VesakDate(int year) =>
        MoonInMonth(year, 5, fullMoon: true);

    /// <summary>
    /// Converts an Islamic (Hijri) civil date to a Gregorian DateTime,
    /// using the standard tabular (Kuwaiti algorithmic) calendar.
    /// Accurate to ±1 day vs actual moon-sighting dates.
    /// </summary>
    private static DateTime IslamicToGregorian(int iy, int im, int id)
    {
        int jdn = id
                + (int)Math.Ceiling(29.5 * (im - 1))
                + (iy - 1) * 354
                + (3 + 11 * iy) / 30
                + 1948438;
        return JdnToDateTime(jdn);
    }

    private static int ApproxHijriYear(int gregorianYear) =>
        (int)((gregorianYear - 622) * 1.030685);

    /// <summary>Eid al-Fitr: 1 Shawwal (Islamic month 10).</summary>
    private static DateTime? EidAlFitr(int year)
    {
        int hy = ApproxHijriYear(year);
        for (int h = hy - 1; h <= hy + 1; h++)
        {
            var dt = IslamicToGregorian(h, 10, 1);
            if (dt.Year == year) return dt;
        }
        return null;
    }

    /// <summary>Eid al-Adha: 10 Dhu al-Hijjah (Islamic month 12).</summary>
    private static DateTime? EidAlAdha(int year)
    {
        int hy = ApproxHijriYear(year);
        for (int h = hy - 1; h <= hy + 1; h++)
        {
            var dt = IslamicToGregorian(h, 12, 10);
            if (dt.Year == year) return dt;
        }
        return null;
    }

    // ── Hebrew calendar (Rosh Hashanah / Yom Kippur / Hanukkah) ──────────────
    // Uses the traditional molad-based calculation (Maimonides / standard
    // rabbinical algorithm). Exact for the proleptic Hebrew calendar.

    private static bool IsHebrewLeapYear(int hy) => (7 * hy + 1) % 19 < 7;

    private static int HebrewElapsedDays(int hy)
    {
        int monthsElapsed = 235 * ((hy - 1) / 19)
                          + 12  * ((hy - 1) % 19)
                          + (7  * ((hy - 1) % 19) + 1) / 19;
        int parts = 204 + 793 * (monthsElapsed % 1080);
        int hours = 5 + 12 * monthsElapsed + 793 * (monthsElapsed / 1080) + parts / 1080;
        int day   = 1 + 29 * monthsElapsed + hours / 24;
        int pMod  = 1080 * (hours % 24) + parts % 1080;

        int alt = day;
        if (pMod >= 19440
            || (day % 7 == 2 && pMod >= 9924  && !IsHebrewLeapYear(hy))
            || (day % 7 == 1 && pMod >= 16789 &&  IsHebrewLeapYear(hy - 1)))
            alt++;

        if (alt % 7 == 0 || alt % 7 == 3 || alt % 7 == 5) alt++;
        return alt;
    }

    private static int HebrewYearDays(int hy) =>
        HebrewElapsedDays(hy + 1) - HebrewElapsedDays(hy);

    private static int HebrewMonthLength(int hy, int hm)
    {
        int yd = HebrewYearDays(hy);
        // Cheshvan (2): 30 only in complete years
        if (hm == 2) return yd % 10 == 5 ? 30 : 29;
        // Kislev (3): 29 only in deficient years
        if (hm == 3) return yd % 10 == 3 ? 29 : 30;
        // Adar (6) is 30 days in leap years, 29 in regular
        if (hm == 6) return IsHebrewLeapYear(hy) ? 30 : 29;
        return hm is 1 or 5 or 7 or 10 or 12 ? 30 : 29;
    }

    /// <summary>
    /// Converts a Hebrew date to a Gregorian DateTime.
    /// Month numbering: Tishrei=1, Cheshvan=2, Kislev=3, Tevet=4, Shevat=5,
    /// Adar(I)=6, [AdarII=7 in leap years], Nisan=7(8), … Elul=12(13).
    /// </summary>
    private static DateTime HebrewToGregorian(int hy, int hm, int hd)
    {
        const int HebrewEpoch = 347997; // JDN of 1 Tishrei AM 1
        int elapsed = HebrewElapsedDays(hy);
        int doy = hd;
        for (int mo = 1; mo < hm; mo++)
            doy += HebrewMonthLength(hy, mo);
        return JdnToDateTime(HebrewEpoch + elapsed + doy - 1);
    }

    private static int ApproxHebrewYear(int gregorianYear) => gregorianYear + 3760;

    /// <summary>Rosh Hashanah: 1 Tishrei of the Hebrew year beginning in <paramref name="year"/>.</summary>
    private static DateTime? RoshHashanah(int year)
    {
        int hy0 = ApproxHebrewYear(year);
        for (int hy = hy0 - 1; hy <= hy0 + 1; hy++)
        {
            var dt = HebrewToGregorian(hy, 1, 1);
            if (dt.Year == year) return dt;
        }
        return null;
    }

    /// <summary>Yom Kippur: 10 Tishrei.</summary>
    private static DateTime? YomKippur(int year)
    {
        int hy0 = ApproxHebrewYear(year);
        for (int hy = hy0 - 1; hy <= hy0 + 1; hy++)
        {
            var dt = HebrewToGregorian(hy, 1, 10);
            if (dt.Year == year) return dt;
        }
        return null;
    }

    /// <summary>Hanukkah: 25 Kislev (first day/night).</summary>
    private static DateTime? HanukkahDate(int year)
    {
        // 25 Kislev of Hebrew year ~(Gregorian + 3761) falls in Nov/Dec.
        int hy0 = ApproxHebrewYear(year) + 1;
        for (int hy = hy0 - 1; hy <= hy0 + 1; hy++)
        {
            var dt = HebrewToGregorian(hy, 3, 25);
            if (dt.Year == year) return dt;
        }
        return null;
    }

    /// <summary>
    /// Diwali: the new moon (Amavasya) of Kartika, falling in October or
    /// early November.
    /// </summary>
    private static DateTime? DiwaliDate(int year)
    {
        // Try October new moon first; accept it if day >= 14 (Kartika new moon
        // is always in the second half of October or early November).
        var oct = MoonInMonth(year, 10, fullMoon: false);
        if (oct != null && oct.Value.Day >= 14) return oct;
        var nov = MoonInMonth(year, 11, fullMoon: false);
        if (nov != null && nov.Value.Day <= 15) return nov;
        return oct; // fallback
    }

    /// <summary>
    /// Sharad Navratri: begins on the day after the new moon of Ashwin
    /// (Shukla Pratipada), which falls in September or early October.
    /// </summary>
    private static DateTime? NavratriDate(int year)
    {
        // Ashwin new moon falls in Sep (day >= 15) or early Oct (day <= 10).
        var sep = MoonInMonth(year, 9, fullMoon: false);
        if (sep != null && sep.Value.Day >= 15) return sep.Value.AddDays(1);
        var oct = MoonInMonth(year, 10, fullMoon: false);
        if (oct != null && oct.Value.Day <= 10) return oct.Value.AddDays(1);
        return null;
    }

    /// <summary>
    /// Returns true when <paramref name="date"/> falls on a Saturday or Sunday
    /// or is sandwiched into a long weekend (i.e. Friday before a Monday holiday
    /// or Tuesday after a Monday holiday, etc.).
    /// </summary>
    private static bool IsWeekend(DateTime date) =>
        date.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday;

    /// <summary>
    /// Returns true when <paramref name="date"/> is a Friday before a long weekend
    /// (i.e. the following Monday is a public holiday for the given country).
    /// </summary>
    private static bool IsLongWeekendEve(DateTime date, string country)
    {
        if (date.DayOfWeek != DayOfWeek.Friday) return false;
        return GetHolidayEntry(date.AddDays(3), country) is not null;
    }

    /// <summary>
    /// Returns true when <paramref name="date"/> is the Tuesday after a long weekend
    /// (i.e. the preceding Monday was a public holiday).
    /// </summary>
    private static bool IsPostLongWeekend(DateTime date, string country)
    {
        if (date.DayOfWeek != DayOfWeek.Tuesday) return false;
        return GetHolidayEntry(date.AddDays(-1), country) is not null;
    }

    // ── Welcome message construction ──────────────────────────────────────────

    private string[] BuildWelcomeMessages()
    {
        // Resolve effective local time, honouring the user's timezone override when set.
        DateTime now;
        if (!string.IsNullOrWhiteSpace(_userTimezoneOffset) &&
            double.TryParse(_userTimezoneOffset.Replace("+", ""), out var offsetHours))
        {
            var offset = TimeSpan.FromHours(offsetHours);
            now = DateTime.UtcNow + offset;
        }
        else
        {
            now = DateTime.Now;
        }

        var tod     = TimeOfDay(now.Hour);
        var country = _userCountry;
        var dow     = now.DayOfWeek;
        var dayName = now.ToString("dddd");   // e.g. "Monday"

        var messages = new List<string>();

        // Personalised name prefix: when a name is set, prepend it to a subset
        // of the greeting pool so they feel personal without being repetitive.
        var name = _userName;
        if (!string.IsNullOrWhiteSpace(name))
        {
            var tod2 = TimeOfDay(now.Hour);

            messages.Add($"Good {tod2}, {name}!");
            messages.Add($"Hey {name}! Ready to build?");
            messages.Add($"Welcome back, {name}!");
            messages.Add($"Let's go, {name}!");

            // Additional personalised greetings
            messages.Add($"Great to see you again, {name}!");
            messages.Add($"Ready for another session, {name}?");
            messages.Add($"Time to be productive, {name}!");
            messages.Add($"Time to build something great, {name}!");
            messages.Add($"Locked in and ready, {name}?");
            messages.Add($"Good to have you back, {name}.");
            messages.Add($"Let's ship something great today, {name}!");
            messages.Add($"Your workspace is ready, {name}.");
        }

        // ── 1. Holiday / special day ──────────────────────────────────────────
        // Added multiple times so that on a special day the relevant greeting
        // has a meaningfully higher chance of being shown than any single
        // generic message, without making it a certainty.
        var holiday = GetHolidayEntry(now, country);
        if (holiday?.Greeting is not null)
        {
            messages.Add(holiday.Greeting);
            messages.Add(holiday.Greeting);
            messages.Add(holiday.Greeting);
        }

        // ── 2. Long weekend hints ─────────────────────────────────────────────
        // Also weighted up so long-weekend messages feel timely when applicable.
        if (IsLongWeekendEve(now, country))
        {
            messages.Add("Long weekend starts tomorrow - one more push!");
            messages.Add("Long weekend starts tomorrow - one more push!");
            messages.Add("Almost there! Long weekend is just around the corner.");
            messages.Add("Almost there! Long weekend is just around the corner.");
            messages.Add($"Happy {dayName}! The long weekend is almost here.");
            messages.Add($"Happy {dayName}! The long weekend is almost here.");
        }

        if (IsPostLongWeekend(now, country))
        {
            messages.Add("Back from the long weekend - fresh start!");
            messages.Add("Back from the long weekend - fresh start!");
            messages.Add("Post-long-weekend. Let's ease back in.");
            messages.Add("Post-long-weekend. Let's ease back in.");
            messages.Add("Hope the long weekend recharged you. Ready to build?");
            messages.Add("Hope the long weekend recharged you. Ready to build?");
        }

        // ── 3. Day-of-week personality ────────────────────────────────────────
        messages.Add(dow switch
        {
            DayOfWeek.Monday    => "Monday? Let's make it count.",
            DayOfWeek.Tuesday   => "Tuesday momentum - keep it going!",
            DayOfWeek.Wednesday => "Midweek check-in - still crushing it?",
            DayOfWeek.Thursday  => "Almost Friday - don't stop now!",
            DayOfWeek.Friday    => "Happy Friday! Let's finish strong.",
            DayOfWeek.Saturday  => "Coding on a Saturday - respect.",
            DayOfWeek.Sunday    => "Sunday coding session - the quiet grind.",
            _                   => $"Happy {dayName}!"
        });

        if (dow == DayOfWeek.Friday)
        {
            messages.Add("It's Friday — let's ship something before the weekend!");
            messages.Add("Friday energy. Let's make the most of it.");
        }
        if (dow is DayOfWeek.Saturday or DayOfWeek.Sunday)
        {
            messages.Add("Weekend warrior mode: activated.");
            messages.Add("No meetings on weekends. Just code.");
        }
        if (dow == DayOfWeek.Monday)
        {
            messages.Add("New week, new bugs to squash.");
            messages.Add("Monday's for the brave. Welcome back.");
        }

        // ── 4. Time-of-day flavour ────────────────────────────────────────────
        if (tod != "night")
        {
            messages.Add($"Good {tod}!");
            messages.Add($"Good {tod}, ready to build?");
            messages.Add($"Good {tod}, let's get to it!");
            messages.Add($"It's a great {tod} to code!");
        }

        if (tod == "morning")
        {
            messages.Add("Hey there, early bird!");
            messages.Add("Rise and shine, let's code!");
            messages.Add("Coffee in hand, let's ship something!");
            messages.Add("A fresh day, a fresh start.");
            messages.Add("Morning focus is unmatched.");
        }
        if (tod == "afternoon")
        {
            messages.Add("Afternoon grind — let's go!");
            messages.Add("Hope the day's treating you well!");
            messages.Add("Halfway through the day, keep it up!");
            messages.Add("Afternoon slump? Not here.");
            messages.Add("Post-lunch focus: activated.");
        }
        if (tod == "evening")
        {
            messages.Add("Fancy coding over a cup of tea?");
            messages.Add("Winding down or just getting started?");
            messages.Add("Evening sessions hit different.");
            messages.Add("The day's winding down, but the code isn't.");
        }
        if (tod == "night")
        {
            messages.Add("Hey there, night owl!");
            messages.Add("Burning the midnight oil?");
            messages.Add("The best code gets written at night.");
            messages.Add("Still at it? Respect.");
            messages.Add("Late night, great code.");
            messages.Add("The quieter the world, the clearer the code.");
            messages.Add("Dark outside, bright ideas inside.");
            messages.Add("Another late one? Worth it.");
        }

        // ── 5. Season-aware messages ──────────────────────────────────────────
        // Hemisphere: 0 = auto-detect from country, 1 = northern, 2 = southern.
        var isSouthern = _userHemisphere == 2
            || (_userHemisphere == 0 && country is "AU" or "NZ" or "ZA" or "AR" or "BR" or "CL");
        var month = now.Month;
        var season = isSouthern
            ? month switch { 12 or 1 or 2 => "summer", 3 or 4 or 5 => "autumn",
                             6 or 7 or 8 => "winter", _ => "spring" }
            : month switch { 12 or 1 or 2 => "winter", 3 or 4 or 5 => "spring",
                             6 or 7 or 8 => "summer", _ => "autumn" };

        messages.Add(season switch
        {
            "winter" => "Warm up your fingers - it's time to code.",
            "spring" => "Spring energy - let's build something fresh.",
            "summer" => "Hot outside, hotter code.",
            "autumn" => "Cozy season, perfect for shipping features.",
            _        => "Great day to write some code."
        });

        // ── 6. Neutral standby messages ───────────────────────────────────────
        // Excluded when a name is set so the pool isn't diluted by messages
        // that could have addressed the user by name instead.
        if (string.IsNullOrWhiteSpace(name))
        {
            messages.Add("Welcome back!");
            messages.Add("Great to see you!");
            messages.Add("Ready to code?");
            messages.Add("Let's build something!");
            messages.Add("What are we building today?");
            messages.Add("Back at it again!");
            messages.Add("Let's get to work!");
            messages.Add("Hey there!");
            messages.Add($"Happy {dayName}!");
        }

        return messages.ToArray();
    }

    private static string TimeOfDay(int hour)
    {
        if (hour < 6)  return "night";
        if (hour < 12) return "morning";
        if (hour < 17) return "afternoon";
        if (hour < 22) return "evening";
        return "night";
    }

    // Lazily constructed per-instance so it can incorporate the personalization settings
    // which are read from settings before DataContext is set.
    private string[]? _welcomeMessagesCache;

    // Evaluated once per launch: true one in a million times, showing the "Code fast. Stay light" tagline.
    private readonly bool _isTaglineGreeting = Random.Shared.Next(1_500) == 0;
    public bool IsTaglineGreeting => _isTaglineGreeting;

    public string WelcomeMessage
    {
        get
        {
            _welcomeMessagesCache ??= BuildWelcomeMessages();
            return _welcomeMessagesCache[Random.Shared.Next(_welcomeMessagesCache.Length)];
        }
    }

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

    // Short human-readable label for the encoding of the active file, shown in the status bar.
    public string EncodingDisplayText
    {
        get
        {
            if (!HasFileOpen) return string.Empty;
            var cp = _currentFileEncoding.CodePage;
            return cp switch
            {
                65001 => _currentFileEncoding is System.Text.UTF8Encoding u && u.GetPreamble().Length > 0
                             ? "UTF-8 BOM"
                             : "UTF-8",
                1200  => "UTF-16 LE",
                1201  => "UTF-16 BE",
                12000 => "UTF-32",
                20127 => "ASCII",
                _     => _currentFileEncoding.WebName.ToUpperInvariant(),
            };
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
        ? "File-backed tabs reopen on launch, and unsaved tabs ask for confirmation before closing."
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

    public string WordCountText
    {
        get => _wordCountText;
        private set
        {
            if (_wordCountText == value) return;
            _wordCountText = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsWordCountVisible));
        }
    }

    // Only visible for plain-text files (.txt / .text / .log) that are open in the editor.
    public bool IsWordCountVisible =>
        IsTextEditorVisible && IsPlainTextFile(_currentFilePath);

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

    // Black or white — whichever contrasts better against the current AccentBrush.
    // Used wherever text or icons sit directly on an AccentBrush background so that
    // dark accent colours (e.g. black, navy) don't make content invisible.
    public IBrush AccentForegroundBrush { get; private set; } = Brushes.White;

    // Returns Brushes.White or Brushes.Black depending on which gives better contrast
    // against the supplied brush, using the WCAG relative-luminance formula.
    private static IBrush GetAccentForeground(IBrush accent)
    {
        if (accent.ToImmutable() is not ISolidColorBrush solid)
            return Brushes.White;
        var c = solid.Color;
        // Convert sRGB channels to linear light
        static double Lin(byte channel)
        {
            var s = channel / 255.0;
            return s <= 0.04045 ? s / 12.92 : Math.Pow((s + 0.055) / 1.055, 2.4);
        }
        var L = 0.2126 * Lin(c.R) + 0.7152 * Lin(c.G) + 0.0722 * Lin(c.B);
        // White (L=1) contrast ratio vs L: (1.05)/(L+0.05)
        // Black (L=0) contrast ratio vs L: (L+0.05)/(0.05)
        // Pick whichever is higher
        return (1.05 / (L + 0.05)) >= ((L + 0.05) / 0.05)
            ? Brushes.White
            : Brushes.Black;
    }

    // Always reflects the live Windows accent colour; used by the Windows blob
    // preview even when another accent mode is active.
    public IBrush WindowsAccentPreviewBrush { get; private set; } =
        Brush.Parse("#0078D4");

    public string AccentColorMode
    {
        get => _accentColorMode;
        set
        {
            if (_accentColorMode == value) return;
            _accentColorMode = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsAccentKodo));
            OnPropertyChanged(nameof(IsAccentWindows));
            OnPropertyChanged(nameof(IsAccentCustom));
        }
    }
    public bool IsAccentKodo    => _accentColorMode == "kodo";
    public bool IsAccentWindows => _accentColorMode == "windows";
    public bool IsAccentCustom  => _accentColorMode == "custom";

    public string CustomAccentHex
    {
        get => _customAccentHex;
        set
        {
            if (_customAccentHex == value) return;
            _customAccentHex = value;
            OnPropertyChanged();
        }
    }

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

    private void FileTreeItems_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (!_suppressExplorerWidthRefresh)
            OnPropertyChanged(nameof(ExplorerPanelWidth));
    }

    private void TerminalSessions_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (e.NewItems is not null)
        {
            foreach (var item in e.NewItems.OfType<TerminalSession>())
                item.PropertyChanged += TerminalSession_OnPropertyChanged;
        }

        if (e.OldItems is not null)
        {
            foreach (var item in e.OldItems.OfType<TerminalSession>())
            {
                item.PropertyChanged -= TerminalSession_OnPropertyChanged;
                item.Dispose();
            }
        }

        OnPropertyChanged(nameof(HasActiveTerminal));
        OnPropertyChanged(nameof(TerminalSessionCount));
        OnPropertyChanged(nameof(TerminalStatusBarText));
        OnPropertyChanged(nameof(ActiveTerminalFooterText));
        RefreshTerminalWindows();
        SaveSettings();
    }

    private void TerminalSession_OnPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName is nameof(TerminalSession.WorkingDirectory) or nameof(TerminalSession.StatusText) or nameof(TerminalSession.WindowHandle))
        {
            OnPropertyChanged(nameof(ActiveTerminalWorkingDirectory));
            OnPropertyChanged(nameof(ActiveTerminalStatusText));
            OnPropertyChanged(nameof(ActiveTerminalFooterText));
            OnPropertyChanged(nameof(TerminalStatusBarText));
            RefreshTerminalWindows();
        }
    }

    // ── State management ─────────────────────────────────────────────────────

    private void RefreshState(bool fullRefresh = false)
    {
        RefreshCaretAndDocumentStats();

        if (!fullRefresh)
            return;

        _pendingFullStateRefresh = false;
        RefreshWordCount();
        RefreshNonCaretState();
    }

    private void RefreshCaretAndDocumentStats()
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
    }

    private void RefreshNonCaretState()
    {
        Title = HasDocumentOpen ? $"{GetDocumentDisplayName()} - Kodo" : "Kodo";
        OnPropertyChanged(nameof(HasDocumentOpen));
        OnPropertyChanged(nameof(IsDocumentViewVisible));
        OnPropertyChanged(nameof(HasImagePreview));
        OnPropertyChanged(nameof(IsImagePreviewVisible));
        OnPropertyChanged(nameof(IsTextEditorVisible));
        OnPropertyChanged(nameof(CanShowFindInFile));
        OnPropertyChanged(nameof(IsFindPanelActive));
        OnPropertyChanged(nameof(CanShowSaveActions));
        OnPropertyChanged(nameof(IsWordCountVisible));
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
        OnPropertyChanged(nameof(EncodingDisplayText));
        OnPropertyChanged(nameof(ActiveTerminalWorkingDirectory));
        OnPropertyChanged(nameof(ActiveTerminalFooterText));
        OnPropertyChanged(nameof(TerminalStatusBarText));
        UpdateDiscordPresence();
    }

    private void QueueRefreshState(bool fullRefresh = false)
    {
        _pendingFullStateRefresh |= fullRefresh;
        _editorStateRefreshTimer.Stop();
        _editorStateRefreshTimer.Start();
    }

    private void EditorStateRefreshTimer_OnTick(object? sender, EventArgs e)
    {
        _editorStateRefreshTimer.Stop();
        RefreshState(fullRefresh: _pendingFullStateRefresh);
    }

    private void QueueWordCountRefresh()
    {
        _wordCountRefreshTimer.Stop();
        _wordCountRefreshTimer.Start();
    }

    private void WordCountRefreshTimer_OnTick(object? sender, EventArgs e)
    {
        _wordCountRefreshTimer.Stop();
        RefreshWordCount();
        OnPropertyChanged(nameof(IsWordCountVisible));
    }

    private void RefreshWordCount()
    {
        if (!HasDocumentOpen || !IsPlainTextFile(_currentFilePath) || EditorTextBox?.Document is null)
        {
            WordCountText = string.Empty;
            return;
        }

        var text = EditorTextBox.Document.Text;
        var wordCount = string.IsNullOrWhiteSpace(text)
            ? 0
            : text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
        WordCountText = $"{wordCount} words";
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

    private void SaveSettings(bool immediate = false)
    {
        if (_suppressSettingsSave) return;

        if (!immediate)
        {
            _settingsSaveDebounceTimer.Stop();
            _settingsSaveDebounceTimer.Start();
            return;
        }

        _settingsSaveDebounceTimer.Stop();
        PersistSettingsSnapshot(BuildSettingsSnapshot());
    }

    private void SettingsSaveDebounceTimer_OnTick(object? sender, EventArgs e)
    {
        _settingsSaveDebounceTimer.Stop();
        if (_suppressSettingsSave)
            return;

        PersistSettingsSnapshot(BuildSettingsSnapshot());
    }

    private AppSettings BuildSettingsSnapshot()
    {
        // Snapshot all UI-thread-owned state here, before the background task,
        // so we don't access ObservableCollections or bound properties from a background thread.
        return new AppSettings
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
            PreferredTerminalShellId                = SelectedTerminalShell?.Id,
            TerminalVisible                         = IsTerminalVisible,
            TerminalPanelHeight                     = TerminalPanelHeight,
            HasCompletedTutorial                    = _hasCompletedTutorial,
            AccentColorMode                         = _accentColorMode,
            CustomAccentHex                         = _customAccentHex,
            UserCountry                             = _userCountry,
            UserHemisphere                          = _userHemisphere,
            UserTimezoneOffset                      = _userTimezoneOffset,
            UserName                                = _userName,
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
    }

    private void PersistSettingsSnapshot(AppSettings snapshot)
    {
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

    private IBrush GetCachedBrush(string colorValue)
    {
        if (_brushCache.TryGetValue(colorValue, out var brush))
            return brush;

        brush = Brush.Parse(colorValue);
        _brushCache[colorValue] = brush;
        return brush;
    }

    // ── Theme application ────────────────────────────────────────────────────

    /// <summary>
    /// Sets all theme brush fields and <see cref="Application.RequestedThemeVariant"/>
    /// without firing <see cref="INotifyPropertyChanged"/>, saving settings, or
    /// triggering a state refresh.  Call this <em>before</em> <c>DataContext = this</c>
    /// so that bindings read the correct colours on their very first evaluation and
    /// the window never renders with the wrong (default) palette.
    /// </summary>
    private void ApplyThemeBrushes(string themeName)
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

            WindowBackgroundBrush = GetCachedBrush(extensionTheme.WindowBackground);
            TopBarBrush           = GetCachedBrush(extensionTheme.TopBar);
            SidebarBrush          = GetCachedBrush(extensionTheme.Sidebar);
            ButtonBrush           = GetCachedBrush(extensionTheme.Button);
            ButtonHoverBrush      = GetCachedBrush(extensionTheme.ButtonHover);
            EditorBackgroundBrush = GetCachedBrush(extensionTheme.EditorBackground);
            CardBrush             = GetCachedBrush(extensionTheme.Card);
            PrimaryTextBrush      = GetCachedBrush(extensionTheme.PrimaryText);
            MutedTextBrush        = GetCachedBrush(extensionTheme.MutedText);
            SurfaceBorderBrush    = GetCachedBrush(extensionTheme.SurfaceBorder);
            AccentBrush           = GetCachedBrush(extensionTheme.Accent);
            _themeAccentHex       = extensionTheme.Accent;
        }
        else
        {
            CurrentThemeName = themeName == "Light" ? "Light" : "Dark";
            Application.Current!.RequestedThemeVariant = CurrentThemeName == "Light"
                ? ThemeVariant.Light
                : ThemeVariant.Dark;

            if (CurrentThemeName == "Light")
            {
                WindowBackgroundBrush = GetCachedBrush("#F3F3F3");
                TopBarBrush           = GetCachedBrush("#FFFFFF");
                SidebarBrush          = GetCachedBrush("#EFF2F7");
                ButtonBrush           = GetCachedBrush("#E3E8F1");
                ButtonHoverBrush      = GetCachedBrush("#D5DDE9");
                EditorBackgroundBrush = GetCachedBrush("#FFFFFF");
                CardBrush             = GetCachedBrush("#F7F9FC");
                PrimaryTextBrush      = GetCachedBrush("#202124");
                MutedTextBrush        = GetCachedBrush("#5F6B7A");
                SurfaceBorderBrush    = GetCachedBrush("#D7DCE5");
                AccentBrush           = GetCachedBrush("#8C00FF");
                _themeAccentHex       = "#8C00FF";
            }
            else
            {
                WindowBackgroundBrush = GetCachedBrush("#1E1E1E");
                TopBarBrush           = GetCachedBrush("#181818");
                SidebarBrush          = GetCachedBrush("#181818");
                ButtonBrush           = GetCachedBrush("#252526");
                ButtonHoverBrush      = GetCachedBrush("#313437");
                EditorBackgroundBrush = GetCachedBrush("#1E1E1E");
                CardBrush             = GetCachedBrush("#252526");
                PrimaryTextBrush      = GetCachedBrush("#F4F4F4");
                MutedTextBrush        = GetCachedBrush("#A0A0A0");
                SurfaceBorderBrush    = GetCachedBrush("#2B2B2B");
                AccentBrush           = GetCachedBrush("#8C00FF");
                _themeAccentHex       = "#8C00FF";
            }
        }

        // Apply the accent override (kodo / windows / custom) silently.
        // We can't call the full ApplyAccentOverride() here because it calls
        // ApplyThemeToEditor(), which touches the AvaloniaEdit TextEditor before
        // it has been fully laid out, but we CAN silently resolve the accent hex
        // so AccentBrush is correct when bindings first read it.
        //
        // WindowsAccentPreviewBrush must also be initialised here from the live
        // registry value. Without this it stays at its field-initialiser default
        // (#0078D4) until the poll timer fires ~2 s later, causing the Windows
        // blob in Settings to flash the wrong colour on every launch.
        var windowsHex = GetWindowsAccentColor() ?? "#0078D4";
        try { WindowsAccentPreviewBrush = GetCachedBrush(windowsHex); }
        catch { WindowsAccentPreviewBrush = GetCachedBrush("#0078D4"); }

        var resolvedAccent = _accentColorMode switch
        {
            "windows" => windowsHex,
            "custom"  => _customAccentHex,
            _         => _themeAccentHex,  // "kodo" - use the theme's own accent
        };
        try { AccentBrush = GetCachedBrush(resolvedAccent); }
        catch { AccentBrush = GetCachedBrush("#8C00FF"); }
        AccentForegroundBrush = GetAccentForeground(AccentBrush);
    }

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

            WindowBackgroundBrush = GetCachedBrush(extensionTheme.WindowBackground);
            TopBarBrush           = GetCachedBrush(extensionTheme.TopBar);
            SidebarBrush          = GetCachedBrush(extensionTheme.Sidebar);
            ButtonBrush           = GetCachedBrush(extensionTheme.Button);
            ButtonHoverBrush      = GetCachedBrush(extensionTheme.ButtonHover);
            EditorBackgroundBrush = GetCachedBrush(extensionTheme.EditorBackground);
            CardBrush             = GetCachedBrush(extensionTheme.Card);
            PrimaryTextBrush      = GetCachedBrush(extensionTheme.PrimaryText);
            MutedTextBrush        = GetCachedBrush(extensionTheme.MutedText);
            SurfaceBorderBrush    = GetCachedBrush(extensionTheme.SurfaceBorder);
            AccentBrush           = GetCachedBrush(extensionTheme.Accent);
            _themeAccentHex       = extensionTheme.Accent;
        }
        else
        {
            CurrentThemeName = themeName == "Light" ? "Light" : "Dark";
            Application.Current!.RequestedThemeVariant = CurrentThemeName == "Light"
                ? ThemeVariant.Light
                : ThemeVariant.Dark;

            if (CurrentThemeName == "Light")
            {
                WindowBackgroundBrush = GetCachedBrush("#F3F3F3");
                TopBarBrush           = GetCachedBrush("#FFFFFF");
                SidebarBrush          = GetCachedBrush("#EFF2F7");
                ButtonBrush           = GetCachedBrush("#E3E8F1");
                ButtonHoverBrush      = GetCachedBrush("#D5DDE9");
                EditorBackgroundBrush = GetCachedBrush("#FFFFFF");
                CardBrush             = GetCachedBrush("#F7F9FC");
                PrimaryTextBrush      = GetCachedBrush("#202124");
                MutedTextBrush        = GetCachedBrush("#5F6B7A");
                SurfaceBorderBrush    = GetCachedBrush("#D7DCE5");
                AccentBrush           = GetCachedBrush("#8C00FF");
                _themeAccentHex       = "#8C00FF";
            }
            else
            {
                WindowBackgroundBrush = GetCachedBrush("#1E1E1E");
                TopBarBrush           = GetCachedBrush("#181818");
                SidebarBrush          = GetCachedBrush("#181818");
                ButtonBrush           = GetCachedBrush("#252526");
                ButtonHoverBrush      = GetCachedBrush("#313437");
                EditorBackgroundBrush = GetCachedBrush("#1E1E1E");
                CardBrush             = GetCachedBrush("#252526");
                PrimaryTextBrush      = GetCachedBrush("#F4F4F4");
                MutedTextBrush        = GetCachedBrush("#A0A0A0");
                SurfaceBorderBrush    = GetCachedBrush("#2B2B2B");
                AccentBrush           = GetCachedBrush("#8C00FF");
                _themeAccentHex       = "#8C00FF";
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
        // Always run ApplyAccentOverride: it updates both AccentBrush (for all
        // three modes) and WindowsAccentPreviewBrush (so the blob stays live
        // even when "kodo" or "custom" mode is active).
        ApplyAccentOverride();
        ApplyThemeToEditor();
        SaveSettings();
        RefreshState(fullRefresh: true);
        RefreshExtensionTheme();
    }

    private void ApplyAccentOverride()
    {
        // Always keep the Windows preview brush current so the blob reflects
        // the real system colour regardless of which mode is active.
        var windowsHex = GetWindowsAccentColor() ?? "#0078D4";
        try { WindowsAccentPreviewBrush = GetCachedBrush(windowsHex); }
        catch { WindowsAccentPreviewBrush = GetCachedBrush("#0078D4"); }
        OnPropertyChanged(nameof(WindowsAccentPreviewBrush));

        // In "kodo" mode, restore the theme's own accent colour.
        if (_accentColorMode == "kodo")
        {
            try { AccentBrush = GetCachedBrush(_themeAccentHex); }
            catch { AccentBrush = GetCachedBrush("#8C00FF"); }
            AccentForegroundBrush = GetAccentForeground(AccentBrush);
            OnPropertyChanged(nameof(AccentBrush));
            OnPropertyChanged(nameof(AccentForegroundBrush));
            ApplyThemeToEditor();
            return;
        }

        var hex = _accentColorMode switch
        {
            "windows" => windowsHex,
            "custom"  => _customAccentHex,
            _         => "#8C00FF"
        };
        try { AccentBrush = GetCachedBrush(hex); }
        catch { AccentBrush = GetCachedBrush("#8C00FF"); }
        AccentForegroundBrush = GetAccentForeground(AccentBrush);
        OnPropertyChanged(nameof(AccentBrush));
        OnPropertyChanged(nameof(AccentForegroundBrush));
        ApplyThemeToEditor();
    }

    private static string? GetWindowsAccentColor()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return null;
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Explorer\Accent");
            if (key?.GetValue("AccentColorMenu") is int raw)
            {
                // AccentColorMenu is stored as AABBGGRR
                var r = (raw)       & 0xFF;
                var g = (raw >> 8)  & 0xFF;
                var b = (raw >> 16) & 0xFF;
                return $"#{r:X2}{g:X2}{b:X2}";
            }
        }
        catch { /* Registry unavailable */ }
        return null;
    }

    private async void AccentColorPickerButton_OnClick(object? sender, RoutedEventArgs e)
    {
        Window? dialog = null;
        var confirmed = false;

        var initialColor = Color.Parse("#8C00FF");
        try { initialColor = Color.Parse(_customAccentHex); } catch { /* use fallback */ }

        RgbToHsv(initialColor.R, initialColor.G, initialColor.B,
            out var hue, out var sat, out var val);

        // ── Hue strip ──────────────────────────────────────────────────────────
        var hueCanvas = new Canvas { Width = 300, Height = 20 };
        var hueGrad   = new LinearGradientBrush
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint   = new RelativePoint(1, 0, RelativeUnit.Relative),
        };
        foreach (var (offset, h) in new (double, double)[]
            { (0,0),(1/6d,60),(2/6d,120),(3/6d,180),(4/6d,240),(5/6d,300),(1,360) })
        {
            HsvToRgb(h, 1, 1, out var hr2, out var hg2, out var hb2);
            hueGrad.GradientStops.Add(new GradientStop(Color.FromRgb(hr2, hg2, hb2), offset));
        }
        var hueRect = new Avalonia.Controls.Shapes.Rectangle
            { Width = 300, Height = 20, Fill = hueGrad, RadiusX = 4, RadiusY = 4 };
        hueCanvas.Children.Add(hueRect);

        var hueCursor = new Avalonia.Controls.Shapes.Rectangle
        {
            Width = 4, Height = 24, Fill = Brushes.White, RadiusX = 2, RadiusY = 2,
            Stroke = new SolidColorBrush(Colors.Black), StrokeThickness = 1,
        };
        Canvas.SetTop(hueCursor, -2);
        Canvas.SetLeft(hueCursor, hue / 360.0 * 296);
        hueCanvas.Children.Add(hueCursor);

        // ── SV square ──────────────────────────────────────────────────────────
        const double svSize   = 300.0;
        const double svHeight = 180.0;
        var svCanvas = new Canvas { Width = svSize, Height = svHeight };

        var svHueFill      = new Avalonia.Controls.Shapes.Rectangle { Width = svSize, Height = svHeight, RadiusX = 4, RadiusY = 4 };
        var svWhiteOverlay = new Avalonia.Controls.Shapes.Rectangle { Width = svSize, Height = svHeight, RadiusX = 4, RadiusY = 4 };
        var svBlackOverlay = new Avalonia.Controls.Shapes.Rectangle
        {
            Width = svSize, Height = svHeight, RadiusX = 4, RadiusY = 4,
            Fill  = new LinearGradientBrush
            {
                StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
                EndPoint   = new RelativePoint(0, 1, RelativeUnit.Relative),
                GradientStops =
                {
                    new GradientStop(Color.FromArgb(0,   0, 0, 0), 0),
                    new GradientStop(Color.FromArgb(255, 0, 0, 0), 1),
                },
            },
        };

        void RefreshSvSquare()
        {
            HsvToRgb(hue, 1, 1, out var hr3, out var hg3, out var hb3);
            svHueFill.Fill = new SolidColorBrush(Color.FromRgb(hr3, hg3, hb3));
            svWhiteOverlay.Fill = new LinearGradientBrush
            {
                StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
                EndPoint   = new RelativePoint(1, 0, RelativeUnit.Relative),
                GradientStops =
                {
                    new GradientStop(Color.FromArgb(255, 255, 255, 255), 0),
                    new GradientStop(Color.FromArgb(0,   255, 255, 255), 1),
                },
            };
        }
        RefreshSvSquare();

        svCanvas.Children.Add(svHueFill);
        svCanvas.Children.Add(svWhiteOverlay);
        svCanvas.Children.Add(svBlackOverlay);

        var svCursor = new Avalonia.Controls.Shapes.Ellipse
        {
            Width = 12, Height = 12,
            Stroke = Brushes.White, StrokeThickness = 2,
            Fill   = new SolidColorBrush(Colors.Transparent),
        };
        Canvas.SetLeft(svCursor, sat * svSize   - 6);
        Canvas.SetTop (svCursor, (1 - val) * svHeight - 6);
        svCanvas.Children.Add(svCursor);

        // ── Preview swatch + hex input ─────────────────────────────────────────
        var previewBorder = new Border
        {
            Width = 36, Height = 36, CornerRadius = new CornerRadius(8),
            BorderBrush = SurfaceBorderBrush, BorderThickness = new Thickness(1),
        };

        var hexInput = new TextBox
        {
            Text            = _customAccentHex,
            PlaceholderText = "#RRGGBB",
            MaxLength       = 7,
            Foreground      = PrimaryTextBrush,
            Background      = ButtonBrush,
            BorderBrush     = SurfaceBorderBrush,
            BorderThickness = new Thickness(1),
            Padding         = new Thickness(8, 6),
            FontSize        = 13,
            CaretBrush      = PrimaryTextBrush,
            Width           = 110,
        };

        // ── Sync helpers ───────────────────────────────────────────────────────
        void UpdateAll()
        {
            HsvToRgb(hue, sat, val, out var r2, out var g2, out var b2);
            var c = Color.FromRgb(r2, g2, b2);
            previewBorder.Background = new SolidColorBrush(c);
            hexInput.Text = $"#{r2:X2}{g2:X2}{b2:X2}";
            Canvas.SetLeft(hueCursor, hue / 360.0 * 296);
            Canvas.SetLeft(svCursor,  sat * svSize   - 6);
            Canvas.SetTop (svCursor,  (1 - val) * svHeight - 6);
            RefreshSvSquare();
        }
        UpdateAll();

        // ── Hue drag ──────────────────────────────────────────────────────────
        hueCanvas.PointerPressed  += (_, pe) =>
        {
            pe.Pointer.Capture(hueCanvas);
            hue = Math.Clamp(pe.GetPosition(hueCanvas).X / 300.0 * 360, 0, 360);
            UpdateAll();
        };
        hueCanvas.PointerMoved    += (_, pe) =>
        {
            if (pe.Pointer.Captured != hueCanvas) return;
            hue = Math.Clamp(pe.GetPosition(hueCanvas).X / 300.0 * 360, 0, 360);
            UpdateAll();
        };
        hueCanvas.PointerReleased += (_, pe) => pe.Pointer.Capture(null);

        // ── SV drag ───────────────────────────────────────────────────────────
        svCanvas.PointerPressed  += (_, pe) =>
        {
            pe.Pointer.Capture(svCanvas);
            var p = pe.GetPosition(svCanvas);
            sat = Math.Clamp(p.X / svSize,        0, 1);
            val = Math.Clamp(1 - p.Y / svHeight,  0, 1);
            UpdateAll();
        };
        svCanvas.PointerMoved    += (_, pe) =>
        {
            if (pe.Pointer.Captured != svCanvas) return;
            var p = pe.GetPosition(svCanvas);
            sat = Math.Clamp(p.X / svSize,        0, 1);
            val = Math.Clamp(1 - p.Y / svHeight,  0, 1);
            UpdateAll();
        };
        svCanvas.PointerReleased += (_, pe) => pe.Pointer.Capture(null);

        // ── Hex sync ──────────────────────────────────────────────────────────
        hexInput.TextChanged += (_, _) =>
        {
            try
            {
                var t = hexInput.Text?.Trim() ?? "";
                if (!t.StartsWith('#')) t = "#" + t;
                var c = Color.Parse(t);
                RgbToHsv(c.R, c.G, c.B, out hue, out sat, out val);
                previewBorder.Background = new SolidColorBrush(c);
                Canvas.SetLeft(hueCursor, hue / 360.0 * 296);
                Canvas.SetLeft(svCursor,  sat * svSize   - 6);
                Canvas.SetTop (svCursor,  (1 - val) * svHeight - 6);
                RefreshSvSquare();
            }
            catch { /* wait for valid hex */ }
        };

        hexInput.KeyDown += (_, ke) =>
        {
            if (ke.Key == Key.Enter)  { confirmed = true; dialog!.Close(); }
            if (ke.Key == Key.Escape) { dialog!.Close(); }
        };

        dialog = new Window
        {
            Width                 = 340,
            SizeToContent         = SizeToContent.Height,
            CanResize             = false,
            ShowInTaskbar         = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Title                 = "Custom Accent Colour",
            Background            = CardBrush,
            Content = new Border
            {
                Padding = new Thickness(20),
                Child   = new StackPanel
                {
                    Spacing  = 14,
                    Children =
                    {
                        new TextBlock { Text = IsAmericanEnglish ? "Choose an accent color" : "Choose an accent colour", FontSize = 15,
                            FontWeight = FontWeight.SemiBold, Foreground = PrimaryTextBrush },
                        svCanvas,
                        hueCanvas,
                        new StackPanel
                        {
                            Orientation = Orientation.Horizontal,
                            Spacing     = 10,
                            Children    = { previewBorder, hexInput },
                        },
                        new StackPanel
                        {
                            Orientation         = Orientation.Horizontal,
                            Spacing             = 10,
                            HorizontalAlignment = HorizontalAlignment.Right,
                            Children =
                            {
                                CreateDialogButton("Cancel", ButtonBrush, SurfaceBorderBrush, PrimaryTextBrush,
                                    () => dialog!.Close()),
                                CreateDialogButton("Apply", AccentBrush, AccentBrush, AccentForegroundBrush,
                                    () => { confirmed = true; dialog!.Close(); }),
                            }
                        }
                    }
                }
            }
        };

        dialog.Opened += (_, _) => { hexInput.Focus(); hexInput.SelectAll(); };
        await dialog.ShowDialog(this);

        if (!confirmed) return;
        var hex = hexInput.Text?.Trim() ?? string.Empty;
        if (!hex.StartsWith('#')) hex = "#" + hex;
        try
        {
            Brush.Parse(hex);
            _customAccentHex = hex;
            CustomAccentHex  = hex;
            AccentColorMode  = "custom";
            ApplyAccentOverride();
            SaveSettings();
        }
        catch { /* invalid hex - ignore */ }
    }

    // Converts RGB (0–255) to HSV (H: 0–360, S/V: 0–1).
    private static void RgbToHsv(byte r, byte g, byte b,
        out double h, out double s, out double v)
    {
        var rf = r / 255.0; var gf = g / 255.0; var bf = b / 255.0;
        var max = Math.Max(rf, Math.Max(gf, bf));
        var min = Math.Min(rf, Math.Min(gf, bf));
        var delta = max - min;
        v = max;
        s = max == 0 ? 0 : delta / max;
        if (delta == 0) { h = 0; return; }
        if      (max == rf) h = 60 * (((gf - bf) / delta) % 6);
        else if (max == gf) h = 60 * (((bf - rf) / delta) + 2);
        else                h = 60 * (((rf - gf) / delta) + 4);
        if (h < 0) h += 360;
    }

    // Converts HSV (H: 0–360, S/V: 0–1) to RGB (0–255).
    private static void HsvToRgb(double h, double s, double v,
        out byte r, out byte g, out byte b)
    {
        if (s == 0) { r = g = b = (byte)(v * 255); return; }
        var i = (int)(h / 60) % 6;
        var f = h / 60 - Math.Floor(h / 60);
        var p = v * (1 - s); var q = v * (1 - f * s); var t = v * (1 - (1 - f) * s);
        var (rf, gf, bf) = i switch
        {
            0 => (v, t, p), 1 => (q, v, p), 2 => (p, v, t),
            3 => (p, q, v), 4 => (t, p, v), _ => (v, p, q),
        };
        r = (byte)(rf * 255); g = (byte)(gf * 255); b = (byte)(bf * 255);
    }

    private void AccentKodoButton_OnClick(object? sender, RoutedEventArgs e)
    {
        AccentColorMode = "kodo";
        ApplyAccentOverride();
        SaveSettings();
    }

    private void AccentWindowsButton_OnClick(object? sender, RoutedEventArgs e)
    {
        AccentColorMode = "windows";
        ApplyAccentOverride();
        SaveSettings();
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
        ActiveEditorTab.TopLineNumber = EditorTextBox.TextArea.TextView.GetDocumentLineByVisualTop(
            EditorTextBox.TextArea.TextView.ScrollOffset.Y)?.LineNumber ?? 1;
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
            RefreshState(fullRefresh: true);
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
        SetFileCorrupted(_corruptedTabs.Contains(tab));
        SetEditorContent(IsImagePreviewFile(_currentFilePath) ? string.Empty : tab.Content);
        // Restore scroll position synchronously. ScrollToLine works immediately after
        // SetEditorContent because it operates on line numbers rather than pixel offsets,
        // so no layout pass is needed and there is no visible delay on tab switch.
        EditorTextBox.ScrollToLine(tab.TopLineNumber);
        UpdateCurrentDocumentPresentation();

        // Directly set the backing field before NavigateTo so the bail-early
        // check doesn't short-circuit when we're already on the editor page.
        _isHomePageVisible = false;
        NavigateTo(Page.Editor);
        RefreshState(fullRefresh: true);

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
        _corruptedTabs.Remove(tab);

        if (!closingActiveTab)
        {
            RefreshState(fullRefresh: true);
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
        SetFileCorrupted(false);
        EditorTextBox.SyntaxHighlighting = null;
        ConfigureRainbowBrackets(null);
        SetEditorContent(string.Empty);
        RefreshState(fullRefresh: true);
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
            AddRecentFile(path);
            ActivateTab(existingTab);
            return;
        }

        string content;
        bool isCorrupted;
        try
        {
            if (IsImagePreviewFile(path))
            {
                content = string.Empty;
                isCorrupted = false;
                _currentFileEncoding = System.Text.Encoding.UTF8;
            }
            else if (IsBinaryContent(path))
            {
                content = string.Empty;
                isCorrupted = true;
                _currentFileEncoding = System.Text.Encoding.UTF8;
            }
            else
            {
                _currentFileEncoding = DetectFileEncoding(path);
                content = await File.ReadAllTextAsync(path, _currentFileEncoding);
                isCorrupted = false;
            }
        }
        catch (Exception ex)
        {
            await ShowWarningDialogAsync("Open file", ex);
            return;
        }

        // Navigate away from home BEFORE adding the tab, so the CollectionChanged
        // notification evaluates IsEditorTabsVisible with IsHomePageVisible already false.
        // This mirrors the same pattern used in NewFile().
        NavigateTo(Page.Editor);

        var tab = new EditorTab(path, Path.GetFileName(path), content);
        if (isCorrupted)
            _corruptedTabs.Add(tab);
        OpenTabs.Add(tab);
        AddRecentFile(path);
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

        // ApplyThemeBrushes() (called during the constructor) resolves AccentBrush
        // correctly but intentionally skips ApplyThemeToEditor() because the
        // AvaloniaEdit TextEditor isn't fully laid out yet at that point.
        // Now that the window is open and the editor exists, apply the editor theme
        // so that SelectionBrush (and other editor-specific properties) reflect the
        // actual accent colour rather than AvaloniaEdit's built-in defaults.
        ApplyThemeToEditor();

        // Suppress saves only during the automated tab-restore and startup-file
        // open sequence. CollectionChanged / ActiveEditorTab fire SaveSettings()
        // on every tab, but at that point the tab list is still being built, so
        // each intermediate save would overwrite OpenTabPaths with a partial set.
        // The flag is cleared in a finally block so an exception can never leave
        // it permanently set (which would silently disable all future saves).
        try
        {
            _suppressSettingsSave = true;

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
        finally
        {
            // Re-enable saves and do one clean write now that the tab list is complete.
            _suppressSettingsSave = false;
            SaveSettings(immediate: true);
        }

        _ = RefreshExtensionsDataAsync();
        _ = RefreshLatestReleaseAsync();

        if (_isFirstLaunch && !_hasCompletedTutorial)
            await ShowTutorialAsync();
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
        RefreshState(fullRefresh: true);
    }

    private void CloseFolder()
    {
        _currentFolderPath = null;
        FileTreeItems.Clear();
        IsFileExplorerVisible = false;
        RefreshState(fullRefresh: true);
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

            RefreshState(fullRefresh: true);
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
        ReplaceFileTreeItems(items);
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

    private void ReplaceFileTreeItems(IReadOnlyList<FileTreeItem> items)
    {
        _suppressExplorerWidthRefresh = true;
        try
        {
            FileTreeItems.Clear();
            foreach (var item in items)
                FileTreeItems.Add(item);
        }
        finally
        {
            _suppressExplorerWidthRefresh = false;
            OnPropertyChanged(nameof(ExplorerPanelWidth));
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
                .OrderBy(d => Path.GetFileName(d), NaturalSortComparer.OrdinalIgnoreCase)
                .ToArray();

            var files = Directory.GetFiles(dirPath)
                .Where(f => !Path.GetFileName(f).StartsWith('.'))
                .OrderBy(f => Path.GetFileName(f), NaturalSortComparer.OrdinalIgnoreCase)
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

    private void AddRecentFile(string? path) =>
        AddRecentPath(path, isFolder: false);

    private void AddRecentFolder(string? path) =>
        AddRecentPath(path, isFolder: true);

    private void AddRecentPath(string? path, bool isFolder)
    {
        if (string.IsNullOrWhiteSpace(path)) return;

        var existing = RecentFiles.FirstOrDefault(f => string.Equals(f.Path, path, StringComparison.OrdinalIgnoreCase));
        if (existing is not null) RecentFiles.Remove(existing);

        RecentFiles.Insert(0, new RecentFileItem(path, isFolder, DateTime.Now));
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

    // ── Terminal helpers ─────────────────────────────────────────────────────

    private void RefreshAvailableTerminalShells(string? preferredShellId = null)
    {
        AvailableTerminalShells.Clear();

        foreach (var shell in DetectTerminalShells())
            AvailableTerminalShells.Add(shell);

        SelectedTerminalShell = AvailableTerminalShells.FirstOrDefault(shell =>
            string.Equals(shell.Id, preferredShellId, StringComparison.OrdinalIgnoreCase))
            ?? AvailableTerminalShells.FirstOrDefault(shell =>
                string.Equals(shell.Id, "powershell", StringComparison.OrdinalIgnoreCase))
            ?? AvailableTerminalShells.FirstOrDefault();
    }

    private IEnumerable<TerminalShellOption> DetectTerminalShells()
    {
        var shells = new List<TerminalShellOption>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void AddShell(string id, string displayName, string? resolvedPath, string arguments)
        {
            if (string.IsNullOrWhiteSpace(resolvedPath) || !File.Exists(resolvedPath) || !seen.Add(resolvedPath))
                return;

            shells.Add(new TerminalShellOption
            {
                Id = id,
                DisplayName = displayName,
                FileName = resolvedPath,
                Arguments = arguments
            });
        }

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            AddShell(
                "powershell",
                "PowerShell",
                ResolveExecutable("pwsh.exe")
                    ?? ResolveExecutable("powershell.exe", Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.System),
                        @"WindowsPowerShell\v1.0\powershell.exe")),
                "-NoLogo");
            AddShell(
                "windows-powershell",
                "Windows PowerShell",
                ResolveExecutable("powershell.exe", Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.System),
                        @"WindowsPowerShell\v1.0\powershell.exe")),
                "-NoLogo");
            AddShell(
                "cmd",
                "Command Prompt",
                ResolveExecutable(Environment.GetEnvironmentVariable("ComSpec") ?? "cmd.exe")
                    ?? ResolveExecutable("cmd.exe"),
                "/K");
            AddShell(
                "bash",
                "Git Bash",
                ResolveExecutable("bash.exe",
                    @"C:\Program Files\Git\bin\bash.exe",
                    @"C:\Program Files\Git\usr\bin\bash.exe"),
                "--login -i");
        }
        else
        {
            AddShell("bash", "Bash", ResolveExecutable("bash"), "-i");
            AddShell("zsh", "Zsh", ResolveExecutable("zsh"), "-i");
            AddShell("sh", "Shell", ResolveExecutable("sh"), "-i");
        }

        return shells;
    }

    private static string? ResolveExecutable(string fileName, params string[] fallbacks)
    {
        if (Path.IsPathFullyQualified(fileName) && File.Exists(fileName))
            return fileName;

        var pathValue = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        var extensions = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? (Environment.GetEnvironmentVariable("PATHEXT") ?? ".EXE;.CMD;.BAT;.COM")
                .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            : [string.Empty];

        foreach (var path in pathValue.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (string.IsNullOrWhiteSpace(path))
                continue;

            if (Path.HasExtension(fileName))
            {
                var candidate = Path.Combine(path, fileName);
                if (File.Exists(candidate))
                    return candidate;
            }
            else
            {
                foreach (var ext in extensions)
                {
                    var candidate = Path.Combine(path, fileName + ext);
                    if (File.Exists(candidate))
                        return candidate;
                }
            }
        }

        foreach (var fallback in fallbacks)
        {
            if (!string.IsNullOrWhiteSpace(fallback) && File.Exists(fallback))
                return fallback;
        }

        return null;
    }

    private string ResolveTerminalWorkingDirectory()
    {
        if (!string.IsNullOrWhiteSpace(ActiveTerminalSession?.WorkingDirectory) && Directory.Exists(ActiveTerminalSession.WorkingDirectory))
            return ActiveTerminalSession.WorkingDirectory;
        if (!string.IsNullOrWhiteSpace(_currentFolderPath) && Directory.Exists(_currentFolderPath))
            return _currentFolderPath;
        if (!string.IsNullOrWhiteSpace(_currentFilePath))
        {
            var fileDirectory = Path.GetDirectoryName(_currentFilePath);
            if (!string.IsNullOrWhiteSpace(fileDirectory) && Directory.Exists(fileDirectory))
                return fileDirectory;
        }

        return Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
    }

    private TerminalShellOption? GetSelectedTerminalShellOrFallback() =>
        SelectedTerminalShell ?? AvailableTerminalShells.FirstOrDefault();

    private void ToggleTerminalPanel(bool ensureVisible = false)
    {
        NavigateTo(Page.Editor);

        if (ensureVisible)
            IsTerminalVisible = true;
        else
            IsTerminalVisible = !IsTerminalVisible;

        // Do NOT auto-spawn a shell when the panel opens — the user must click
        // "Create terminal" or use Ctrl+Shift+` to start one explicitly.
        if (IsTerminalVisible && ActiveTerminalSession is not null)
            FocusActiveTerminal();
    }

    private void CreateTerminalSession(TerminalShellOption? shell = null, TerminalSession? replaceExisting = null)
    {
        if (!IsTerminalSupported)
            return;

        shell ??= GetSelectedTerminalShellOrFallback();
        if (shell is null)
            return;

        var workingDirectory = ResolveTerminalWorkingDirectory();
        var title = _nextTerminalNumber == 1 ? shell.DisplayName : $"{shell.DisplayName} {_nextTerminalNumber}";
        _nextTerminalNumber++;

        var session = new TerminalSession(shell.Id, shell.DisplayName, title, workingDirectory);
        StartTerminalProcess(session, shell);

        if (replaceExisting is not null)
        {
            var index = TerminalSessions.IndexOf(replaceExisting);
            if (index >= 0)
            {
                CloseTerminalSession(replaceExisting, activateReplacement: false);
                TerminalSessions.Insert(Math.Min(index, TerminalSessions.Count), session);
            }
            else
            {
                TerminalSessions.Add(session);
            }
        }
        else
        {
            TerminalSessions.Add(session);
        }

        ActiveTerminalSession = session;
        IsTerminalVisible = true;
        FocusActiveTerminal();
    }

    private void StartTerminalProcess(TerminalSession session, TerminalShellOption shell)
    {
        // The actual ConPTY Start() is called by the ActiveTerminalSession setter
        // (set just after this method returns in CreateTerminalSession). Here we
        // just wire up the exit event and mark the session as launching so the UI
        // shows the right status text while the process is starting.
        TerminalHostControl.SessionExited += OnSessionExited;

        void OnSessionExited(object? s, EventArgs _)
        {
            // Unhook so stale lambdas from replaced sessions don't pile up.
            TerminalHostControl.SessionExited -= OnSessionExited;
            session.IsRunning = false;
            session.StatusText = "Exited";
            OnPropertyChanged(nameof(TerminalStatusBarText));
            OnPropertyChanged(nameof(ActiveTerminalFooterText));
        }

        session.IsRunning = true;
        session.StatusText = "Launching...";
    }



    private void ClearActiveTerminal()
    {
        if (ActiveTerminalSession is null)
            return;

        SendTextToTerminal(ActiveTerminalSession, GetClearCommandForShell(ActiveTerminalSession.ShellId));
    }

    private static string GetClearCommandForShell(string shellId) =>
        shellId switch
        {
            "bash" or "zsh" or "sh" => "clear\r",
            _ => "cls\r"
        };

    private void RestartActiveTerminal()
    {
        if (ActiveTerminalSession is null)
        {
            CreateTerminalSession();
            return;
        }

        var shell = AvailableTerminalShells.FirstOrDefault(option =>
            string.Equals(option.Id, ActiveTerminalSession.ShellId, StringComparison.OrdinalIgnoreCase))
            ?? GetSelectedTerminalShellOrFallback();
        CreateTerminalSession(shell, ActiveTerminalSession);
    }

    private void RefreshTerminalWindows()
    {
        // PseudoConsoleTerminal is a native Avalonia control — it handles its own
        // layout and rendering. Showing / hiding is driven by IsVisible bindings on
        // the host Grid in AXAML, so there is nothing to manually synchronise here.
        // The method is kept so call-sites that still reference it compile cleanly.
    }

    private void FocusActiveTerminal()
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (IsTerminalVisible && ActiveTerminalSession is not null)
                TerminalHostControl.Focus();
        }, DispatcherPriority.Input);
    }

    private void SendTextToTerminal(TerminalSession session, string text)
    {
        if (string.IsNullOrEmpty(text))
            return;

        TerminalHostControl.SendInput(text);
    }

    private void CloseTerminalSession(TerminalSession session, bool activateReplacement = true)
    {
        // Stop the ConPTY process for this session.
        if (ReferenceEquals(session, ActiveTerminalSession))
            TerminalHostControl.Stop();

        session.IsRunning = false;
        session.Dispose();

        var index = TerminalSessions.IndexOf(session);
        TerminalSessions.Remove(session);

        if (!activateReplacement)
            return;

        if (ReferenceEquals(ActiveTerminalSession, session))
        {
            ActiveTerminalSession = TerminalSessions.Count == 0
                ? null
                : TerminalSessions[Math.Clamp(index - 1, 0, TerminalSessions.Count - 1)];
        }
    }

    private void CloseAllTerminalSessions()
    {
        foreach (var session in TerminalSessions.ToList())
            CloseTerminalSession(session);

        ActiveTerminalSession = null;
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

    // Switches the visible page in one pass - sets all backing fields before firing
    // any notifications, so the UI only re-renders once instead of once per property set.
    private enum Page { Home, Editor, Settings, Extensions, Tutorial, WhatsNew }

    private void NavigateTo(Page page)
    {
        var newHome       = page == Page.Home;
        var newSettings   = page == Page.Settings;
        var newExtensions = page == Page.Extensions;
        var newTutorial   = page == Page.Tutorial;
        var newWhatsNew   = page == Page.WhatsNew;

        // Bail early if nothing actually changed
        if (_isHomePageVisible       == newHome       &&
            _isSettingsPageVisible   == newSettings   &&
            _isExtensionsPageVisible == newExtensions &&
            _isTutorialPageVisible   == newTutorial   &&
            _isWhatsNewPageVisible   == newWhatsNew)
            return;

        _isHomePageVisible       = newHome;
        _isSettingsPageVisible   = newSettings;
        _isExtensionsPageVisible = newExtensions;
        _isTutorialPageVisible   = newTutorial;
        _isWhatsNewPageVisible   = newWhatsNew;

        OnPropertyChanged(nameof(IsHomePageVisible));
        OnPropertyChanged(nameof(IsSettingsPageVisible));
        OnPropertyChanged(nameof(IsExtensionsPageVisible));
        OnPropertyChanged(nameof(IsTutorialPageVisible));
        OnPropertyChanged(nameof(IsWhatsNewPageVisible));
        OnPropertyChanged(nameof(IsEditorPageVisible));
        OnPropertyChanged(nameof(IsEditorTabsVisible));
        OnPropertyChanged(nameof(IsDocumentViewVisible));
        OnPropertyChanged(nameof(IsEmptyStateVisible));
        OnPropertyChanged(nameof(CanShowSaveActions));
        OnPropertyChanged(nameof(FileSummaryText));
        OnPropertyChanged(nameof(FilePathText));
        RefreshState(fullRefresh: true);
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

    private void ToggleTerminalButton_OnClick(object? sender, RoutedEventArgs e) =>
        ToggleTerminalPanel();

    private void NewTerminalButton_OnClick(object? sender, RoutedEventArgs e) =>
        CreateTerminalSession();

    private void ClearTerminalButton_OnClick(object? sender, RoutedEventArgs e) =>
        ClearActiveTerminal();

    private void RestartTerminalButton_OnClick(object? sender, RoutedEventArgs e) =>
        RestartActiveTerminal();

    private void StatusBar_OnPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        e.Handled = true;
    }



    private void CloseTerminalButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (ActiveTerminalSession is not null)
            CloseTerminalSession(ActiveTerminalSession);
    }

    private void OpenTerminalSessionButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: TerminalSession session })
        {
            ActiveTerminalSession = session;
            IsTerminalVisible = true;
            FocusActiveTerminal();
        }
    }

    private void CloseTerminalSessionButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: TerminalSession session })
            CloseTerminalSession(session);
    }

    private async void RenameTerminalSessionMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<TerminalSession>(sender) is not { } session) return;
        var newName = await ShowRenameDialogAsync(session.Title);
        if (newName is null || string.Equals(newName, session.Title, StringComparison.Ordinal)) return;
        session.Title = newName;
    }

    private void DuplicateTerminalSessionMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<TerminalSession>(sender) is not { } session) return;
        var shell = AvailableTerminalShells.FirstOrDefault(s =>
            string.Equals(s.Id, session.ShellId, StringComparison.OrdinalIgnoreCase))
            ?? GetSelectedTerminalShellOrFallback();
        CreateTerminalSession(shell);
    }

    private void RestartTerminalSessionMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<TerminalSession>(sender) is not { } session) return;
        var shell = AvailableTerminalShells.FirstOrDefault(s =>
            string.Equals(s.Id, session.ShellId, StringComparison.OrdinalIgnoreCase))
            ?? GetSelectedTerminalShellOrFallback();
        CreateTerminalSession(shell, session);
    }

    private void CloseOtherTerminalSessionsMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TryGetTaggedData<TerminalSession>(sender) is not { } pivotSession) return;
        var others = TerminalSessions.Where(s => !ReferenceEquals(s, pivotSession)).ToList();
        foreach (var s in others)
            CloseTerminalSession(s);
    }

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
        OpenExtensionsPage(showMarketplaceTab: false, forceRefresh: false);
    }

    private async void RefreshExtensionsButton_OnClick(object? sender, RoutedEventArgs e) =>
        await RefreshExtensionsDataAsync(force: true);

    private void InstalledTabButton_OnClick(object? sender, RoutedEventArgs e) =>
        IsMarketplaceTabSelected = false;

    // Used by the tab strip inside the Extensions page - only switches the tab
    private void MarketplaceTabButton_OnClick(object? sender, RoutedEventArgs e)
    {
        IsMarketplaceTabSelected = true;
        RefreshMarketplaceConnectivityState();
    }

    // Used by the "Visit Marketplace" button on the home screen -
    // opens the Extensions page AND switches to the Marketplace tab
    private void OpenMarketplaceButton_OnClick(object? sender, RoutedEventArgs e)
    {
        OpenExtensionsPage(showMarketplaceTab: true, forceRefresh: true);
    }

    private void OpenTutorialButton_OnClick(object? sender, RoutedEventArgs e)
    {
        _tutorialOpenedFromSettings = true;
        TutorialStepIndex = 0;
        NavigateTo(Page.Tutorial);
    }

    private void OpenWhatsNewButton_OnClick(object? sender, RoutedEventArgs e)
    {
        NavigateTo(Page.WhatsNew);
        IsWhatsNewExpanded = true;
        _ = RefreshLatestReleaseAsync();
    }

    private void BackToEditorButton_OnClick(object? sender, RoutedEventArgs e)
    {
        NavigateTo(Page.Editor);
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
        OpenExtensionsPage(showMarketplaceTab: true, forceRefresh: false);
    }

    private void OpenExtensionsPage(bool showMarketplaceTab, bool forceRefresh)
    {
        NavigateTo(Page.Extensions);
        IsMarketplaceTabSelected = showMarketplaceTab;
        RefreshMarketplaceConnectivityState();
        _ = RefreshExtensionsDataAsync(force: forceRefresh);
    }

    private void OpenReleasesPageButton_OnClick(object? sender, RoutedEventArgs e) =>
        OpenUrl(ReleasesPageUrl);

    private void OpenDiscordButton_OnClick(object? sender, RoutedEventArgs e) =>
        OpenUrl(DiscordServerUrl);

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

    // Convenience handlers used by the tutorial setup step's theme buttons.
    private void ThemeDarkButton_OnClick(object? sender, RoutedEventArgs e)  => ApplyTheme("Dark");
    private void ThemeLightButton_OnClick(object? sender, RoutedEventArgs e) => ApplyTheme("Light");

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
            RefreshState(fullRefresh: true);
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

    private string GetExplorerRootDirectory()
    {
        if (string.IsNullOrWhiteSpace(_currentFolderPath))
            throw new InvalidOperationException("No folder is currently open in the explorer.");

        return _currentFolderPath;
    }

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
        {
            // Snapshot which directories are expanded before wiping the tree,
            // then restore them afterward so the user's expansion state is preserved.
            var expandedPaths = FileTreeItems
                .Where(i => i.IsDirectory && i.IsExpanded)
                .Select(i => i.FullPath)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            await PopulateFileTreeAsync(_currentFolderPath);

            if (expandedPaths.Count > 0)
                await RestoreExpandedPathsAsync(expandedPaths);
        }
    }

    // Re-expands directories that were open before a tree refresh.
    // Works top-down: a parent must be expanded before its children are visible.
    private async Task RestoreExpandedPathsAsync(HashSet<string> expandedPaths)
    {
        // Keep expanding until no more progress can be made (handles nested directories).
        bool anyExpanded;
        do
        {
            anyExpanded = false;
            // Take a snapshot - ToggleDirectoryExpansionAsync mutates FileTreeItems
            var candidates = FileTreeItems
                .Where(i => i.IsDirectory && !i.IsExpanded && expandedPaths.Contains(i.FullPath))
                .ToList();

            foreach (var item in candidates)
            {
                await ToggleDirectoryExpansionAsync(item);
                anyExpanded = true;
            }
        }
        while (anyExpanded);
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
                RefreshState(fullRefresh: true);
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

        var confirmButton = CreateDialogButton("Rename", AccentBrush, AccentBrush, AccentForegroundBrush, () =>
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
            // Collapse: repopulate from scratch - fastest correct collapse
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

    private async void ExplorerHeaderNewFileButton_OnClick(object? sender, RoutedEventArgs e)
    {
        try
        {
            var directory = GetExplorerRootDirectory();
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

    private async void ExplorerHeaderNewFolderButton_OnClick(object? sender, RoutedEventArgs e)
    {
        try
        {
            var directory = GetExplorerRootDirectory();
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

    // ── Encoding detection & change ──────────────────────────────────────────

    /// <summary>
    /// Detects the encoding of <paramref name="path"/> by inspecting its BOM.
    /// Falls back to UTF-8 (no BOM) when no BOM is present.
    /// </summary>
    private static System.Text.Encoding DetectFileEncoding(string path)
    {
        try
        {
            Span<byte> bom = stackalloc byte[4];
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var read = fs.Read(bom);

            if (read >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)
                return new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: true);   // UTF-8 BOM
            if (read >= 2 && bom[0] == 0xFF && bom[1] == 0xFE)
                return System.Text.Encoding.Unicode;          // UTF-16 LE
            if (read >= 2 && bom[0] == 0xFE && bom[1] == 0xFF)
                return System.Text.Encoding.BigEndianUnicode; // UTF-16 BE
            if (read >= 4 && bom[0] == 0x00 && bom[1] == 0x00 && bom[2] == 0xFE && bom[3] == 0xFF)
                return System.Text.Encoding.UTF32;

            // No BOM - default to UTF-8 without BOM
            return new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        }
        catch
        {
            return System.Text.Encoding.UTF8;
        }
    }

    /// <summary>
    /// Shows a small encoding-picker dialog and, if a different encoding is chosen,
    /// immediately re-saves the file with that encoding.
    /// </summary>
    private async void EncodingStatusBarButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (!HasFileOpen) return;

        // CodePagesEncodingProvider is required for non-Unicode encodings (e.g. 1252) on
        // .NET Core / .NET 5+. Registering it more than once is safe - it's a no-op.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Build the list defensively: skip any encoding the current runtime can't supply.
        var candidateEncodings = new (string Label, Func<System.Text.Encoding> Factory)[]
        {
            ("UTF-8",          () => new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false)),
            ("UTF-8 with BOM", () => new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: true)),
            ("UTF-16 LE",      () => System.Text.Encoding.Unicode),
            ("UTF-16 BE",      () => System.Text.Encoding.BigEndianUnicode),
            ("UTF-32",         () => System.Text.Encoding.UTF32),
            ("Windows-1252",   () => System.Text.Encoding.GetEncoding(1252)),
            ("ASCII",          () => System.Text.Encoding.ASCII),
        };

        var encodings = candidateEncodings
            .Select(c =>
            {
                try   { return ((string Label, System.Text.Encoding Enc)?)(c.Label, c.Factory()); }
                catch { return null; }
            })
            .Where(x => x is not null)
            .Select(x => x!.Value)
            .ToArray();

        System.Text.Encoding? chosen = null;
        Window? dialog = null;

        var panel = new StackPanel { Spacing = 6, Margin = new Thickness(16) };

        panel.Children.Add(new TextBlock
        {
            Text = "Save file with encoding:",
            FontSize = 13,
            FontWeight = FontWeight.SemiBold,
            Foreground = Brushes.White,
            Margin = new Thickness(0, 0, 0, 8),
        });

        foreach (var (label, enc) in encodings)
        {
            var isCurrent = enc.CodePage == _currentFileEncoding.CodePage &&
                            enc.GetPreamble().Length == _currentFileEncoding.GetPreamble().Length;
            var btn = new Button
            {
                Content = isCurrent ? $"{label}  ✓" : label,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                HorizontalContentAlignment = HorizontalAlignment.Left,
                Background = isCurrent
                    ? new SolidColorBrush(Color.Parse("#2D1F4A"))
                    : new SolidColorBrush(Color.Parse("#252526")),
                Foreground = isCurrent
                    ? new SolidColorBrush(Color.Parse("#C084FC"))
                    : Brushes.White,
                BorderBrush = new SolidColorBrush(Color.Parse("#3A3A3A")),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(12, 7),
            };
            var capturedEnc = enc;
            btn.Click += (_, _) =>
            {
                chosen = capturedEnc;
                dialog?.Close();
            };
            panel.Children.Add(btn);
        }

        panel.Children.Add(new TextBlock
        {
            Text = "The file will be re-saved immediately with the chosen encoding.",
            FontSize = 11,
            Foreground = new SolidColorBrush(Color.Parse("#606060")),
            Margin = new Thickness(0, 8, 0, 0),
            TextWrapping = TextWrapping.Wrap,
        });

        dialog = new Window
        {
            Title = "Change File Encoding",
            Width = 280,
            SizeToContent = SizeToContent.Height,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background = new SolidColorBrush(Color.Parse("#1E1E1E")),
            Content = panel,
        };

        await dialog.ShowDialog(this);

        if (chosen is null) return;

        // Preamble length distinguishes UTF-8 vs UTF-8 BOM even when CodePage matches.
        if (chosen.CodePage == _currentFileEncoding.CodePage &&
            chosen.GetPreamble().Length == _currentFileEncoding.GetPreamble().Length)
            return;

        try
        {
            await File.WriteAllTextAsync(_currentFilePath!, EditorTextBox.Document.Text, chosen);
            _currentFileEncoding = chosen;
            OnPropertyChanged(nameof(EncodingDisplayText));
        }
        catch (Exception ex)
        {
            await ShowWarningDialogAsync("Change encoding", ex);
        }
    }

    private void FindNextButton_OnClick(object? sender, RoutedEventArgs e) =>
        FindInEditor(forward: true);

    private void FindPrevButton_OnClick(object? sender, RoutedEventArgs e) =>
        FindInEditor(forward: false);

    private void CloseFindPanel_OnClick(object? sender, RoutedEventArgs e) =>
        IsFindPanelVisible = false;

    // ── Editor context menu (right-click) ─────────────────────────────────────

    private void EditorContextMenu_Opening(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        var hasSelection = EditorTextBox?.TextArea?.Selection is { IsEmpty: false };
        if (sender is not ContextMenu menu) return;
        foreach (var item in menu.Items.OfType<MenuItem>())
        {
            if (item.Name is "EditorCutMenuItem" or "EditorCopyMenuItem")
                item.IsEnabled = hasSelection;
        }
    }

    private void EditorCutMenuItem_OnClick(object? sender, RoutedEventArgs e) =>
        EditorTextBox?.TextArea?.Selection?.ReplaceSelectionWithText(string.Empty);

    private async void EditorCopyMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (EditorTextBox?.TextArea?.Selection is not { } sel) return;
        var text = sel.GetText();
        if (string.IsNullOrEmpty(text)) return;
        var clipboard = TopLevel.GetTopLevel(this)?.Clipboard;
        if (clipboard is not null)
            await clipboard.SetTextAsync(text);
    }

    private async void EditorPasteMenuItem_OnClick(object? sender, RoutedEventArgs e)
    {
        if (EditorTextBox?.TextArea is not { } textArea) return;
        var clipboard = TopLevel.GetTopLevel(this)?.Clipboard;
        if (clipboard is null) return;
        var text = await clipboard.TryGetTextAsync();
        if (!string.IsNullOrEmpty(text))
            textArea.Selection.ReplaceSelectionWithText(text);
    }

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

    private async void UpdateAllMarketplaceExtensionsButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (IsUpdatingAllExtensions)
            return;

        var pendingUpdateIds = MarketplaceExtensions
            .Where(extension => extension.IsUpdateAvailable && extension.IsInstallEnabled)
            .Select(extension => extension.Id)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (pendingUpdateIds.Count == 0)
        {
            ExtensionsStatusText = "All marketplace extensions are already up to date.";
            return;
        }

        IsUpdatingAllExtensions = true;
        var successfulUpdates = 0;
        var failedUpdates = 0;

        try
        {
            for (var index = 0; index < pendingUpdateIds.Count; index++)
            {
                var extensionId = pendingUpdateIds[index];
                var marketplaceExtension = MarketplaceExtensions.FirstOrDefault(entry =>
                    entry.Id.Equals(extensionId, StringComparison.OrdinalIgnoreCase));

                if (marketplaceExtension is null || !marketplaceExtension.IsUpdateAvailable)
                    continue;

                ExtensionsStatusText = $"Updating {index + 1} of {pendingUpdateIds.Count}: {marketplaceExtension.Name}...";
                await InstallMarketplaceExtensionAsync(marketplaceExtension);

                var refreshedExtension = MarketplaceExtensions.FirstOrDefault(entry =>
                    entry.Id.Equals(extensionId, StringComparison.OrdinalIgnoreCase));

                if (refreshedExtension is not null && refreshedExtension.IsInstalled && !refreshedExtension.IsUpdateAvailable)
                    successfulUpdates++;
                else
                    failedUpdates++;
            }

            ExtensionsStatusText = failedUpdates == 0
                ? $"Updated {successfulUpdates} extension{(successfulUpdates == 1 ? string.Empty : "s")}."
                : $"Updated {successfulUpdates} extension{(successfulUpdates == 1 ? string.Empty : "s")}. {failedUpdates} still need attention.";
        }
        finally
        {
            IsUpdatingAllExtensions = false;
        }
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

    // TextEditor fires EventHandler (not RoutedEventHandler) - signature must match exactly
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
        QueueRefreshState(fullRefresh: true);
        QueueWordCountRefresh();
        RestartAutoSaveTimerIfNeeded();
    }

	// Fires BEFORE the character is written into the document.
    // Used to skip-over an already-present auto-inserted closing character
    // instead of inserting a duplicate.
    private void EditorTextArea_OnTextEntering(object? sender, TextInputEventArgs e)
    {
        if (!IsSmartSyntaxEnabled()) return;
        if (string.IsNullOrEmpty(e.Text)) return;
        var ch     = e.Text[0];
        var caret  = EditorTextBox.TextArea.Caret;
        var doc    = EditorTextBox.Document;
        var offset = caret.Offset;
        var selection = EditorTextBox.TextArea.Selection;

        if (!selection.IsEmpty && BracketPairs.TryGetValue(ch, out var selectionClosing))
        {
            var segment = selection.SurroundingSegment;
            if (segment is not null)
            {
                var selectedText = selection.GetText();
                doc.Replace(segment, $"{ch}{selectedText}{selectionClosing}");
                caret.Offset = segment.Offset + selectedText.Length + 2;
                e.Handled = true;
                return;
            }
        }

        if (!ClosingChars.Contains(ch)) return;

        if (ch == '}' && TryAlignClosingDelimiterBeforeInsert(doc, caret, '}'))
            offset = caret.Offset;

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
        if (!IsSmartSyntaxEnabled()) return;
        if (string.IsNullOrEmpty(e.Text)) return;
        var ch = e.Text[0];

        if (!BracketPairs.TryGetValue(ch, out var closing)) return;

        var caret  = EditorTextBox.TextArea.Caret;
        var doc    = EditorTextBox.Document;
        var offset = caret.Offset;

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
            case Key.Enter when IsSmartSyntaxEnabled() && (e.KeyModifiers & KeyModifiers.Shift) != KeyModifiers.Shift:
                HandleSmartEnter(doc, caret);
                e.Handled = true;
                return;

            case Key.Tab when IsSmartSyntaxEnabled() && e.KeyModifiers == KeyModifiers.Shift:
                HandleTabKey(() => HandleOutdent(doc, textArea.Selection, caret), doc, caret);
                e.Handled = true;
                return;

            case Key.Tab when IsSmartSyntaxEnabled() && e.KeyModifiers == KeyModifiers.None:
                HandleTabKey(() => HandleIndent(doc, textArea.Selection, caret), doc, caret);
                e.Handled = true;
                return;

            case Key.Back when IsSmartSyntaxEnabled():
                if (HandleSmartBackspace(doc, caret))
                {
                    e.Handled = true;
                    return;
                }
                break;

            case Key.V when IsSmartSyntaxEnabled() && e.KeyModifiers == KeyModifiers.Control:
                e.Handled = true;
                _ = HandleSmartPasteAsync(doc, textArea, caret);
                return;

            case Key.Oem2 when IsSmartSyntaxEnabled() && (e.KeyModifiers & KeyModifiers.Control) == KeyModifiers.Control:
                ToggleLineComment(doc, textArea, textArea.Selection, caret);
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

    private bool HandleSmartBackspace(AvaloniaEdit.Document.TextDocument doc, AvaloniaEdit.Editing.Caret caret)
    {
        var selection = EditorTextBox.TextArea.Selection;
        if (selection is not null && !selection.IsEmpty)
            return false;

        var offset = caret.Offset;
        if (offset <= 0 || offset >= doc.TextLength)
            return false;

        var opening = doc.GetCharAt(offset - 1);
        if (!BracketPairs.TryGetValue(opening, out var closing))
            return false;

        if (doc.GetCharAt(offset) != closing)
            return false;

        doc.Remove(offset - 1, 2);
        SetCaretOffsetSafely(caret, doc, offset - 1);
        return true;
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

    private bool TryAlignClosingDelimiterBeforeInsert(AvaloniaEdit.Document.TextDocument doc, AvaloniaEdit.Editing.Caret caret, char closing)
    {
        var offset = caret.Offset;
        var line = doc.GetLineByOffset(offset);
        var lineText = doc.GetText(line);
        var caretColumnInLine = offset - line.Offset;
        var textBeforeCaret = lineText[..Math.Min(caretColumnInLine, lineText.Length)];

        if (textBeforeCaret.Length == 0 || !string.IsNullOrWhiteSpace(textBeforeCaret))
            return false;

        var removable = GetOutdentLength(textBeforeCaret, textBeforeCaret.Length);
        if (removable <= 0)
            return false;

        doc.Remove(line.Offset, removable);
        SetCaretOffsetSafely(caret, doc, offset - removable);
        return true;
    }

    private static string NormalizeLineEndings(string text) =>
        text.Replace("\r\n", "\n").Replace('\r', '\n');

    private string ReindentPastedText(string text, AvaloniaEdit.Document.TextDocument doc, int offset)
    {
        var normalized = NormalizeLineEndings(text);
        if (!normalized.Contains('\n'))
            return text;

        var line = doc.GetLineByOffset(Math.Clamp(offset, 0, doc.TextLength));
        var lineText = doc.GetText(line);
        var caretColumnInLine = Math.Clamp(offset - line.Offset, 0, lineText.Length);
        var textBeforeCaret = lineText[..caretColumnInLine];
        var baseIndent = GetLeadingWhitespace(textBeforeCaret);
        var pasteLines = normalized.Split('\n');

        if (pasteLines.Length <= 1)
            return text;

        var firstNonEmptyIndex = Array.FindIndex(pasteLines, static l => !string.IsNullOrWhiteSpace(l));
        if (firstNonEmptyIndex < 0)
            return text;

        var commonIndent = GetLeadingWhitespace(pasteLines[firstNonEmptyIndex]);
        for (var i = firstNonEmptyIndex + 1; i < pasteLines.Length; i++)
        {
            if (string.IsNullOrWhiteSpace(pasteLines[i]))
                continue;

            commonIndent = GetSharedIndent(commonIndent, GetLeadingWhitespace(pasteLines[i]));
            if (commonIndent.Length == 0)
                break;
        }

        for (var i = 1; i < pasteLines.Length; i++)
        {
            if (pasteLines[i].Length == 0)
                continue;

            var trimmedLine = pasteLines[i];
            if (commonIndent.Length > 0 && trimmedLine.StartsWith(commonIndent, StringComparison.Ordinal))
                trimmedLine = trimmedLine[commonIndent.Length..];

            pasteLines[i] = baseIndent + trimmedLine;
        }

        return string.Join(Environment.NewLine, pasteLines);
    }

    private static string GetSharedIndent(string left, string right)
    {
        var max = Math.Min(left.Length, right.Length);
        var length = 0;
        while (length < max && left[length] == right[length])
            length++;

        return left[..length];
    }

    private async Task HandleSmartPasteAsync(AvaloniaEdit.Document.TextDocument doc, AvaloniaEdit.Editing.TextArea textArea, AvaloniaEdit.Editing.Caret caret)
    {
        var clipboard = TopLevel.GetTopLevel(this)?.Clipboard;
        if (clipboard is null)
            return;

        var text = await clipboard.TryGetTextAsync();
        if (string.IsNullOrEmpty(text))
            return;

        var insertionText = ReindentPastedText(text, doc, caret.Offset);
        var selection = textArea.Selection;
        if (selection is not null && !selection.IsEmpty && selection.SurroundingSegment is not null)
        {
            var segment = selection.SurroundingSegment;
            doc.Replace(segment, insertionText);
            SetCaretOffsetSafely(caret, doc, segment.Offset + insertionText.Length);
            return;
        }

        var safeOffset = Math.Clamp(caret.Offset, 0, doc.TextLength);
        doc.Insert(safeOffset, insertionText);
        SetCaretOffsetSafely(caret, doc, safeOffset + insertionText.Length);
    }

    private void ToggleLineComment(AvaloniaEdit.Document.TextDocument doc, AvaloniaEdit.Editing.TextArea textArea, AvaloniaEdit.Editing.Selection? selection, AvaloniaEdit.Editing.Caret caret)
    {
        var lineCommentToken = CurrentLanguageExtension?.CommentLine;
        if (string.IsNullOrWhiteSpace(lineCommentToken))
            return;

        var startOffset = selection is not null && !selection.IsEmpty && selection.SurroundingSegment is not null
            ? selection.SurroundingSegment.Offset
            : caret.Offset;
        var endOffset = selection is not null && !selection.IsEmpty && selection.SurroundingSegment is not null
            ? selection.SurroundingSegment.EndOffset
            : caret.Offset;

        var lines = GetSelectedLines(doc, startOffset, endOffset);
        if (lines.Count == 0)
            return;

        var shouldUncomment = lines
            .Where(line => !string.IsNullOrWhiteSpace(doc.GetText(line)))
            .All(line =>
            {
                var text = doc.GetText(line);
                var indent = GetLeadingWhitespace(text);
                return text[indent.Length..].StartsWith(lineCommentToken, StringComparison.Ordinal);
            });

        var delta = 0;
        foreach (var line in lines.OrderByDescending(l => l.Offset))
        {
            var text = doc.GetText(line);
            if (string.IsNullOrWhiteSpace(text))
                continue;

            var indent = GetLeadingWhitespace(text);
            var commentOffset = line.Offset + indent.Length;
            if (shouldUncomment)
            {
                if (text[indent.Length..].StartsWith(lineCommentToken, StringComparison.Ordinal))
                {
                    var removedForLine = lineCommentToken.Length;
                    doc.Remove(commentOffset, lineCommentToken.Length);
                    if (text.Length > indent.Length + lineCommentToken.Length && text[indent.Length + lineCommentToken.Length] == ' ')
                    {
                        doc.Remove(commentOffset, 1);
                        removedForLine++;
                    }

                    delta -= removedForLine;
                }
            }
            else
            {
                doc.Insert(commentOffset, lineCommentToken + " ");
                delta += lineCommentToken.Length + 1;
            }
        }

        if (selection is not null && !selection.IsEmpty && selection.SurroundingSegment is not null)
        {
            var segment = selection.SurroundingSegment;
            var newEnd = Math.Max(segment.Offset, segment.EndOffset + delta);
            textArea.Selection = AvaloniaEdit.Editing.Selection.Create(textArea, segment.Offset, newEnd);
        }
        else
        {
            SetCaretOffsetSafely(caret, doc, caret.Offset + delta);
        }
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
                    return; // User aborted - leave the window open.

                case UnsavedTabAction.Save:
                    ActivateTab(tab, focusEditor: false);
                    if (!await SaveAsync(allowPromptForPath: true, forcePromptForPath: false))
                        return; // Save was cancelled - leave the window open.
                    break;

                // UnsavedTabAction.Discard - just continue to the next tab.
            }
        }

        // All dirty tabs resolved - close for real.
        _isConfirmedClose = true;
        Close();
    }

    private void MainWindow_OnClosed(object? sender, EventArgs e)
    {
        SaveSettings(immediate: true);
        _autoSaveTimer.Stop();
        _autoSaveStatusTimer.Stop();
        _discordReconnectTimer.Stop();
        _extensionsRefreshDebounceTimer.Stop();
        _wordCountRefreshTimer.Stop();
        _settingsSaveDebounceTimer.Stop();
        _windowsAccentPollTimer.Stop();
        NetworkChange.NetworkAvailabilityChanged -= NetworkChange_OnNetworkAvailabilityChanged;
        NetworkChange.NetworkAddressChanged -= NetworkChange_OnNetworkAddressChanged;
        CloseAllTerminalSessions();
        DisposeExtensionFolderWatchers();
        DisposeDiscordPresence();
        CurrentImagePreview = null;
    }

    // Runs on the UI thread every 2 s; re-applies the accent only when the
    // registry value has actually changed, so there's no unnecessary work.
    private void WindowsAccentPollTimer_OnTick(object? sender, EventArgs e)
    {
        var current = GetWindowsAccentColor() ?? string.Empty;
        if (current == _lastSeenWindowsAccentHex) return;
        _lastSeenWindowsAccentHex = current;
        ApplyAccentOverride();
        if (_accentColorMode == "windows")
        {
            ApplyThemeToEditor();
            RefreshExtensionTheme();
        }
    }

    private void NetworkChange_OnNetworkAvailabilityChanged(object? sender, NetworkAvailabilityEventArgs e) =>
        Dispatcher.UIThread.Post(() => RefreshMarketplaceConnectivityState());

    private void NetworkChange_OnNetworkAddressChanged(object? sender, EventArgs e) =>
        Dispatcher.UIThread.Post(() => RefreshMarketplaceConnectivityState());

    private static bool HasActiveWirelessConnection() =>
        NetworkInterface.GetAllNetworkInterfaces().Any(networkInterface =>
            networkInterface.NetworkInterfaceType == NetworkInterfaceType.Wireless80211 &&
            networkInterface.OperationalStatus == OperationalStatus.Up &&
            networkInterface.GetIPProperties().UnicastAddresses.Any(address => !System.Net.IPAddress.IsLoopback(address.Address)));

    private static bool HasActiveInternetConnection() =>
        NetworkInterface.GetIsNetworkAvailable() &&
        NetworkInterface.GetAllNetworkInterfaces().Any(networkInterface =>
            networkInterface.OperationalStatus == OperationalStatus.Up &&
            networkInterface.NetworkInterfaceType is not NetworkInterfaceType.Loopback &&
            networkInterface.NetworkInterfaceType is not NetworkInterfaceType.Tunnel);

    private static bool IsConnectivityFailure(Exception exception) =>
        exception is HttpRequestException or TaskCanceledException;

    private void RefreshMarketplaceConnectivityState(string? operation = null, Exception? exception = null)
    {
        var hasWirelessConnection = HasActiveWirelessConnection();
        var hasInternetConnection = HasActiveInternetConnection();

        string message = string.Empty;

        if (!hasInternetConnection)
        {
            message = hasWirelessConnection
                ? "Kodo cannot reach the internet right now. Marketplace installs and updates may fail until your connection comes back."
                : "No Wi-Fi or internet connection detected. Marketplace installs and downloads may fail until you are back online.";
        }
        else if (!hasWirelessConnection && exception is not null && IsConnectivityFailure(exception))
        {
            message = "No Wi-Fi connection detected. If you are expecting wireless access, reconnect first. Marketplace downloads can fail while the app is offline or the network is unstable.";
        }

        if (!string.IsNullOrWhiteSpace(operation) && !string.IsNullOrWhiteSpace(message))
            message = $"{message} Latest issue: {operation}.";

        MarketplaceConnectivityMessage = message;
        IsMarketplaceConnectivityWarningVisible = !string.IsNullOrWhiteSpace(message);
    }

    private TutorialStep CurrentTutorialStep => TutorialSteps[TutorialStepIndex];

    private void OnTutorialStepChanged()
    {
        OnPropertyChanged(nameof(TutorialStepIndex));
        OnPropertyChanged(nameof(TutorialStepLabel));
        OnPropertyChanged(nameof(TutorialProgressDotsText));
        OnPropertyChanged(nameof(TutorialSectionTitle));
        OnPropertyChanged(nameof(TutorialTitle));
        OnPropertyChanged(nameof(TutorialBody));
        OnPropertyChanged(nameof(TutorialShortcutText));
        OnPropertyChanged(nameof(TutorialSpotlightTitle));
        OnPropertyChanged(nameof(TutorialHighlightOne));
        OnPropertyChanged(nameof(TutorialHighlightTwo));
        OnPropertyChanged(nameof(TutorialHighlightThree));
        OnPropertyChanged(nameof(CanGoToPreviousTutorialStep));
        OnPropertyChanged(nameof(TutorialPrimaryButtonText));
        OnPropertyChanged(nameof(IsTutorialSetupStep));
        OnPropertyChanged(nameof(IsNotTutorialSetupStep));
        OnPropertyChanged(nameof(IsTutorialWelcomeStep));
        OnPropertyChanged(nameof(IsNotTutorialWelcomeStep));
        OnPropertyChanged(nameof(IsTutorialHeaderVisible));
    }

    private void CompleteTutorialAndReturnHome()
    {
        _hasCompletedTutorial = true;
        SaveSettings();
        NavigateTo(Page.Home);
    }

    private void PreviousTutorialStepButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TutorialStepIndex > 0)
            TutorialStepIndex--;
    }

    private void NextTutorialStepButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (TutorialStepIndex < TutorialSteps.Length - 1)
        {
            TutorialStepIndex++;
            return;
        }

        CompleteTutorialAndReturnHome();
    }

    private void SkipTutorialButton_OnClick(object? sender, RoutedEventArgs e) =>
        CompleteTutorialAndReturnHome();

    private async void MainWindow_OnKeyDown(object? sender, KeyEventArgs e)
    {
        var hasControl = (e.KeyModifiers & KeyModifiers.Control) == KeyModifiers.Control;
        var hasShift   = (e.KeyModifiers & KeyModifiers.Shift)   == KeyModifiers.Shift;

        // Escape - dismiss Settings / Extensions / Tutorial / WhatsNew and return to editor
        if (e.Key == Key.Escape && !hasControl)
        {
            if (IsTutorialPageVisible)
            {
                CompleteTutorialAndReturnHome();
                e.Handled = true;
            }
            else if (IsSettingsPageVisible || IsExtensionsPageVisible || IsWhatsNewPageVisible)
            {
                NavigateTo(Page.Editor);
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
                    // Ctrl+Shift+E - back to editor (original behaviour)
                    NavigateTo(Page.Editor);
                    FocusEditor();
                    e.Handled = true;
                }
                else
                {
                    // Ctrl+E - open Extensions
                    NavigateTo(Page.Extensions);
                    RefreshMarketplaceConnectivityState();
                    _ = RefreshExtensionsDataAsync();
                    e.Handled = true;
                }
                break;

            case Key.OemComma:
                // Ctrl+, - open Settings
                NavigateTo(Page.Settings);
                e.Handled = true;
                break;

            case Key.H:
                // Ctrl+H - go to Home
                NavigateTo(Page.Home);
                e.Handled = true;
                break;

            case Key.S:
                if (hasShift)
                {
                    // Ctrl+Shift+S - Save As (always prompt for path)
                    e.Handled = true;
                    await SaveAsAsync();
                }
                else
                {
                    // Ctrl+S - Save
                    e.Handled = true;
                    await SaveAsync();
                }
                break;

            case Key.O:
                e.Handled = true;
                await OpenFileAsync();
                break;

            case Key.K:
                // Ctrl+K - toggle folder open/close
                e.Handled = true;
                if (IsFolderOpen)
                    CloseFolder();
                else
                    await OpenFolderAsync();
                break;

            case Key.B:
                // Ctrl+B - toggle file explorer sidebar
                IsFileExplorerVisible = !IsFileExplorerVisible;
                e.Handled = true;
                break;

            case Key.J:
                // Ctrl+J - toggle the bottom terminal panel
                ToggleTerminalPanel();
                e.Handled = true;
                break;

            case Key.W:
                // Ctrl+W - close current tab
                if (ActiveEditorTab is not null)
                    await RequestCloseTabAsync(ActiveEditorTab);
                e.Handled = true;
                break;

            case Key.F:
                // Ctrl+F - toggle find panel
                if (CanShowFindInFile)
                    IsFindPanelVisible = !IsFindPanelVisible;
                else
                    IsFindPanelVisible = false;
                e.Handled = true;
                break;

            case Key.X when hasShift:
                // Ctrl+Shift+X - open Extensions (secondary binding)
                NavigateTo(Page.Extensions);
                RefreshMarketplaceConnectivityState();
                _ = RefreshExtensionsDataAsync();
                e.Handled = true;
                break;

            case Key.Oem3:
                if (hasShift)
                    CreateTerminalSession();
                else
                    ToggleTerminalPanel();
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
                CreateDialogButton("Save", AccentBrush, AccentBrush, AccentForegroundBrush, saveAction)
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


    // ── First-launch tutorial ────────────────────────────────────────────────
    //
    // Shown once when settings.json doesn't exist (genuine first run). The flag
    // HasCompletedTutorial is written to settings afterward so subsequent launches
    // skip it even if the user never explicitly dismissed the window.

    private Task ShowTutorialAsync()
    {
        try
        {
            _tutorialOpenedFromSettings = false;
            TutorialStepIndex = 0;
            NavigateTo(Page.Tutorial);
        }
        catch
        {
            // Tutorial failure must never crash the app.
        }

        return Task.CompletedTask;
    }

    // Shows a non-fatal warning dialog that mirrors the crash dialog in App.axaml.cs
    // but uses softer wording. Call this from any recoverable error path where the
    // user needs to know something went wrong.
    private async Task ShowWarningDialogAsync(string context, Exception exception)
    {
        var source = "MainWindow.Warning";
        KodoDiagnostics.WriteDiagnosticLog(source, exception, false, "Warning", context);

        if (ShouldSuppressWarningDialog(context, exception))
        {
            KodoDiagnostics.WriteDebugFallback($"Suppressed duplicate warning dialog for '{context}'.", exception);
            return;
        }

        try
        {
            // Determine whether this context is file-critical (data may be at risk).
            var isFileOperation = context.StartsWith("File save", StringComparison.OrdinalIgnoreCase)
                               || context.StartsWith("Auto-save", StringComparison.OrdinalIgnoreCase);
 
            var subtitleMessage = isFileOperation
                ? "Kodo could not complete this file operation. Your in-editor content is still intact — try saving again or use Save As to choose a different location."
                : "Kodo ran into a problem with this operation. No data was lost — you can try again.";
 
            // --- Header ---
            var titleText = new TextBlock
            {
                Text         = "Something went wrong",
                FontSize     = 16,
                FontWeight   = FontWeight.SemiBold,
                Foreground   = PrimaryTextBrush,
                TextWrapping = TextWrapping.Wrap,
            };
 
            var subtitleText = new TextBlock
            {
                Text         = subtitleMessage,
                FontSize     = 13,
                Foreground   = MutedTextBrush,
                TextWrapping = TextWrapping.Wrap,
                Margin       = new Thickness(0, 4, 0, 0),
            };
 
            // Context badge (e.g. "File save", "Extension install — MyLang")
            var contextBadge = new Border
            {
                Background      = ButtonBrush,
                BorderBrush     = SurfaceBorderBrush,
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(6),
                Padding         = new Thickness(10, 5),
                HorizontalAlignment = HorizontalAlignment.Left,
                Child = new TextBlock
                {
                    Text       = context,
                    FontSize   = 12,
                    FontFamily = new FontFamily("Cascadia Code,Consolas,Menlo,monospace"),
                    Foreground = new SolidColorBrush(Color.Parse("#9CDCFE")),
                },
            };

            var metadataText = new SelectableTextBlock
            {
                Text         = KodoDiagnostics.BuildDiagnosticSummary(source, false, context),
                FontSize     = 11,
                FontFamily   = new FontFamily("Cascadia Code,Consolas,Menlo,monospace"),
                Foreground   = MutedTextBrush,
                TextWrapping = TextWrapping.Wrap,
            };
 
            // Human-readable error message — shown above the raw stack trace so the
            // user gets an immediate plain-English explanation before seeing the detail.
            var errorMessageText = new TextBlock
            {
                Text         = string.IsNullOrWhiteSpace(exception.Message)
                                   ? "An unexpected error occurred."
                                   : exception.Message,
                FontSize     = 13,
                Foreground   = PrimaryTextBrush,
                TextWrapping = TextWrapping.Wrap,
            };
 
            // Collapsible stack trace — SelectableTextBlock so users can copy it.
            var exceptionText = new SelectableTextBlock
            {
                Text         = KodoDiagnostics.BuildDiagnosticPayload(source, exception, false, "Warning", context),
                FontSize     = 12,
                FontFamily   = new FontFamily("Cascadia Code,Consolas,Menlo,monospace"),
                Foreground   = new SolidColorBrush(Color.Parse("#CE9178")),
                TextWrapping = TextWrapping.Wrap,
            };
 
            var exceptionScroll = new ScrollViewer
            {
                Content  = exceptionText,
                MaxHeight = 200,
                VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            };
 
            // Uses CardBrush so the background honours the active theme (Light/Dark/extension).
            var exceptionBorder = new Border
            {
                Background      = CardBrush,
                BorderBrush     = SurfaceBorderBrush,
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(8),
                Padding         = new Thickness(12),
                Child           = exceptionScroll,
            };

            var logPathText = new TextBlock
            {
                Text         = $"Full log written to: {KodoDiagnostics.LogFilePath}",
                FontSize     = 11,
                Foreground   = MutedTextBrush,
                TextWrapping = TextWrapping.Wrap,
            };
 
            // --- Action buttons ---
            var copyButton = new Button
            {
                Content             = "Copy to Clipboard",
                HorizontalAlignment = HorizontalAlignment.Left,
                Padding             = new Thickness(16, 8),
                Background          = ButtonBrush,
                Foreground          = MutedTextBrush,
                BorderBrush         = SurfaceBorderBrush,
                BorderThickness     = new Thickness(1),
                CornerRadius        = new CornerRadius(8),
            };
 
            var dismissButton = new Button
            {
                Content             = "Dismiss",
                HorizontalAlignment = HorizontalAlignment.Right,
                Padding             = new Thickness(20, 8),
                Background          = AccentBrush,
                Foreground          = AccentForegroundBrush,
                BorderThickness     = new Thickness(0),
                CornerRadius        = new CornerRadius(8),
            };
 
            var buttonRow = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
            buttonRow.Children.Add(copyButton);
            Grid.SetColumn(dismissButton, 1);
            buttonRow.Children.Add(dismissButton);
 
            var content = new StackPanel
            {
                Spacing  = 12,
                Margin   = new Thickness(20),
                Children =
                {
                    titleText,
                    subtitleText,
                    contextBadge,
                    metadataText,
                    errorMessageText,
                    exceptionBorder,
                    logPathText,
                    buttonRow,
                },
            };
 
            Window? dialog = null;
            dialog = new Window
            {
                // Em-dash consistent with the crash dialog title.
                Title  = "Kodo — Error",
                Width  = 520,
                SizeToContent = SizeToContent.Height,
                MinWidth  = 380,
                MinHeight = 180,
                MaxHeight = 660,
                CanResize = true,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Background = CardBrush,
                Content    = content,
            };
 
            copyButton.Click += async (_, _) =>
            {
                try
                {
                    var clip = TopLevel.GetTopLevel(dialog)?.Clipboard;
                    if (clip is not null)
                    {
                        var text = KodoDiagnostics.BuildDiagnosticPayload(source, exception, false, "Warning", context);
                        await clip.SetTextAsync(text);
                        copyButton.Content   = "Copied!";
                        copyButton.Foreground = PrimaryTextBrush;
                    }
                }
                catch
                {
                    // Clipboard failures must not crash the error dialog.
                }
            };
 
            dismissButton.Click += (_, _) => dialog!.Close();
            await dialog.ShowDialog(this);
        }
        catch (Exception dialogEx)
        {
            KodoDiagnostics.WriteDiagnosticLog(source, dialogEx, false, "Warning Dialog Failure", context);
            KodoDiagnostics.WriteDebugFallback($"ShowWarningDialogAsync failed to display for context '{context}'.", dialogEx);
        }
    }

    private bool ShouldSuppressWarningDialog(string context, Exception exception)
    {
        var key = $"{context}|{exception.GetType().FullName}|{exception.Message}";
        var now = DateTime.UtcNow;
        if (_warningDialogCooldowns.TryGetValue(key, out var lastShownUtc) &&
            now - lastShownUtc < WarningDialogCooldown)
        {
            return true;
        }

        _warningDialogCooldowns[key] = now;
        return false;
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
        public string? PreferredTerminalShellId { get; set; }
        public bool TerminalVisible { get; set; }
        public double TerminalPanelHeight { get; set; } = 250;
        public List<string> OpenTabPaths { get; set; } = [];
        public string? ActiveTabPath { get; set; }
        public List<RecentFileEntry> RecentFiles { get; set; } = [];
        // False on first launch (settings file didn't exist yet); set to true after the
        // tutorial is dismissed so it never shows again on subsequent launches.
        public bool HasCompletedTutorial { get; set; }
        public string AccentColorMode { get; set; } = "kodo";
        public string CustomAccentHex { get; set; } = "#8C00FF";
        // Personalization - optional; empty/0 means "use OS defaults".
        public string? UserCountry        { get; set; }
        public int     UserHemisphere     { get; set; }
        public string? UserTimezoneOffset { get; set; }
        public string? UserName         { get; set; }
    }

    private sealed record ExtensionScanResult(
        List<LoadedExtension> Extensions,
        List<string> LoadErrors);

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

    private sealed record TutorialStep(
        string SectionTitle,
        string Title,
        string Body,
        string Shortcut,
        string SpotlightTitle,
        string HighlightOne,
        string HighlightTwo,
        string HighlightThree);
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

    // Disabled for unsaved (untitled) files and plain-text files (.txt / .log / .text)
    // so that smart visual features don't activate where no language context exists.
    public bool IsEnabled { get; set; } = true;

    public void UpdateSyntax(LoadedExtension? extension)
    {
        if (extension is null)
        {
            IsEnabled = false;
            _commentLine = "//";
            _commentBlockStart = "/*";
            _commentBlockEnd = "*/";
            _stringDelimiters = ["\"", "'"];
            _multiLineStringDelimiters = [];
        }
        else
        {
            IsEnabled = true;
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
        if (!IsEnabled)
            return;

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

public sealed class InterpolatedStringColorizer : DocumentColorizingTransformer
{
    private static readonly MethodInfo? SetTextRunPropertiesMethod =
        typeof(VisualLineElement).GetMethod("SetTextRunProperties", BindingFlags.Instance | BindingFlags.NonPublic);
    private const string VariableIdentifierBodyPattern =
        "[\\p{L}_][\\p{L}\\p{Nd}_]*";
    private const string CommonStringPrefixPattern =
        "(?i)(?<![\\p{L}\\p{Nd}_])(?:fr|rf|br|rb|ur|ru|cr|rc|f|r|u|b|c)(?=(?:\\\"\\\"\\\"|'''|\\\"|'|#+\\\"))";
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

    private readonly List<SyntaxBrushRule> _rules = [];
    private InterpolationSnapshot? _snapshot;
    private InterpolationSupport _support;
    private IBrush _keywordBrush = Brushes.White;
    private IBrush _punctuationBrush = Brushes.White;

    public bool IsEnabled { get; set; }

    public void UpdateSyntax(LoadedExtension? extension)
    {
        _rules.Clear();
        _snapshot = null;
        _support = default;

        if (extension is null)
        {
            IsEnabled = false;
            _keywordBrush = Brushes.White;
            _punctuationBrush = Brushes.White;
            return;
        }

        IsEnabled = true;
        _keywordBrush = BrushFor(extension, "keyword", "#569CD6");
        _punctuationBrush = BrushFor(extension, "punctuation", "#D4D4D4");
        BuildRules(extension);
        _support = BuildSupport(extension);
    }

    protected override void ColorizeLine(AvaloniaEdit.Document.DocumentLine line)
    {
        if (!IsEnabled || !_support.HasAny)
            return;

        var document = CurrentContext.Document;
        if (document is null || line.Length <= 0)
            return;

        var snapshot = EnsureSnapshot(document.Text ?? string.Empty);
        var lineState = snapshot.GetLineState(line.LineNumber);
        var text = document.GetText(line.Offset, line.Length);
        ScanLine(text, line.Offset, lineState.ActiveInterpolation);
    }

    private void BuildRules(LoadedExtension extension)
    {
        var keywordBrush = BrushFor(extension, "keyword", "#569CD6");
        var typeBrush = BrushFor(extension, "type", "#4EC9B0");
        var numberBrush = BrushFor(extension, "number", "#B5CEA8");
        var functionBrush = BrushFor(extension, "function", "#DCDCAA");
        var namespaceBrush = BrushFor(extension, "namespace", "#4FC1FF");
        var propertyBrush = BrushFor(extension, "property", "#9CDCFE");
        var attributeBrush = BrushFor(extension, "attribute", "#C586C0");
        var operatorBrush = BrushFor(extension, "operator", "#D4D4D4");
        var preprocessorBrush = BrushFor(extension, "preprocessor", "#C586C0");
        var stringBrush = BrushFor(extension, "string", "#CE9178");
        var variableBrush = BrushFor(extension, "variable", "#A0DBFD");
        if (extension.Keywords.Length > 0)
            _rules.Add(new SyntaxBrushRule(BuildTokenRegex(extension.Keywords), keywordBrush));

        if (extension.Types.Length > 0)
            _rules.Add(new SyntaxBrushRule(BuildTokenRegex(extension.Types), typeBrush));

        if (extension.Namespaces.Length > 0)
            _rules.Add(new SyntaxBrushRule(BuildTokenRegex(extension.Namespaces), namespaceBrush));

        if (extension.Properties.Length > 0)
            _rules.Add(new SyntaxBrushRule(BuildTokenRegex(extension.Properties), propertyBrush));

        if (extension.Functions.Length > 0)
            _rules.Add(new SyntaxBrushRule(BuildTokenRegex(extension.Functions), functionBrush));

        if (extension.StringDelimiters.Contains("\"") ||
            extension.StringDelimiters.Contains("'") ||
            extension.MultiLineStringDelimiters.Contains("\"\"\"") ||
            extension.MultiLineStringDelimiters.Contains("'''"))
        {
            _rules.Add(new SyntaxBrushRule(
                new Regex(CommonStringPrefixPattern, RegexOptions.Compiled),
                stringBrush));
        }

        // XML/MSBuild version string rule - only active for XML-family languages
        // (identified by <!-- block comment syntax). Matches the full content of an
        // element body when it looks like a version string (e.g. 1.0.0, v1.2.3-DEV,
        // net10.0, 12.0.1) so dots, dashes, and letter suffixes are not fragmented
        // across number/operator/variable token colours. Inserted before the number
        // rule so it claims the whole token first.
        if (extension.CommentBlockStart == "<!--")
        {
            _rules.Add(new SyntaxBrushRule(
                new Regex(@"(?<=>)\s*v?\d[\d._-]*[A-Za-z0-9]*\s*(?=<)", RegexOptions.Compiled),
                numberBrush));
        }

        _rules.Add(new SyntaxBrushRule(
            new Regex(@"(?<![\p{L}\p{Nd}_])(?:0[xX][0-9A-Fa-f]+|0[bB][01]+|0[oO][0-7]+|\d+(?:\.\d+)?(?:[eE][+\-]?\d+)?)(?![\p{L}\p{Nd}_])", RegexOptions.Compiled),
            numberBrush));
        // Property and namespace rules fire on dot-separated identifiers (e.g. foo.Bar).
        // In XML-family languages dots appear in file paths and version strings between
        // element tags, so these rules would wrongly colour path segments as namespace/
        // property tokens. Skip them entirely for XML - no chained member access exists.
        if (extension.CommentBlockStart != "<!--")
        {
            _rules.Add(new SyntaxBrushRule(
                new Regex(@"(?<=\.|->|::)[\p{L}_][\p{L}\p{Nd}_:-]*", RegexOptions.Compiled),
                propertyBrush));
            _rules.Add(new SyntaxBrushRule(
                new Regex(@"(?<![\p{L}\p{Nd}_])[\p{L}_][\p{L}\p{Nd}_]*(?=\.)", RegexOptions.Compiled),
                namespaceBrush));
        }
        _rules.Add(new SyntaxBrushRule(
            new Regex(@"(?<![\p{L}\p{Nd}_])[@#][\p{L}_][\p{L}\p{Nd}_-]*", RegexOptions.Compiled),
            preprocessorBrush));
        _rules.Add(new SyntaxBrushRule(
            new Regex(@"(?<=\[)[\p{L}_][\p{L}\p{Nd}_:.]*(?=[,\]\(])|(?<=<)[\p{L}_][\p{L}\p{Nd}_:-]*(?=[^>]*>)", RegexOptions.Compiled),
            attributeBrush));
        _rules.Add(new SyntaxBrushRule(
            new Regex(@"(?<![\p{L}\p{Nd}_])[\p{L}_][\p{L}\p{Nd}_]*(?=\s*\()", RegexOptions.Compiled),
            functionBrush));
        _rules.Add(new SyntaxBrushRule(
            new Regex(@"=>|->|::|\+\+|--|\+=|-=|\*=|/=|%=|&&|\|\||<<|>>|<=|>=|==|!=|=|\+|-|\*|/|%|!|\?|:|<|>|&|\||\^|~", RegexOptions.Compiled),
            operatorBrush));
        // For XML-family languages the dot appears in file paths and version strings
        // between element tags, so exclude it from punctuation to avoid grey fragments.
        _rules.Add(new SyntaxBrushRule(
            new Regex(extension.CommentBlockStart == "<!--" ? @"[;,]" : @"[;,.]", RegexOptions.Compiled),
            _punctuationBrush));
        _rules.Add(new SyntaxBrushRule(
            new Regex(@"(?<=\b(?:using|import|include|require|use|from)\b\s+(?:[\p{L}_][\p{L}\p{Nd}_./\\]*\s*[./\\]\s*)?)[\p{L}_][\p{L}\p{Nd}_]*(?=\s*(?:;|$))", RegexOptions.Compiled),
            namespaceBrush));
        _rules.Add(new SyntaxBrushRule(
            BuildVariableRegex(extension.Keywords.Concat(extension.Types).Concat(extension.Functions).Concat(extension.Properties).Concat(extension.Namespaces)),
            variableBrush));
    }

    private static Regex BuildVariableRegex(IEnumerable<string> reservedTokens)
    {
        var reserved = reservedTokens
            .Where(token => !string.IsNullOrWhiteSpace(token))
            .Distinct(StringComparer.Ordinal)
            .OrderByDescending(token => token.Length)
            .Select(Regex.Escape)
            .ToArray();

        var reservedPrefix = reserved.Length > 0
            ? $"(?!{string.Join("|", reserved.Select(r => r + "(?![\\p{L}\\p{Nd}_])"))})"
            : string.Empty;

        return new Regex(
            $"(?<![.\\p{{L}}\\p{{Nd}}_]){reservedPrefix}{VariableIdentifierBodyPattern}(?!\\s*[\\.(\"'`]|[\\p{{L}}\\p{{Nd}}_])",
            RegexOptions.Compiled);
    }

    private static InterpolationSupport BuildSupport(LoadedExtension extension)
    {
        var stringDelimiters = extension.StringDelimiters
            .Where(d => !string.IsNullOrEmpty(d))
            .ToHashSet(StringComparer.Ordinal);
        var multiLineDelimiters = extension.MultiLineStringDelimiters
            .Where(d => !string.IsNullOrEmpty(d))
            .ToHashSet(StringComparer.Ordinal);

        var supportsJavaScriptTemplate = multiLineDelimiters.Contains("`") &&
            !string.Equals(extension.Id, "markdown-kodo-extension", StringComparison.OrdinalIgnoreCase);
        var supportsPythonStyleInterpolation =
            string.Equals(extension.CommentLine, "#", StringComparison.Ordinal) &&
            !extension.DisableSingleQuoteStrings &&
            (stringDelimiters.Contains("\"") || stringDelimiters.Contains("'")) &&
            (multiLineDelimiters.Contains("\"\"\"") || multiLineDelimiters.Contains("'''"));
        var supportsDollarPrefixedInterpolation =
            extension.DisableSingleQuoteStrings &&
            stringDelimiters.Contains("\"");

        return new InterpolationSupport(supportsJavaScriptTemplate, supportsPythonStyleInterpolation, supportsDollarPrefixedInterpolation);
    }

    private InterpolationSnapshot EnsureSnapshot(string text)
    {
        if (_snapshot is { } snapshot && string.Equals(snapshot.Text, text, StringComparison.Ordinal))
            return snapshot;

        _snapshot = BuildSnapshot(text);
        return _snapshot;
    }

    private InterpolationSnapshot BuildSnapshot(string text)
    {
        var lineStates = new List<InterpolationLineState> { new(null) };
        ActiveInterpolation? active = null;

        for (var index = 0; index < text.Length; index++)
        {
            if (TryConsumeLineBreak(text, ref index))
            {
                lineStates.Add(new(active));
                continue;
            }

            if (active is not null)
            {
                var activeValue = active.Value;
                if (IsInterpolationTerminator(text, index, activeValue))
                {
                    index += activeValue.Quote.Length - 1;
                    active = null;
                }

                continue;
            }

            if (TryMatchInterpolationStart(text, index, out var interpolation, out var contentStart))
            {
                if (interpolation.CanSpanMultipleLines)
                {
                    active = interpolation;
                    index = contentStart - 1;
                }
            }
        }

        return new InterpolationSnapshot(text, lineStates);
    }

    private void ScanLine(string text, int lineOffset, ActiveInterpolation? initialState)
    {
        var index = 0;
        var active = initialState;

        while (index < text.Length)
        {
            if (active is not null)
            {
                var activeValue = active.Value;
                var closingIndex = FindClosingIndex(text, index, activeValue);
                var stringEnd = closingIndex >= 0 ? closingIndex : text.Length;
                ColorizeInterpolationSegments(text, lineOffset, activeValue, index, stringEnd);

                if (closingIndex >= 0)
                {
                    index = closingIndex + activeValue.Quote.Length;
                    active = null;
                    continue;
                }

                break;
            }

            if (TryMatchInterpolationStart(text, index, out var interpolation, out var contentStart))
            {
                ApplyInterpolationPrefix(lineOffset, index, contentStart, interpolation.Quote.Length);
                var closingIndex = FindClosingIndex(text, contentStart, interpolation);
                var stringEnd = closingIndex >= 0 ? closingIndex : text.Length;
                ColorizeInterpolationSegments(text, lineOffset, interpolation, contentStart, stringEnd);

                if (closingIndex >= 0)
                {
                    index = closingIndex + interpolation.Quote.Length;
                    continue;
                }

                break;
            }

            index++;
        }
    }

    private void ColorizeInterpolationSegments(string text, int lineOffset, ActiveInterpolation interpolation, int contentStart, int stringEnd)
    {
        if (interpolation.Kind == InterpolationKind.JavaScriptTemplate)
        {
            for (var index = contentStart; index < stringEnd - 1; index++)
            {
                if (text[index] == '\\')
                {
                    index++;
                    continue;
                }

                if (text[index] == '$' && text[index + 1] == '{')
                {
                    var closeIndex = FindClosingBrace(text, index + 2, allowDoubledEscapes: false, limit: stringEnd);
                    if (closeIndex > index + 1)
                    {
                        ApplyBrush(lineOffset + index + 1, lineOffset + index + 2, GetRainbowBrush(0));
                        ApplyBrush(lineOffset + closeIndex, lineOffset + closeIndex + 1, GetRainbowBrush(0));
                        ApplyExpressionRules(text, lineOffset, index + 2, closeIndex);
                        index = closeIndex;
                    }
                }
            }

            return;
        }

        for (var index = contentStart; index < stringEnd; index++)
        {
            if (text[index] == '{')
            {
                if (index + 1 < stringEnd && text[index + 1] == '{')
                {
                    index++;
                    continue;
                }

                var closeIndex = FindClosingBrace(text, index + 1, allowDoubledEscapes: true, limit: stringEnd);
                if (closeIndex > index)
                {
                    ApplyBrush(lineOffset + index, lineOffset + index + 1, GetRainbowBrush(0));
                    ApplyBrush(lineOffset + closeIndex, lineOffset + closeIndex + 1, GetRainbowBrush(0));
                    ApplyExpressionRules(text, lineOffset, index + 1, closeIndex);
                    index = closeIndex;
                }
            }
            else if (!interpolation.IsVerbatim && text[index] == '\\')
            {
                index++;
            }
        }
    }

    private void ApplyExpressionRules(string text, int lineOffset, int expressionStart, int expressionEnd)
    {
        if (expressionEnd <= expressionStart)
            return;

        var expressionText = text.Substring(expressionStart, expressionEnd - expressionStart);
        foreach (var rule in _rules)
        {
            foreach (Match match in rule.Regex.Matches(expressionText))
            {
                if (!match.Success || match.Length <= 0)
                    continue;

                ApplyBrush(
                    lineOffset + expressionStart + match.Index,
                    lineOffset + expressionStart + match.Index + match.Length,
                    rule.Brush);
            }
        }

        ApplyRainbowBrackets(text, lineOffset, expressionStart, expressionEnd);
    }

    private void ApplyInterpolationPrefix(int lineOffset, int startIndex, int contentStart, int quoteLength)
    {
        var prefixLength = contentStart - startIndex - quoteLength;
        if (prefixLength <= 0)
            return;

        ApplyBrush(lineOffset + startIndex, lineOffset + startIndex + prefixLength, _keywordBrush);
    }

    private void ApplyRainbowBrackets(string text, int lineOffset, int expressionStart, int expressionEnd)
    {
        var stack = new Stack<char>();
        for (var index = expressionStart; index < expressionEnd; index++)
        {
            var ch = text[index];
            if (OpeningToClosing.ContainsKey(ch))
            {
                ApplyBrush(lineOffset + index, lineOffset + index + 1, GetRainbowBrush(stack.Count));
                stack.Push(ch);
                continue;
            }

            if (ClosingToOpening.TryGetValue(ch, out var opening) &&
                stack.Count > 0 &&
                stack.Peek() == opening)
            {
                ApplyBrush(lineOffset + index, lineOffset + index + 1, GetRainbowBrush(stack.Count - 1));
                stack.Pop();
            }
        }
    }

    private bool TryMatchInterpolationStart(string text, int index, out ActiveInterpolation interpolation, out int contentStart)
    {
        if (_support.SupportsJavaScriptTemplate && text[index] == '`')
        {
            interpolation = new ActiveInterpolation(InterpolationKind.JavaScriptTemplate, "`", false, true);
            contentStart = index + 1;
            return true;
        }

        if (_support.SupportsPythonStyle && TryMatchPythonFStringStart(text, index, out var quote, out contentStart))
        {
            interpolation = new ActiveInterpolation(InterpolationKind.PythonFString, quote, false, quote.Length > 1);
            return true;
        }

        if (_support.SupportsDollarPrefixed && TryMatchDollarPrefixedStart(text, index, out var isVerbatim, out contentStart))
        {
            interpolation = new ActiveInterpolation(InterpolationKind.DollarPrefixed, "\"", isVerbatim, isVerbatim);
            return true;
        }

        interpolation = default;
        contentStart = 0;
        return false;
    }

    private static int FindClosingIndex(string text, int index, ActiveInterpolation interpolation)
    {
        for (var i = index; i <= text.Length - interpolation.Quote.Length; i++)
        {
            if (!interpolation.IsVerbatim && text[i] == '\\')
            {
                i++;
                continue;
            }

            if (interpolation.IsVerbatim && interpolation.Quote == "\"" &&
                i + 1 < text.Length &&
                text[i] == '"' &&
                text[i + 1] == '"')
            {
                i++;
                continue;
            }

            if (IsInterpolationTerminator(text, i, interpolation))
                return i;
        }

        return -1;
    }

    private static bool IsInterpolationTerminator(string text, int index, ActiveInterpolation interpolation)
    {
        if (index < 0 || index + interpolation.Quote.Length > text.Length)
            return false;

        if (interpolation.Kind == InterpolationKind.JavaScriptTemplate)
            return text[index] == '`';

        if (!interpolation.IsVerbatim && IsEscaped(text, index))
            return false;

        return string.CompareOrdinal(text, index, interpolation.Quote, 0, interpolation.Quote.Length) == 0;
    }

    private static int FindClosingBrace(string text, int index, bool allowDoubledEscapes, int limit)
    {
        var depth = 0;
        for (var i = index; i < limit; i++)
        {
            if (allowDoubledEscapes && i + 1 < limit && text[i] == '{' && text[i + 1] == '{')
            {
                i++;
                continue;
            }

            if (allowDoubledEscapes && i + 1 < limit && text[i] == '}' && text[i + 1] == '}')
            {
                i++;
                continue;
            }

            if (text[i] == '{')
            {
                depth++;
            }
            else if (text[i] == '}')
            {
                if (depth == 0)
                    return i;

                depth--;
            }
        }

        return -1;
    }

    private static bool TryMatchPythonFStringStart(string text, int index, out string quote, out int contentStart)
    {
        quote = string.Empty;
        contentStart = 0;

        if (index >= text.Length)
            return false;

        var prefixLength = 0;
        if ((text[index] == 'f' || text[index] == 'F') &&
            index + 1 < text.Length &&
            (text[index + 1] == 'r' || text[index + 1] == 'R'))
        {
            prefixLength = 2;
        }
        else if ((text[index] == 'r' || text[index] == 'R') &&
                 index + 1 < text.Length &&
                 (text[index + 1] == 'f' || text[index + 1] == 'F'))
        {
            prefixLength = 2;
        }
        else if (text[index] == 'f' || text[index] == 'F')
        {
            prefixLength = 1;
        }

        if (prefixLength == 0)
            return false;

        var quoteIndex = index + prefixLength;
        if (quoteIndex + 2 < text.Length && text[quoteIndex] == '"' && text[quoteIndex + 1] == '"' && text[quoteIndex + 2] == '"')
        {
            quote = "\"\"\"";
            contentStart = quoteIndex + 3;
            return true;
        }

        if (quoteIndex + 2 < text.Length && text[quoteIndex] == '\'' && text[quoteIndex + 1] == '\'' && text[quoteIndex + 2] == '\'')
        {
            quote = "'''";
            contentStart = quoteIndex + 3;
            return true;
        }

        if (quoteIndex < text.Length && (text[quoteIndex] == '"' || text[quoteIndex] == '\''))
        {
            quote = text[quoteIndex].ToString();
            contentStart = quoteIndex + 1;
            return true;
        }

        return false;
    }

    private static bool TryMatchDollarPrefixedStart(string text, int index, out bool isVerbatim, out int contentStart)
    {
        isVerbatim = false;
        contentStart = 0;

        if (index + 1 >= text.Length)
            return false;

        if (text[index] == '$' && text[index + 1] == '"')
        {
            contentStart = index + 2;
            return true;
        }

        if (index + 2 < text.Length && text[index] == '$' && text[index + 1] == '@' && text[index + 2] == '"')
        {
            isVerbatim = true;
            contentStart = index + 3;
            return true;
        }

        if (index + 2 < text.Length && text[index] == '@' && text[index + 1] == '$' && text[index + 2] == '"')
        {
            isVerbatim = true;
            contentStart = index + 3;
            return true;
        }

        return false;
    }

    private static bool TryConsumeLineBreak(string text, ref int index)
    {
        if (text[index] == '\r')
        {
            if (index + 1 < text.Length && text[index + 1] == '\n')
                index++;

            return true;
        }

        return text[index] == '\n';
    }

    private static bool IsEscaped(string text, int index)
    {
        var slashCount = 0;
        for (var i = index - 1; i >= 0 && text[i] == '\\'; i--)
            slashCount++;

        return slashCount % 2 == 1;
    }

    private static IBrush GetRainbowBrush(int depth) => RainbowBrushes[Math.Abs(depth) % RainbowBrushes.Length];

    private void ApplyBrush(int startOffset, int endOffset, IBrush brush)
    {
        if (endOffset <= startOffset)
            return;

        ChangeLinePart(startOffset, endOffset, element =>
        {
            var properties = element.TextRunProperties.Clone();
            properties.SetForegroundBrush(brush);
            SetTextRunPropertiesMethod?.Invoke(element, [properties]);
        });
    }

    private static IBrush BrushFor(LoadedExtension extension, string tokenName, string fallback)
    {
        var hex = extension.ColorTokens.TryGetValue(tokenName, out var value) ? value : fallback;
        return Brush.Parse(hex);
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

    private readonly record struct SyntaxBrushRule(Regex Regex, IBrush Brush);
    private readonly record struct InterpolationSupport(bool SupportsJavaScriptTemplate, bool SupportsPythonStyle, bool SupportsDollarPrefixed)
    {
        public bool HasAny => SupportsJavaScriptTemplate || SupportsPythonStyle || SupportsDollarPrefixed;
    }

    private sealed record InterpolationSnapshot(string Text, List<InterpolationLineState> LineStates)
    {
        public InterpolationLineState GetLineState(int lineNumber) =>
            lineNumber > 0 && lineNumber <= LineStates.Count
                ? LineStates[lineNumber - 1]
                : new InterpolationLineState(null);
    }

    private readonly record struct InterpolationLineState(ActiveInterpolation? ActiveInterpolation);
    private readonly record struct ActiveInterpolation(InterpolationKind Kind, string Quote, bool IsVerbatim, bool CanSpanMultipleLines);

    private enum InterpolationKind
    {
        JavaScriptTemplate,
        PythonFString,
        DollarPrefixed
    }
}

public sealed class MarkdownColorizer : DocumentColorizingTransformer
{
    private const string CommonStringPrefixPattern =
        "(?i)(?<![\\p{L}\\p{Nd}_])(?:fr|rf|br|rb|ur|ru|cr|rc|f|r|u|b|c)(?=(?:\\\"\\\"\\\"|'''|\\\"|'|#+\\\"))";
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
    private static readonly Regex HeadingRegex = new(@"^(?<indent>\s{0,3})(?<marker>#{1,6})(?<space>\s+)(?<content>.+?)\s*(?:#+\s*)?$", RegexOptions.Compiled);
    private static readonly Regex SetextHeadingUnderlineRegex = new(@"^\s{0,3}(?:=+|-+)\s*$", RegexOptions.Compiled);
    private static readonly Regex HorizontalRuleRegex = new(@"^\s{0,3}(?:([-*_])(?:\s*\1){2,})\s*$", RegexOptions.Compiled);
    private static readonly Regex OrderedListRegex = new(@"^\s{0,3}\d+\.\s+", RegexOptions.Compiled);
    private static readonly Regex UnorderedListRegex = new(@"^\s{0,3}[-+*]\s+", RegexOptions.Compiled);
    private static readonly Regex TaskListRegex = new(@"^(\s{0,3}[-+*]\s+)\[(?<state>[ xX])\]\s+", RegexOptions.Compiled);
    private static readonly Regex BlockquoteRegex = new(@"^(?<indent>\s{0,3})(?<markers>(?:>\s?)+)", RegexOptions.Compiled);
    private static readonly Regex TablePipeRegex = new(@"\|", RegexOptions.Compiled);
    private static readonly Regex TableAlignmentRegex = new(@"^\s*\|?(?:\s*:?-{3,}:?\s*\|)+\s*:?-{3,}:?\s*\|?\s*$", RegexOptions.Compiled);
    private static readonly Regex InlineCodeRegex = new(@"(?<!`)(`+)(?!`)(?<content>.*?[^`])\1(?!`)", RegexOptions.Compiled);
    private static readonly Regex LinkRegex = new(@"!\[(?<alt>[^\]]*)\]\((?<image>[^)\r\n]+)\)|\[(?<label>[^\]]+)\]\((?<url>[^)\r\n]+)\)|\[(?<refLabel>[^\]]+)\]\[(?<refId>[^\]]*)\]", RegexOptions.Compiled);
    private static readonly Regex LinkReferenceDefinitionRegex = new(@"^\s{0,3}\[(?<id>[^\]]+)\]:\s*(?<url>\S+)(?:\s+(?<title>""[^""]*""|'[^']*'|\([^)]*\)))?\s*$", RegexOptions.Compiled);
    private static readonly Regex AutoLinkRegex = new(@"(?<!\()https?://[^\s)>\]]+|<(?<auto>(?:https?://|mailto:)[^>\r\n]+)>", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex StrongEmphasisRegex = new(@"(?<!\*)\*\*\*(?=\S)(?<content>.+?)(?<=\S)\*\*\*(?!\*)|(?<!_)___(?=\S)(?<content2>.+?)(?<=\S)___(?!_)", RegexOptions.Compiled);
    private static readonly Regex StrongRegex = new(@"(?<!\*)\*\*(?=\S)(?<content>.+?)(?<=\S)\*\*(?!\*)|(?<!_)__(?=\S)(?<content2>.+?)(?<=\S)__(?!_)", RegexOptions.Compiled);
    private static readonly Regex EmphasisRegex = new(@"(?<!\*)\*(?=\S)(?<content>.+?)(?<=\S)\*(?!\*)|(?<!_)_(?=\S)(?<content2>.+?)(?<=\S)_(?!_)", RegexOptions.Compiled);
    private static readonly Regex StrikethroughRegex = new(@"~~(?=\S)(?<content>.+?)(?<=\S)~~", RegexOptions.Compiled);
    private static readonly Regex HtmlCommentRegex = new(@"<!--.*?-->", RegexOptions.Compiled);

    private MarkdownSnapshot? _snapshot;
    private Func<string, LoadedExtension?>? _languageResolver;
    private Func<string, LoadedExtension?>? _inlineLanguageResolver;
    private readonly Dictionary<string, EmbeddedSyntaxProfile> _embeddedProfileCache = new(StringComparer.OrdinalIgnoreCase);
    private IBrush _keywordBrush = Brushes.White;
    private IBrush _typeBrush = Brushes.White;
    private IBrush _stringBrush = Brushes.White;
    private IBrush _commentBrush = Brushes.White;
    private IBrush _operatorBrush = Brushes.White;
    private IBrush _punctuationBrush = Brushes.White;
    private IBrush _variableBrush = Brushes.White;
    private IBrush _mutedBrush = Brushes.White;

    public bool IsEnabled { get; private set; }

    public void UpdateSyntax(LoadedExtension? extension, Func<string, LoadedExtension?>? languageResolver, Func<string, LoadedExtension?>? inlineLanguageResolver)
    {
        _snapshot = null;
        _embeddedProfileCache.Clear();
        _languageResolver = languageResolver;
        _inlineLanguageResolver = inlineLanguageResolver;

        if (extension is null || !IsMarkdownExtension(extension))
        {
            IsEnabled = false;
            return;
        }

        IsEnabled = true;
        _keywordBrush = BrushFor(extension, "keyword", "#569CD6");
        _typeBrush = BrushFor(extension, "type", "#BAE6FD");
        _stringBrush = BrushFor(extension, "string", "#CE9178");
        _commentBrush = BrushFor(extension, "comment", "#6A9955");
        _operatorBrush = BrushFor(extension, "operator", "#C586C0");
        _punctuationBrush = BrushFor(extension, "punctuation", "#569CD6");
        _variableBrush = BrushFor(extension, "variable", "#F4F4F4");
        _mutedBrush = Brush.Parse("#7A7A7A");
    }

    protected override void ColorizeLine(AvaloniaEdit.Document.DocumentLine line)
    {
        if (!IsEnabled)
            return;

        var document = CurrentContext.Document;
        if (document is null)
            return;

        var text = document.GetText(line.Offset, line.Length);
        var state = EnsureSnapshot(document.Text ?? string.Empty).GetLineState(line.LineNumber);

        if (state.Delimiter is not null)
        {
            ColorizeFenceDelimiter(text, line.Offset, state.Delimiter.Value);
            return;
        }

        if (state.ActiveFence is { } fence)
        {
            ColorizeEmbeddedCode(text, line.Offset, fence);
            return;
        }

        ColorizeMarkdownLine(text, line.Offset);
    }

    private void ColorizeMarkdownLine(string text, int lineOffset)
    {
        if (string.IsNullOrWhiteSpace(text))
            return;

        var protectedRanges = new bool[text.Length];

        var headingMatch = HeadingRegex.Match(text);
        if (headingMatch.Success)
        {
            ApplyBrush(lineOffset, headingMatch.Groups["marker"].Index, headingMatch.Groups["marker"].Index + headingMatch.Groups["marker"].Length, _keywordBrush);
            ApplyBrush(lineOffset, headingMatch.Groups["content"].Index, headingMatch.Groups["content"].Index + headingMatch.Groups["content"].Length, _keywordBrush);
            var trailingMarkerStart = headingMatch.Groups["content"].Index + headingMatch.Groups["content"].Length;
            if (trailingMarkerStart < text.Length)
                ApplyBrush(lineOffset, trailingMarkerStart, text.Length, _mutedBrush);
            return;
        }

        if (SetextHeadingUnderlineRegex.IsMatch(text))
        {
            ApplyBrush(lineOffset, 0, text.Length, _keywordBrush);
            return;
        }

        if (HorizontalRuleRegex.IsMatch(text))
        {
            ApplyBrush(lineOffset, 0, text.Length, _mutedBrush);
            return;
        }

        if (TableAlignmentRegex.IsMatch(text))
        {
            ApplyBrush(lineOffset, 0, text.Length, _mutedBrush);
            foreach (Match match in TablePipeRegex.Matches(text))
                ApplyBrush(lineOffset, match.Index, match.Index + match.Length, _punctuationBrush);
            return;
        }

        var linkReferenceDefinitionMatch = LinkReferenceDefinitionRegex.Match(text);
        if (linkReferenceDefinitionMatch.Success)
        {
            ApplyBrush(lineOffset, 0, text.Length, _variableBrush);
            ApplyBrush(lineOffset, linkReferenceDefinitionMatch.Groups["id"].Index, linkReferenceDefinitionMatch.Groups["id"].Index + linkReferenceDefinitionMatch.Groups["id"].Length, _typeBrush);
            ApplyBrush(lineOffset, linkReferenceDefinitionMatch.Groups["url"].Index, linkReferenceDefinitionMatch.Groups["url"].Index + linkReferenceDefinitionMatch.Groups["url"].Length, _stringBrush);
            if (linkReferenceDefinitionMatch.Groups["title"].Success)
            {
                ApplyBrush(lineOffset, linkReferenceDefinitionMatch.Groups["title"].Index, linkReferenceDefinitionMatch.Groups["title"].Index + linkReferenceDefinitionMatch.Groups["title"].Length, _stringBrush);
            }
            foreach (var ch in new[] { '[', ']', ':' })
            {
                foreach (var index in AllIndexesOf(text, ch))
                    ApplyBrush(lineOffset, index, index + 1, _punctuationBrush);
            }
            return;
        }

        var taskMatch = TaskListRegex.Match(text);
        if (taskMatch.Success)
        {
            ApplyBrush(lineOffset, 0, taskMatch.Groups[1].Length, _keywordBrush);
            ApplyBrush(lineOffset, taskMatch.Groups[1].Length, taskMatch.Length, _punctuationBrush);
            var stateGroup = taskMatch.Groups["state"];
            ApplyBrush(lineOffset, stateGroup.Index, stateGroup.Index + stateGroup.Length, _stringBrush);
            MarkRange(protectedRanges, 0, taskMatch.Length);
        }
        else
        {
            var orderedMatch = OrderedListRegex.Match(text);
            if (orderedMatch.Success)
            {
                ApplyBrush(lineOffset, 0, orderedMatch.Length, _keywordBrush);
                MarkRange(protectedRanges, 0, orderedMatch.Length);
            }

            var unorderedMatch = UnorderedListRegex.Match(text);
            if (unorderedMatch.Success)
            {
                ApplyBrush(lineOffset, 0, unorderedMatch.Length, _keywordBrush);
                MarkRange(protectedRanges, 0, unorderedMatch.Length);
            }
        }

        var quoteMatch = BlockquoteRegex.Match(text);
        if (quoteMatch.Success)
        {
            var markers = quoteMatch.Groups["markers"];
            ApplyBrush(lineOffset, markers.Index, markers.Index + markers.Length, _commentBrush);
            MarkRange(protectedRanges, markers.Index, markers.Index + markers.Length);
        }

        foreach (Match match in HtmlCommentRegex.Matches(text))
        {
            if (!TryReserveRange(protectedRanges, match.Index, match.Index + match.Length))
                continue;

            ApplyBrush(lineOffset, match.Index, match.Index + match.Length, _commentBrush);
        }

        foreach (Match match in TablePipeRegex.Matches(text))
        {
            ApplyBrush(lineOffset, match.Index, match.Index + match.Length, _mutedBrush);
            MarkRange(protectedRanges, match.Index, match.Index + match.Length);
        }

        foreach (Match match in LinkRegex.Matches(text))
        {
            if (!TryReserveRange(protectedRanges, match.Index, match.Index + match.Length))
                continue;

            ApplyBrush(lineOffset, match.Index, match.Index + match.Length, _variableBrush);

            if (match.Value.StartsWith("!", StringComparison.Ordinal))
                ApplyBrush(lineOffset, match.Index, match.Index + 1, _punctuationBrush);

            var labelGroup = match.Groups["label"];
            if (labelGroup.Success)
                ApplyBrush(lineOffset, labelGroup.Index, labelGroup.Index + labelGroup.Length, _typeBrush);

            var altGroup = match.Groups["alt"];
            if (altGroup.Success)
                ApplyBrush(lineOffset, altGroup.Index, altGroup.Index + altGroup.Length, _typeBrush);

            var urlGroup = match.Groups["url"];
            if (urlGroup.Success)
                ApplyBrush(lineOffset, urlGroup.Index, urlGroup.Index + urlGroup.Length, _stringBrush);

            var imageGroup = match.Groups["image"];
            if (imageGroup.Success)
                ApplyBrush(lineOffset, imageGroup.Index, imageGroup.Index + imageGroup.Length, _stringBrush);

            var refIdGroup = match.Groups["refId"];
            if (refIdGroup.Success)
                ApplyBrush(lineOffset, refIdGroup.Index, refIdGroup.Index + refIdGroup.Length, _stringBrush);

            foreach (var ch in new[] { '[', ']', '(', ')' })
            {
                foreach (var index in AllIndexesOf(match.Value, ch))
                {
                    var absolute = match.Index + index;
                    ApplyBrush(lineOffset, absolute, absolute + 1, _punctuationBrush);
                }
            }
        }

        foreach (Match match in InlineCodeRegex.Matches(text))
        {
            if (!TryReserveRange(protectedRanges, match.Index, match.Index + match.Length))
                continue;

            ColorizeInlineCode(match, lineOffset);
        }

        foreach (Match match in AutoLinkRegex.Matches(text))
        {
            if (!TryReserveRange(protectedRanges, match.Index, match.Index + match.Length))
                continue;

            ApplyBrush(lineOffset, match.Index, match.Index + match.Length, _stringBrush);
            if (match.Value.StartsWith("<", StringComparison.Ordinal) && match.Value.EndsWith(">", StringComparison.Ordinal))
            {
                ApplyBrush(lineOffset, match.Index, match.Index + 1, _punctuationBrush);
                ApplyBrush(lineOffset, match.Index + match.Length - 1, match.Index + match.Length, _punctuationBrush);
            }
        }

        ApplyDelimitedMarkdownRegex(text, lineOffset, protectedRanges, StrongEmphasisRegex, _operatorBrush);
        ApplyDelimitedMarkdownRegex(text, lineOffset, protectedRanges, StrongRegex, _keywordBrush);
        ApplyDelimitedMarkdownRegex(text, lineOffset, protectedRanges, EmphasisRegex, _operatorBrush);
        ApplyDelimitedMarkdownRegex(text, lineOffset, protectedRanges, StrikethroughRegex, _mutedBrush);
    }

    private void ColorizeFenceDelimiter(string text, int lineOffset, FenceDelimiterInfo delimiter)
    {
        ApplyBrush(lineOffset, 0, text.Length, _mutedBrush);

        var trimmedStart = text.Length - text.TrimStart().Length;
        var markerLength = delimiter.MarkerLength;
        ApplyBrush(lineOffset, trimmedStart, trimmedStart + markerLength, _punctuationBrush);

        if (!string.IsNullOrWhiteSpace(delimiter.LanguageLabel))
        {
            var languageStart = trimmedStart + markerLength;
            while (languageStart < text.Length && char.IsWhiteSpace(text[languageStart]))
                languageStart++;

            var languageEnd = languageStart + delimiter.LanguageLabel.Length;
            if (languageEnd <= text.Length)
                ApplyBrush(lineOffset, languageStart, languageEnd, _typeBrush);
        }
    }

    private void ColorizeEmbeddedCode(string text, int lineOffset, FenceState fence)
    {
        if (string.IsNullOrEmpty(text))
            return;

        if (fence.Profile is null)
        {
            ApplyBrush(lineOffset, 0, text.Length, _variableBrush);
            return;
        }

        ColorizeEmbeddedSegment(text, lineOffset, fence.Profile, fence.BracketStack);
    }

    private void ColorizeEmbeddedSegment(string text, int lineOffset, EmbeddedSyntaxProfile profile, IReadOnlyList<char>? initialBracketStack = null)
    {
        var protectedRanges = BuildEmbeddedProtectedRanges(text, profile, lineOffset, applyColors: true);

        foreach (var rule in profile.TokenRules)
        {
            foreach (Match match in rule.Regex.Matches(text))
            {
                if (IsProtected(protectedRanges, match.Index, match.Index + match.Length))
                    continue;

                ApplyBrush(lineOffset, match.Index, match.Index + match.Length, rule.Brush);
            }
        }

        var stack = initialBracketStack is { Count: > 0 }
            ? new Stack<char>(initialBracketStack.Reverse())
            : new Stack<char>();
        for (var index = 0; index < text.Length; index++)
        {
            if (protectedRanges[index])
                continue;

            var ch = text[index];
            if (OpeningToClosing.ContainsKey(ch))
            {
                ApplyBrush(lineOffset, index, index + 1, GetRainbowBrush(stack.Count));
                stack.Push(ch);
                continue;
            }

            if (ClosingToOpening.TryGetValue(ch, out var opening) &&
                stack.Count > 0 &&
                stack.Peek() == opening)
            {
                ApplyBrush(lineOffset, index, index + 1, GetRainbowBrush(stack.Count - 1));
                stack.Pop();
            }
        }
    }

    private bool[] BuildEmbeddedProtectedRanges(string text, EmbeddedSyntaxProfile profile, int lineOffset, bool applyColors)
    {
        var protectedRanges = new bool[text.Length];

        if (profile.SingleLineCommentRegex is { } commentRegex)
        {
            foreach (Match match in commentRegex.Matches(text))
            {
                if (!TryReserveRange(protectedRanges, match.Index, match.Index + match.Length))
                    continue;

                if (applyColors)
                    ApplyBrush(lineOffset, match.Index, match.Index + match.Length, profile.CommentBrush);
            }
        }

        foreach (var regex in profile.StringRegexes)
        {
            foreach (Match match in regex.Matches(text))
            {
                if (!TryReserveRange(protectedRanges, match.Index, match.Index + match.Length))
                    continue;

                if (applyColors)
                    ApplyBrush(lineOffset, match.Index, match.Index + match.Length, profile.StringBrush);
            }
        }

        return protectedRanges;
    }

    private void ApplyDelimitedMarkdownRegex(string text, int lineOffset, bool[] protectedRanges, Regex regex, IBrush brush)
    {
        foreach (Match match in regex.Matches(text))
        {
            if (!TryReserveRange(protectedRanges, match.Index, match.Index + match.Length))
                continue;

            ApplyBrush(lineOffset, match.Index, match.Index + match.Length, brush);
        }
    }

    private MarkdownSnapshot EnsureSnapshot(string text)
    {
        if (_snapshot is { } snapshot && string.Equals(snapshot.Text, text, StringComparison.Ordinal))
            return snapshot;

        _snapshot = BuildSnapshot(text);
        return _snapshot;
    }

    private MarkdownSnapshot BuildSnapshot(string text)
    {
        var lines = text.Split(["\r\n", "\n", "\r"], StringSplitOptions.None);
        var states = new List<MarkdownLineState>(lines.Length);
        FenceState? activeFence = null;
        var bracketStack = new Stack<char>();

        foreach (var line in lines)
        {
            var trimmed = line.TrimStart();

            if (activeFence is null)
            {
                if (TryParseFenceOpening(trimmed, out var opening))
                {
                    states.Add(new MarkdownLineState(null, opening));
                    activeFence = new FenceState(
                        opening.MarkerChar,
                        opening.MarkerLength,
                        opening.LanguageLabel,
                        ResolveEmbeddedProfile(opening.LanguageLabel),
                        []);
                    bracketStack.Clear();
                    continue;
                }

                states.Add(new MarkdownLineState(null, null));
                continue;
            }

            if (TryParseFenceClosing(trimmed, activeFence, out var closing))
            {
                states.Add(new MarkdownLineState(null, closing));
                activeFence = null;
                bracketStack.Clear();
                continue;
            }

            var lineFenceState = activeFence with { BracketStack = bracketStack.Reverse().ToArray() };
            states.Add(new MarkdownLineState(lineFenceState, null));
            UpdateEmbeddedBracketStack(line, activeFence.Profile, bracketStack);
        }

        return new MarkdownSnapshot(text, states);
    }

    private void UpdateEmbeddedBracketStack(string text, EmbeddedSyntaxProfile? profile, Stack<char> bracketStack)
    {
        if (string.IsNullOrEmpty(text))
            return;

        bool[]? protectedRanges = null;
        if (profile is not null)
            protectedRanges = BuildEmbeddedProtectedRanges(text, profile, 0, applyColors: false);

        for (var index = 0; index < text.Length; index++)
        {
            if (protectedRanges is not null && protectedRanges[index])
                continue;

            var ch = text[index];
            if (OpeningToClosing.ContainsKey(ch))
            {
                bracketStack.Push(ch);
                continue;
            }

            if (ClosingToOpening.TryGetValue(ch, out var opening) &&
                bracketStack.Count > 0 &&
                bracketStack.Peek() == opening)
            {
                bracketStack.Pop();
            }
        }
    }

    private void ColorizeInlineCode(Match match, int lineOffset)
    {
        var openingTicks = match.Groups[1];
        var content = match.Groups["content"];
        var closingTicks = match.Groups[1];
        var closingIndex = match.Index + match.Length - closingTicks.Length;

        ApplyBrush(lineOffset, openingTicks.Index, openingTicks.Index + openingTicks.Length, _stringBrush);
        ApplyBrush(lineOffset, closingIndex, closingIndex + closingTicks.Length, _stringBrush);

        if (!content.Success || string.IsNullOrWhiteSpace(content.Value))
        {
            ApplyBrush(lineOffset, match.Index, match.Index + match.Length, _stringBrush);
            return;
        }

        ApplyBrush(lineOffset, content.Index, content.Index + content.Length, _stringBrush);
    }

    private EmbeddedSyntaxProfile? ResolveEmbeddedProfile(string? languageLabel)
    {
        if (string.IsNullOrWhiteSpace(languageLabel) || _languageResolver is null)
            return null;

        var extension = _languageResolver(languageLabel);
        return ResolveEmbeddedProfile(extension);
    }

    private EmbeddedSyntaxProfile? ResolveEmbeddedProfile(LoadedExtension? extension)
    {
        if (extension is null || IsMarkdownExtension(extension))
            return null;

        var cacheKey = $"{extension.Id}|{extension.Version}";
        if (_embeddedProfileCache.TryGetValue(cacheKey, out var cached))
            return cached;

        var profile = BuildEmbeddedProfile(extension);
        _embeddedProfileCache[cacheKey] = profile;
        return profile;
    }

    private static EmbeddedSyntaxProfile BuildEmbeddedProfile(LoadedExtension extension)
    {
        var stringBrush = BrushFor(extension, "string", "#CE9178");
        var commentBrush = BrushFor(extension, "comment", "#6A9955");
        var tokenRules = new List<EmbeddedBrushRule>();
        var supportsCommonStringPrefixes =
            extension.StringDelimiters.Contains("\"") ||
            extension.StringDelimiters.Contains("'") ||
            extension.MultiLineStringDelimiters.Contains("\"\"\"") ||
            extension.MultiLineStringDelimiters.Contains("'''");

        if (extension.Keywords.Length > 0)
            tokenRules.Add(new EmbeddedBrushRule(BuildTokenRegex(extension.Keywords), BrushFor(extension, "keyword", "#569CD6")));
        if (extension.Types.Length > 0)
            tokenRules.Add(new EmbeddedBrushRule(BuildTokenRegex(extension.Types), BrushFor(extension, "type", "#4EC9B0")));
        if (extension.Namespaces.Length > 0)
            tokenRules.Add(new EmbeddedBrushRule(BuildTokenRegex(extension.Namespaces), BrushFor(extension, "namespace", "#4FC1FF")));
        if (extension.Properties.Length > 0)
            tokenRules.Add(new EmbeddedBrushRule(BuildTokenRegex(extension.Properties), BrushFor(extension, "property", "#9CDCFE")));
        if (extension.Functions.Length > 0)
            tokenRules.Add(new EmbeddedBrushRule(BuildTokenRegex(extension.Functions), BrushFor(extension, "function", "#DCDCAA")));

        if (supportsCommonStringPrefixes)
        {
            tokenRules.Add(new EmbeddedBrushRule(
                new Regex(CommonStringPrefixPattern, RegexOptions.Compiled),
                stringBrush));
        }

        if (extension.CommentBlockStart == "<!--")
        {
            tokenRules.Add(new EmbeddedBrushRule(
                new Regex(@"(?<=>)\s*v?\d[\d._-]*[A-Za-z0-9]*\s*(?=<)", RegexOptions.Compiled),
                BrushFor(extension, "number", "#B5CEA8")));
        }

        tokenRules.Add(new EmbeddedBrushRule(
            new Regex(@"(?<![\p{L}\p{Nd}_])(?:0[xX][0-9A-Fa-f]+|0[bB][01]+|0[oO][0-7]+|\d+(?:\.\d+)?(?:[eE][+\-]?\d+)?)(?![\p{L}\p{Nd}_])", RegexOptions.Compiled),
            BrushFor(extension, "number", "#B5CEA8")));

        if (extension.CommentBlockStart != "<!--")
        {
            tokenRules.Add(new EmbeddedBrushRule(
                new Regex(@"(?<=\.|->|::)[\p{L}_][\p{L}\p{Nd}_:-]*", RegexOptions.Compiled),
                BrushFor(extension, "property", "#9CDCFE")));
            tokenRules.Add(new EmbeddedBrushRule(
                new Regex(@"(?<![\p{L}\p{Nd}_])[\p{L}_][\p{L}\p{Nd}_]*(?=\.)", RegexOptions.Compiled),
                BrushFor(extension, "namespace", "#4FC1FF")));
        }

        tokenRules.Add(new EmbeddedBrushRule(
            new Regex(@"(?<![\p{L}\p{Nd}_])[@#][\p{L}_][\p{L}\p{Nd}_-]*", RegexOptions.Compiled),
            BrushFor(extension, "preprocessor", "#C586C0")));
        tokenRules.Add(new EmbeddedBrushRule(
            new Regex(@"(?<=\[)[\p{L}_][\p{L}\p{Nd}_:.]*(?=[,\]\(])|(?<=<)[\p{L}_][\p{L}\p{Nd}_:-]*(?=[^>]*>)", RegexOptions.Compiled),
            BrushFor(extension, "attribute", "#C586C0")));
        tokenRules.Add(new EmbeddedBrushRule(
            new Regex(@"(?<![\p{L}\p{Nd}_])[\p{L}_][\p{L}\p{Nd}_]*(?=\s*\()", RegexOptions.Compiled),
            BrushFor(extension, "function", "#DCDCAA")));
        tokenRules.Add(new EmbeddedBrushRule(
            new Regex(@"=>|->|::|\+\+|--|\+=|-=|\*=|/=|%=|&&|\|\||<<|>>|<=|>=|==|!=|=|\+|-|\*|/|%|!|\?|:|<|>|&|\||\^|~", RegexOptions.Compiled),
            BrushFor(extension, "operator", "#D4D4D4")));
        tokenRules.Add(new EmbeddedBrushRule(
            new Regex(extension.CommentBlockStart == "<!--" ? @"[{}\[\]();,]" : @"[{}\[\]();,.]", RegexOptions.Compiled),
            BrushFor(extension, "punctuation", "#D4D4D4")));
        tokenRules.Add(new EmbeddedBrushRule(
            new Regex(@"(?<=\b(?:using|import|include|require|use|from)\b\s+(?:[\p{L}_][\p{L}\p{Nd}_./\\]*\s*[./\\]\s*)?)[\p{L}_][\p{L}\p{Nd}_]*(?=\s*(?:;|$))", RegexOptions.Compiled),
            BrushFor(extension, "namespace", "#4FC1FF")));

        var variableTokens = extension.Keywords
            .Concat(extension.Types)
            .Concat(extension.Functions)
            .Concat(extension.Properties)
            .Concat(extension.Namespaces);
        tokenRules.Add(new EmbeddedBrushRule(
            BuildVariableRegex(variableTokens),
            BrushFor(extension, "variable", "#A0DBFD")));

        var stringRegexes = new List<Regex>();

        foreach (var delimiter in extension.MultiLineStringDelimiters.Where(d => !string.IsNullOrWhiteSpace(d)).Distinct())
        {
            var escaped = Regex.Escape(delimiter);
            stringRegexes.Add(new Regex($@"{escaped}.*?{escaped}", RegexOptions.Compiled));
        }

        foreach (var delimiter in extension.StringDelimiters.Where(d => !string.IsNullOrWhiteSpace(d)).Distinct())
        {
            if (delimiter == "'" && extension.DisableSingleQuoteStrings)
                continue;

            stringRegexes.Add(BuildSingleLineStringRegex(delimiter));
        }

        if (extension.DisableSingleQuoteStrings)
        {
            tokenRules.Add(new EmbeddedBrushRule(
                new Regex(@"'(?:\\(?:u[0-9A-Fa-f]{4}|U[0-9A-Fa-f]{8}|x[0-9A-Fa-f]{1,4}|[0-7]{1,3}|[abfnrtv\\""'0])|[^\\'])'", RegexOptions.Compiled),
                BrushFor(extension, "charLiteral", "#CE9178")));
        }

        Regex? singleLineCommentRegex = null;
        if (!string.IsNullOrWhiteSpace(extension.CommentLine))
            singleLineCommentRegex = new Regex(Regex.Escape(extension.CommentLine) + @".*$", RegexOptions.Compiled);

        return new EmbeddedSyntaxProfile(commentBrush, stringBrush, singleLineCommentRegex, stringRegexes, tokenRules);
    }

    private static IBrush GetRainbowBrush(int depth) => RainbowBrushes[Math.Abs(depth) % RainbowBrushes.Length];

    private static Regex BuildSingleLineStringRegex(string delimiter)
    {
        var escaped = Regex.Escape(delimiter);
        return delimiter switch
        {
            "\"" => new Regex("\"(?:\\\\.|[^\"\\\\])*\"", RegexOptions.Compiled),
            "'" => new Regex("'(?:\\\\.|[^'\\\\])*'", RegexOptions.Compiled),
            "`" => new Regex("`(?:\\\\.|[^`\\\\])*`", RegexOptions.Compiled),
            _ => new Regex($@"{escaped}.*?{escaped}", RegexOptions.Compiled)
        };
    }

    private void ApplyBrush(int lineOffset, int startOffset, int endOffset, IBrush brush)
    {
        if (endOffset <= startOffset)
            return;

        ChangeLinePart(lineOffset + startOffset, lineOffset + endOffset, element =>
        {
            var properties = element.TextRunProperties.Clone();
            properties.SetForegroundBrush(brush);
            SetTextRunPropertiesMethod?.Invoke(element, [properties]);
        });
    }

    private static bool TryParseFenceOpening(string trimmed, out FenceDelimiterInfo delimiter)
    {
        delimiter = default;
        if (string.IsNullOrWhiteSpace(trimmed))
            return false;

        var markerChar = trimmed[0];
        if (markerChar is not ('`' or '~'))
            return false;

        var markerLength = 0;
        while (markerLength < trimmed.Length && trimmed[markerLength] == markerChar)
            markerLength++;

        if (markerLength < 3)
            return false;

        var info = trimmed[markerLength..].Trim();
        delimiter = new FenceDelimiterInfo(markerChar, markerLength, info);
        return true;
    }

    private static bool TryParseFenceClosing(string trimmed, FenceState activeFence, out FenceDelimiterInfo delimiter)
    {
        delimiter = default;
        if (string.IsNullOrWhiteSpace(trimmed) || trimmed[0] != activeFence.MarkerChar)
            return false;

        var markerLength = 0;
        while (markerLength < trimmed.Length && trimmed[markerLength] == activeFence.MarkerChar)
            markerLength++;

        if (markerLength < activeFence.MarkerLength || !string.IsNullOrWhiteSpace(trimmed[markerLength..]))
            return false;

        delimiter = new FenceDelimiterInfo(activeFence.MarkerChar, markerLength, string.Empty);
        return true;
    }

    private static bool IsMarkdownExtension(LoadedExtension? extension) =>
        string.Equals(extension?.Id, "markdown-kodo-extension", StringComparison.OrdinalIgnoreCase);

    private static IBrush BrushFor(LoadedExtension extension, string tokenName, string fallback)
    {
        var hex = extension.ColorTokens.TryGetValue(tokenName, out var value) ? value : fallback;
        return Brush.Parse(hex);
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

    private static Regex BuildVariableRegex(IEnumerable<string> reservedTokens)
    {
        const string variableIdentifierBodyPattern = "[\\p{L}_][\\p{L}\\p{Nd}_]*";

        var reserved = reservedTokens
            .Where(token => !string.IsNullOrWhiteSpace(token))
            .Distinct(StringComparer.Ordinal)
            .OrderByDescending(token => token.Length)
            .Select(Regex.Escape)
            .ToArray();

        var reservedPrefix = reserved.Length > 0
            ? $"(?!{string.Join("|", reserved.Select(r => r + "(?![\\p{L}\\p{Nd}_])"))})"
            : string.Empty;

        return new Regex(
            $"(?<![.\\p{{L}}\\p{{Nd}}_]){reservedPrefix}{variableIdentifierBodyPattern}(?!\\s*[\\.(\"'`]|[\\p{{L}}\\p{{Nd}}_])",
            RegexOptions.Compiled);
    }

    private static bool TryReserveRange(bool[] protectedRanges, int start, int end)
    {
        if (IsProtected(protectedRanges, start, end))
            return false;

        MarkRange(protectedRanges, start, end);
        return true;
    }

    private static bool IsProtected(bool[] protectedRanges, int start, int end)
    {
        for (var index = start; index < end && index < protectedRanges.Length; index++)
        {
            if (protectedRanges[index])
                return true;
        }

        return false;
    }

    private static void MarkRange(bool[] protectedRanges, int start, int end)
    {
        for (var index = start; index < end && index < protectedRanges.Length; index++)
            protectedRanges[index] = true;
    }

    private static IEnumerable<int> AllIndexesOf(string text, char character)
    {
        for (var index = 0; index < text.Length; index++)
        {
            if (text[index] == character)
                yield return index;
        }
    }

    private sealed record MarkdownSnapshot(string Text, List<MarkdownLineState> LineStates)
    {
        public MarkdownLineState GetLineState(int lineNumber) =>
            lineNumber > 0 && lineNumber <= LineStates.Count
                ? LineStates[lineNumber - 1]
                : new MarkdownLineState(null, null);
    }

    private readonly record struct MarkdownLineState(FenceState? ActiveFence, FenceDelimiterInfo? Delimiter);
    private readonly record struct FenceDelimiterInfo(char MarkerChar, int MarkerLength, string? LanguageLabel);
    private sealed record FenceState(char MarkerChar, int MarkerLength, string? LanguageLabel, EmbeddedSyntaxProfile? Profile, IReadOnlyList<char> BracketStack);
    private sealed record EmbeddedSyntaxProfile(
        IBrush CommentBrush,
        IBrush StringBrush,
        Regex? SingleLineCommentRegex,
        List<Regex> StringRegexes,
        List<EmbeddedBrushRule> TokenRules);
    private readonly record struct EmbeddedBrushRule(Regex Regex, IBrush Brush);
}

// ── Syntax highlighting ──────────────────────────────────────────────────────
// Builds an AvaloniaEdit IHighlightingDefinition at runtime from the data
// declared in a LoadedExtension (keywords, types, comment markers, color tokens).
public sealed class KodoHighlightingDefinition : IHighlightingDefinition
{
    private const string VariableIdentifierBodyPattern =
        "[\\p{L}_][\\p{L}\\p{Nd}_]*";
    private const string CommonStringPrefixPattern =
        "(?i)(?<![\\p{L}\\p{Nd}_])(?:fr|rf|br|rb|ur|ru|cr|rc|f|r|u|b|c)(?=(?:\\\"\\\"\\\"|'''|\\\"|'|#+\\\"))";

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
        var supportsCommonStringPrefixes =
            ext.StringDelimiters.Contains("\"") ||
            ext.StringDelimiters.Contains("'") ||
            ext.MultiLineStringDelimiters.Contains("\"\"\"") ||
            ext.MultiLineStringDelimiters.Contains("'''");

        var isMarkdown = string.Equals(ext.Id, "markdown-kodo-extension", StringComparison.OrdinalIgnoreCase);

        // ── Inner rulesets ────────────────────────────────────────────────────────────
        // codeRuleSet  - holds keyword/type/number rules; used as the inner ruleset of
        //                spans that should still syntax-colour their contents.
        //                (Currently unused as inner ruleset, but kept for clarity.)
        // emptyRuleSet - no rules; used as the inner ruleset of comment and string spans
        //                so that keyword/number rules cannot fire inside them.
        //                AvaloniaEdit applies SpanColor to the entire span body, so the
        //                correct span colour is provided by the outer SpanColor property.
        //                The inner ruleset only needs to be empty - no DefaultColor needed.
        var codeRuleSet  = new HighlightingRuleSet();
        var emptyRuleSet = new HighlightingRuleSet();

        if (!isMarkdown && ext.Keywords.Length > 0)
        {
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = BuildTokenRegex(ext.Keywords),
                Color = keywordColor
            });
        }

        if (!isMarkdown && ext.Types.Length > 0)
        {
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = BuildTokenRegex(ext.Types),
                Color = typeColor
            });
        }

        if (!isMarkdown && ext.Namespaces.Length > 0)
        {
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = BuildTokenRegex(ext.Namespaces),
                Color = namespaceColor
            });
        }

        if (!isMarkdown && ext.Properties.Length > 0)
        {
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = BuildTokenRegex(ext.Properties),
                Color = propertyColor
            });
        }

        if (!isMarkdown && ext.Functions.Length > 0)
        {
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = BuildTokenRegex(ext.Functions),
                Color = functionColor
            });
        }

        // XML/MSBuild version string rule - only active for XML-family languages
        // (identified by <!-- block comment syntax). Must be inserted before the number
        // rule so the full version token (e.g. v1.0.0-DEV, net10.0) is claimed as a
        // single unit rather than being fragmented across number/operator/variable rules.
        if (ext.CommentBlockStart == "<!--")
        {
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(@"(?<=>)\s*v?\d[\d._-]*[A-Za-z0-9]*\s*(?=<)", RegexOptions.Compiled),
                Color = numberColor
            });
        }

        if (!isMarkdown)
        {
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(@"(?<![\p{L}\p{Nd}_])(?:0[xX][0-9A-Fa-f]+|0[bB][01]+|0[oO][0-7]+|\d+(?:\.\d+)?(?:[eE][+\-]?\d+)?)(?![\p{L}\p{Nd}_])", RegexOptions.Compiled),
                Color = numberColor
            });
        }

        // Property and namespace rules fire on dot-separated identifiers (e.g. foo.Bar).
        // In XML-family languages dots appear in file paths and version strings between
        // element tags, so these rules would wrongly colour path segments as namespace/
        // property tokens. Skip them entirely for XML - no chained member access exists.
        // Also skip for Markdown where dots in URLs and file extensions cause false matches.
        if (ext.CommentBlockStart != "<!--" && !isMarkdown)
        {
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
        }

        if (!isMarkdown)
        {
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
        }

        if (!isMarkdown)
        {
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

            // For XML-family languages the dot appears in file paths and version strings
            // between element tags, so exclude it from punctuation to avoid grey fragments.
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(ext.CommentBlockStart == "<!--" ? @"[{}\[\]();,]" : @"[{}\[\]();,.]", RegexOptions.Compiled),
                Color = punctuationColor
            });

            // Last segment of an import/using directive for any language.
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(
                    @"(?<=\b(?:using|import|include|require|use|from)\b\s+(?:[\p{L}_][\p{L}\p{Nd}_./\\]*\s*[./\\]\s*)?)[\p{L}_][\p{L}\p{Nd}_]*(?=\s*(?:;|$))",
                    RegexOptions.Compiled),
                Color = namespaceColor
            });

            // User-defined variables - bare identifiers not preceded by dot, not followed
            // by '(' or '.', and not already claimed by keyword/type/number rules.
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = BuildVariableRegex(ext.Keywords.Concat(ext.Types).Concat(ext.Functions).Concat(ext.Properties).Concat(ext.Namespaces)),
                Color = variableColor
            });

            // Char literal rule (C#-style languages with disableSingleQuoteStrings).
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
        }
        else
        {
            // ── Markdown-specific inline rules ─────────────────────────────────────────
            // Bold/bold-italic markers: ** *** __ ___
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(@"\*{2,3}|_{2,3}", RegexOptions.Compiled),
                Color = operatorColor
            });
            // Italic markers: single * or _ not adjacent to another
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(@"(?<!\*)\*(?!\*)|(?<!_)_(?!_)", RegexOptions.Compiled),
                Color = operatorColor
            });
            // Strikethrough markers: ~~
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(@"~~", RegexOptions.Compiled),
                Color = operatorColor
            });
            // Link/image bracket and paren delimiters
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(@"!?\[|\]\(|\)", RegexOptions.Compiled),
                Color = punctuationColor
            });
            // Table pipe separators
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(@"\|", RegexOptions.Compiled),
                Color = punctuationColor
            });
            // Unordered list markers at start of line: - + *
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(@"(?<=^|\n)[ \t]*[-+*](?=[ \t])", RegexOptions.Compiled),
                Color = operatorColor
            });
        }

        // Char literal rule for non-Markdown languages with disableSingleQuoteStrings
        // (already handled inside the !isMarkdown block above; this guard is now a no-op
        // for Markdown but kept outside to satisfy the original structure expectation).

        // ── Main ruleset ──────────────────────────────────────────────────────────────
        // Spans are checked in order - first match wins. SpanColor is applied by
        // AvaloniaEdit to the entire span (start delimiter + body + end delimiter),
        // modulated by SpanColorIncludesStart / SpanColorIncludesEnd for the delimiters.
        // The inner emptyRuleSet ensures no keyword/number rules fire inside the span.
        var mainRuleSet = new HighlightingRuleSet();

        // Block comment /* … */ - added first so it takes priority over // on the same line.
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

        // ── Markdown heading spans ────────────────────────────────────────────────────
        // Added before the generic commentLine span so that #-headings are coloured
        // with keywordColor rather than commentColor (which handles blockquotes via >).
        if (isMarkdown)
        {
            mainRuleSet.Spans.Add(new HighlightingSpan
            {
                StartExpression        = new Regex(@"^#{1,6}(?=\s)", RegexOptions.Compiled | RegexOptions.Multiline),
                EndExpression          = new Regex(@"$", RegexOptions.Compiled),
                SpanColor              = keywordColor,
                SpanColorIncludesStart = true,
                SpanColorIncludesEnd   = false,
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

        if (supportsCommonStringPrefixes)
        {
            if (ext.MultiLineStringDelimiters.Contains("\"\"\""))
                mainRuleSet.Spans.Add(CreateRegexStringSpan("(?i)(?:fr|rf|br|rb|ur|ru|cr|rc|f|r|u|b|c)\\\"\\\"\\\"", "\\\"\\\"\\\"", stringColor, emptyRuleSet, allowEndOfLineFallback: false));

            if (ext.MultiLineStringDelimiters.Contains("'''"))
                mainRuleSet.Spans.Add(CreateRegexStringSpan(@"(?i)(?:fr|rf|br|rb|ur|ru|cr|rc|f|r|u|b|c)'''", @"'''", stringColor, emptyRuleSet, allowEndOfLineFallback: false));

            if (ext.StringDelimiters.Contains("\""))
                mainRuleSet.Spans.Add(CreateRegexStringSpan("(?i)(?:fr|rf|br|rb|ur|ru|cr|rc|f|r|u|b|c)\\\"", "\\\"", stringColor, emptyRuleSet, allowEndOfLineFallback: true));

            if (ext.StringDelimiters.Contains("'"))
                mainRuleSet.Spans.Add(CreateRegexStringSpan(@"(?i)(?:fr|rf|br|rb|ur|ru|cr|rc|f|r|u|b|c)'", @"'", stringColor, emptyRuleSet, allowEndOfLineFallback: true));
        }

        if (ext.DisableSingleQuoteStrings && ext.StringDelimiters.Contains("\""))
        {
            mainRuleSet.Spans.Add(CreateRegexStringSpan(@"(?:\$@|@\$)""", @"""(?!"")", stringColor, emptyRuleSet, allowEndOfLineFallback: false));
            mainRuleSet.Spans.Add(CreateRegexStringSpan(@"\$""", @"""", stringColor, emptyRuleSet, allowEndOfLineFallback: true));
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

    private static Regex BuildVariableRegex(IEnumerable<string> reservedTokens)
    {
        var reserved = reservedTokens
            .Where(token => !string.IsNullOrWhiteSpace(token))
            .Distinct(StringComparer.Ordinal)
            .OrderByDescending(token => token.Length)
            .Select(Regex.Escape)
            .ToArray();

        var reservedPrefix = reserved.Length > 0
            ? $"(?!{string.Join("|", reserved.Select(r => r + "(?![\\p{L}\\p{Nd}_])"))})"
            : string.Empty;

        return new Regex(
            $"(?<![.\\p{{L}}\\p{{Nd}}_]){reservedPrefix}{VariableIdentifierBodyPattern}(?!\\s*[\\.(\"'`]|[\\p{{L}}\\p{{Nd}}_])",
            RegexOptions.Compiled);
    }

    private static HighlightingSpan CreateStringSpan(
        string delimiter,
        HighlightingColor stringColor,
        HighlightingRuleSet emptyRuleSet,
        bool allowEndOfLineFallback)
    {
        var escapedDelimiter = Regex.Escape(delimiter);
        return CreateRegexStringSpan(escapedDelimiter, escapedDelimiter, stringColor, emptyRuleSet, allowEndOfLineFallback);
    }

    private static HighlightingSpan CreateRegexStringSpan(
        string startPattern,
        string endDelimiterPattern,
        HighlightingColor stringColor,
        HighlightingRuleSet emptyRuleSet,
        bool allowEndOfLineFallback)
    {
        var endPattern = allowEndOfLineFallback
            ? $@"(?<!\\){endDelimiterPattern}|$"
            : $@"(?<!\\){endDelimiterPattern}";

        return new HighlightingSpan
        {
            StartExpression = new Regex(startPattern, RegexOptions.Compiled),
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
    // name will ever be looked up - return an empty set for all names.
    public HighlightingRuleSet GetNamedRuleSet(string name) => new HighlightingRuleSet();
}

public sealed class IndentGuideBackgroundRenderer : IBackgroundRenderer
{
    public KnownLayer Layer => KnownLayer.Background;

    public int TabSize { get; set; } = 4;

    // Disabled for unsaved (untitled) files and plain-text files (.txt / .log / .text)
    // so that smart visual features don't activate where no language context exists.
    public bool IsEnabled { get; set; } = true;

    public IBrush GuideBrush { get; set; } = new SolidColorBrush(Color.Parse("#808080"), 0.4);

    public void Draw(TextView textView, DrawingContext drawingContext)
    {
        if (!IsEnabled)
            return;

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