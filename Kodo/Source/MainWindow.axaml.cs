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
using System.Net;
using System.Net.NetworkInformation;
using System.Runtime.CompilerServices;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Text;
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
    public string ChevronText => IsDirectory ? (_isExpanded ? "↓" : "→") : string.Empty;

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
    // Raw PNG or SVG bytes read from icon.png / icon.svg on the background scan thread.
    // Decoded into IconImage (PNG) or SvgData (SVG) on the UI thread by ApplyLoadedExtensionsResult.
    public byte[]? IconBytes { get; set; }
    // Optional icon loaded from icon.png / icon.svg inside the .kox / folder
    public Bitmap? IconImage { get; set; }
    // SVG text for icons sourced from the marketplace index or local icon.svg
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

/// <summary>
/// A named group of one or more theme cards from a single extension.
/// Groups with more than one theme are shown as a collapsible section
/// under Installed Theme Extensions.
/// </summary>
public class ThemeExtensionGroup : INotifyPropertyChanged
{
    private bool _isExpanded;

    public string GroupName { get; }
    public IReadOnlyList<LoadedExtension> Themes { get; }

    /// <summary>True when this extension packs more than one theme.</summary>
    public bool IsMultiTheme => Themes.Count > 1;

    /// <summary>
    /// Controls whether the card row is visible.
    /// Single-theme groups are always expanded (no collapse chrome shown).
    /// </summary>
    public bool IsExpanded
    {
        get => _isExpanded;
        set
        {
            if (_isExpanded == value) return;
            _isExpanded = value;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsExpanded)));
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ChevronGlyph)));
        }
    }

    /// <summary>▶ when collapsed, ▼ when expanded.</summary>
    public string ChevronGlyph => _isExpanded ? "▾" : "▸";

    public ThemeExtensionGroup(string groupName, IReadOnlyList<LoadedExtension> themes)
    {
        GroupName = groupName;
        Themes    = themes;
        // Multi-theme groups start collapsed; single-theme groups are always open.
        _isExpanded = !IsMultiTheme;
    }

    public event PropertyChangedEventHandler? PropertyChanged;
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

// One inline run (bold or normal) within a release-notes paragraph.
public sealed class FormattedRun
{
    public string Text   { get; init; } = string.Empty;
    public bool   IsBold { get; init; }
}

// One paragraph (or bullet/ordered-list item) in the release notes.
// Runs are rendered inline (WrapPanel) so bold/normal text flows together.
// The list marker (bullet glyph or "1.") is kept separate from the runs and
// rendered in its own fixed-width column so wrapped lines hang-indent and
// align under the first word instead of sliding back under the marker.
public sealed class FormattedParagraph
{
    public IReadOnlyList<FormattedRun> Runs      { get; init; } = [];
    // Extra top margin so paragraphs breathe; bullet items get slightly less.
    public Thickness TopMargin { get; init; } = new Thickness(0, 6, 0, 0);
    // "•" for bullets, "1." / "2." / ... for ordered items, or empty for a
    // plain paragraph/heading (in which case MarkerColumnWidth is 0).
    public string Marker { get; init; } = string.Empty;
    // Fixed width of the marker column, shared across rows so every bullet's
    // wrapped text lines up under its own first line instead of under "• ".
    public double MarkerColumnWidth { get; init; }
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
    private const int MaxRecentFiles = 6;
    private const string DefaultDiscordClientId = "1495509170756255744";
    private const string DefaultDiscordLargeImageKey = "kodo_logo";
    private const string DefaultDiscordLargeImageText = "Kodo";
    private const string SettingsFileName = "kodosettings.json";
    // Bounds for the drag-resizable terminal panel. Kept in sync with the
    // RowDefinition's MaxHeight in MainWindow.axaml - if one changes, change both.
    private const double MinTerminalPanelHeight = 120;
    private const double MaxTerminalPanelHeight = 420;
    private const double DefaultTerminalPanelHeight = 300;
    private const string DiscordClientIdEnvironmentVariable = "KODO_DISCORD_CLIENT_ID";
    private const string AutoSaveSavedMessage = "Saved.";
    private const string AutoSaveSavingMessage = "Saving...";
    private const string AutoSaveFailedMessagePrefix = "Save failed:";
    // Read from <InformationalVersion> in Kodo.csproj (e.g. "v1.3.0-BETA").
    // To update the app version, change only that tag in the csproj.
    // Resolved via KodoDiagnostics.AppVersion (strips the +<git-hash> suffix).
    private static readonly string CurrentAppVersion = KodoDiagnostics.AppVersion;
    public string CopyrightText => $"© {DateTime.Now.Year} KerbalMissile and SS-YYC. Licensed under KPL-v1.1.";
    // GitHub Contents API endpoint for the extension index JSON.
    // All Kodo-hosted assets (index JSON, extension .kox packages, and repo-hosted
    // icon images) are fetched via the Contents API with the raw+json Accept header,
    // which makes GitHub return file bytes directly - no base64 unwrapping needed.
    // Third-party icon URLs (e.g. Wikipedia, Wikimedia) are fetched as plain GETs
    // since the Contents API only applies to files inside this repository.
    private const string DefaultMarketplaceIndexUrl = "https://api.github.com/repos/KerbalMissile/Kodo/contents/Indexs/ExtensionsIndex.json";
    private const string LatestReleaseApiUrl = "https://api.github.com/repos/KerbalMissile/Kodo/releases/latest";
    private const string ReleasesApiUrl = "https://api.github.com/repos/KerbalMissile/Kodo/releases";
    private const string ReleasesPageUrl = "https://github.com/KerbalMissile/Kodo/releases";
    private const string DiscordServerUrl = "https://discord.gg/cUQ6C88Z9C";
    private const string WebsiteUrl = "https://kerbalmissile.github.io/Kodo-Website/";
    // GitHub Contents API endpoint for ANNOUNCEMENTS.md.  Uses the same raw+json
    // Accept header as the marketplace index so GitHub returns the file bytes directly
    // with no base64 unwrapping, and benefits from the same generous rate limits.
    private const string AnnouncementsUrl = "https://api.github.com/repos/KerbalMissile/Kodo/contents/Announcements/ANNOUNCEMENTS.md";

    private bool _isNewsLoading = true;
    private bool _isNewsError;

    private string? _currentFilePath;
    // Encoding detected (or chosen) for the currently open file. Defaults to UTF-8.
    private System.Text.Encoding _currentFileEncoding = System.Text.Encoding.UTF8;
    private string? _currentFolderPath;
    private DiscordRpcClient? _discordRpcClient;
    private readonly DispatcherTimer _autoSaveTimer = new() { Interval = TimeSpan.FromMilliseconds(300) };
    private readonly DispatcherTimer _autoSaveStatusTimer = new() { Interval = TimeSpan.FromSeconds(3) };
    private readonly DispatcherTimer _discordReconnectTimer = new() { Interval = TimeSpan.FromSeconds(10) };
    private readonly DispatcherTimer _editorStateRefreshTimer = new() { Interval = TimeSpan.FromMilliseconds(75) };
    private readonly DispatcherTimer _wordCountRefreshTimer = new() { Interval = TimeSpan.FromMilliseconds(175) };
    private readonly DispatcherTimer _settingsSaveDebounceTimer = new() { Interval = TimeSpan.FromMilliseconds(400) };
    // Polls the Windows accent registry key so the blob and active accent stay
    // live without requiring the Microsoft.Win32.SystemEvents NuGet package.
    private readonly DispatcherTimer _windowsAccentPollTimer = new() { Interval = TimeSpan.FromSeconds(2) };
    private string _lastSeenWindowsAccentHex = string.Empty;
    // Polls the Windows "AppsUseLightTheme" registry value (same cadence/approach
    // as the accent poll above) so the System Default theme blob's preview swatch
    // - and the active palette, when System mode is selected - track a live
    // Windows light/dark toggle without requiring an app restart.
    private readonly DispatcherTimer _windowsThemePollTimer = new() { Interval = TimeSpan.FromSeconds(2) };
    private string _lastSeenWindowsThemeName = string.Empty;
    private readonly RainbowBracketColorizer _rainbowBracketColorizer = new();
    private readonly InterpolatedStringColorizer _interpolatedStringColorizer = new();
    private readonly HtmlEmbeddedColorizer _htmlEmbeddedColorizer = new();
    private readonly MarkdownColorizer _markdownColorizer = new();
    private readonly EmojiTypefaceColorizer _emojiTypefaceColorizer = new();
    private EditorTab? _activeEditorTab;
    private int _nextUntitledTabNumber = 1;
    private string? _autoSaveStatusMessage;
    private bool _isAutoSaveEnabled;
    private bool _isDirty;
    private bool _isSaving;
    private bool _isDiscordRichPresenceEnabled;
    private bool _isDiscordImprovedRpcEnabled;
    private bool _hasUntitledDocument;
    private bool _isRefreshingExtensions;
    private bool _isUpdatingAllExtensions;
    // Guards the silent background sweep (AutoUpdateExtensionsIfEnabledAsync)
    // so it never overlaps with itself or with the manual "Update All" button.
    private bool _isAutoUpdatingExtensions;
    private bool _isAutoUpdateExtensionsEnabled;
    // Sub-setting under IsAutoUpdateExtensionsEnabled: when on, the silent
    // extension-update sweep doesn't touch ExtensionsStatusText while it
    // works, so nothing visibly changes on the Extensions page either.
    private bool _isAutoUpdateExtensionsInBackgroundEnabled;
    // Mirrors _isAutoUpdateExtensionsEnabled, but for whole-app updates
    // (downloading and installing newer Kodo releases from GitHub) rather
    // than marketplace extensions. Kept as a separate flag since a user may
    // want one without the other.
    private bool _isAutoUpdateAppEnabled = true;
    // Sub-setting under _isAutoUpdateAppEnabled: when on, a found update is
    // downloaded and installed straight away (UpdateService.SilentlyInstallAsync)
    // instead of showing UpdateDialog's "Update Now" / "Later" prompt.
    private bool _isAutoUpdateAppInBackgroundEnabled;
    // Backs the manual "Check for Updates" button in Settings. Separate from
    // the silent startup check (App.axaml.cs) and from the auto-update toggle
    // above - clicking the button should always give feedback regardless of
    // whether automatic checking is on.
    private bool _isCheckingForUpdatesManually;
    private string _checkForUpdatesStatusText = string.Empty;
    private string _developerOptionsStatusText = string.Empty;
    private bool _isRefreshingLatestRelease;
    private bool _isSettingsPageVisible;
    private bool _isExtensionsPageVisible;
    private bool _isTutorialPageVisible;
    private bool _tutorialOpenedFromSettings;
    private bool _isWhatsNewPageVisible;
    private bool _isWhatsNewExpanded;
    private bool _isUpdateSplashVisible;
    // The version string that was running the last time the user launched Kodo.
    // When it differs from CurrentAppVersion we show the update splash once.
    private string _lastSeenVersion = string.Empty;
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
    private bool _isDeveloperOptionsVisible;
    private bool _isVerboseLoggingEnabled;
    private int _tabSize = 4;
    private int _editorFontSize = 14;
    private string _accentColorMode = "kodo";   // "kodo" | "windows" | "custom"
    private string _customAccentHex = "#8C00FF";
    // The accent colour supplied by the active theme; restored when switching back to "kodo" mode.
    private string _themeAccentHex = "#8C00FF";
    private bool   _hasThemeAccent  = false;
    private string _currentThemeName = "Dark";
    private string _requestedThemeName = "Dark";
    private string _editorStatsText = "0 lines";
    private string _wordCountText = string.Empty;
    private bool _pendingFullStateRefresh = true;
    private string _lastDiscordPresenceDetails = string.Empty;
    private string _lastDiscordPresenceState = string.Empty;
    private (string?, string?, int, bool, string?, bool, bool, bool, bool) _lastDiscordPresenceKey;
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
    private double _imageZoomLevel = 1.0;
    private const double ImageZoomMin = 0.1;
    private const double ImageZoomMax = 10.0;
    private const double ImageZoomStep = 0.25;
    private FileSystemWatcher? _extensionsFolderWatcher;
    private FileSystemWatcher? _projectExtensionsFolderWatcher;
    private readonly DispatcherTimer _extensionsRefreshDebounceTimer = new() { Interval = TimeSpan.FromMilliseconds(250) };
    // Periodic background check for extension updates while Kodo stays open.
    // Only runs (see UpdateExtensionAutoUpdateLifecycle) when the user has
    // opted into "Automatically update extensions" in Settings.
    private readonly DispatcherTimer _extensionAutoUpdateTimer = new() { Interval = TimeSpan.FromHours(6) };
    // Periodic background check for new Kodo releases while the app stays
    // open. Encapsulated in AppUpdateScheduler (Updater.cs), which owns its
    // own DispatcherTimer and start/stop lifecycle gated by
    // IsAutoUpdateAppEnabled. The launch-time check itself still lives in
    // App.axaml.cs (CheckForUpdatesInBackground) - this covers everything
    // after that. Constructed in the constructor body (needs settings-backed
    // properties to already exist), not here as a field initializer.
    private readonly AppUpdateScheduler _appUpdateScheduler;
    // Refreshes the marketplace listing (not extension auto-install) once an
    // hour while Kodo stays open. Always runs - unlike _extensionAutoUpdateTimer
    // this isn't gated by IsAutoUpdateExtensionsEnabled, since it only refreshes
    // what's shown in the Marketplace tab rather than installing anything.
    private readonly DispatcherTimer _marketplaceRefreshTimer = new() { Interval = TimeSpan.FromHours(1) };
    private readonly IndentGuideBackgroundRenderer _indentGuideRenderer = new();
    private readonly List<string> _startupOpenTabPaths = [];
    private readonly Dictionary<string, IBrush> _brushCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, byte[]> _marketplaceIconBytesCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, DateTime> _warningDialogCooldowns = new(StringComparer.OrdinalIgnoreCase);
    private readonly SemaphoreSlim _iconFetchSemaphore = new(4, 4);
    private static readonly TimeSpan ExtensionsRefreshCooldown = TimeSpan.FromSeconds(8);
    // Disk cache for the marketplace index JSON.  Sits next to settings in
    // %LocalAppData%\Kodo so it survives across sessions and network outages.
    // The sidecar .etag file holds the ETag returned by the last successful
    // GitHub Contents API response; sent as If-None-Match on subsequent requests
    // so a 304 Not Modified short-circuits the fetch without consuming a rate-limit
    // slot (304s are free against GitHub's 60 req/hr anonymous quota).
    private string MarketplaceIndexCachePath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Kodo", "marketplace-index.json");
    private string MarketplaceIndexETagPath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Kodo", "marketplace-index.json.etag");
    // In-memory ETag kept in sync with every successful 200 response.
    // Null means no cache exists yet; loaded lazily from disk on first fetch.
    private string? _marketplaceIndexETag;
    // Prevents the exact same error (same context + exception type + message) from
    // spawning duplicate dialogs within a short burst.  3 s is enough to debounce
    // rapid retries while still letting a genuinely new occurrence through quickly.
    private static readonly TimeSpan WarningDialogCooldown   = TimeSpan.FromSeconds(3);

    // Hard ceiling for any GitHub network operation (index fetch, announcements,
    // release info, icon downloads).  If the operation does not *complete* - both
    // headers and body - within this window we cancel it, log a warning, and show
    // the standard Kodo error dialog so the user knows the panel is stale.
    private static readonly TimeSpan GitHubOperationTimeout  = TimeSpan.FromSeconds(7);
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
    private bool _isPointerOverEditorLink;
    private readonly HashSet<EditorTab> _corruptedTabs = new(ReferenceEqualityComparer.Instance);
    private TerminalSession? _activeTerminalSession;
    // Holds the currently subscribed SessionExited handler so it can be
    // explicitly unsubscribed before a new one is attached on every Start().
    // Without this, switching sessions accumulates stale handlers that fire
    // on the wrong session when any future process exits.
    private EventHandler<IntPtr>? _activeSessionExitedHandler;
    private TerminalShellOption? _selectedTerminalShell;
    private int _nextTerminalNumber = 1;
    private bool _isTerminalVisible;
    private bool _isTerminalSupported = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
    private double _terminalPanelHeight = DefaultTerminalPanelHeight;
    private bool _isResizingTerminalPanel;
    private double _terminalPanelDragStartPointerY;
    private double _terminalPanelDragStartHeight;

    // Caches compiled KodoHighlightingDefinition instances by LoadedExtension identity.
    // Building one involves compiling multiple Regex objects (RegexOptions.Compiled), which
    // is expensive enough to cause a noticeable delay on every tab switch. The cache lives
    // for the session and is cleared whenever extensions are reloaded so it never goes stale.
    private readonly Dictionary<LoadedExtension, KodoHighlightingDefinition> _highlightingCache =
        new(ReferenceEqualityComparer.Instance);
    private readonly Dictionary<LoadedExtension, CompiledSyntaxProfile> _compiledSyntaxProfileCache =
        new(ReferenceEqualityComparer.Instance);
    // Caches the result of the content-sniff (first-line language detection) per absolute
    // file path. Avoids re-opening and reading the file on every tab switch for files that
    // have no matching extension (e.g. Makefile, Dockerfile, .csproj opened from a folder).
    // Entries are never stale within a session because the language mapping is fixed once
    // the file is first opened; the cache is small (one string key per opened file).
    private readonly Dictionary<string, LoadedExtension?> _contentSniffCache =
        new(StringComparer.OrdinalIgnoreCase);
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
            "Use the keyboard shortcut chips on Home for quick access to the most common actions."
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
        ["markdown"] = "md",
        // Explicit plain-text markers - map to empty string so no extension is matched
        ["text"] = "",
        ["plain"] = "",
        ["txt"] = "",
        ["plaintext"] = "",
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
    public ObservableCollection<NewsItem> NewsItems { get; } = new();

    public bool IsNewsLoading
    {
        get => _isNewsLoading;
        private set { if (_isNewsLoading == value) return; _isNewsLoading = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsNewsContentVisible)); OnPropertyChanged(nameof(IsNewsEmpty)); }
    }

    public bool IsNewsError
    {
        get => _isNewsError;
        private set { if (_isNewsError == value) return; _isNewsError = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsNewsContentVisible)); OnPropertyChanged(nameof(IsNewsEmpty)); }
    }

    public bool IsNewsContentVisible => !IsNewsLoading && !IsNewsError && NewsItems.Count > 0;
    public bool IsNewsEmpty => !IsNewsLoading && !IsNewsError && NewsItems.Count == 0;

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
            OnPropertyChanged(nameof(ImageZoomedWidth));
            OnPropertyChanged(nameof(ImageZoomedHeight));
            OnPropertyChanged(nameof(ImageZoomPercent));
        }
    }

    public double ImageZoomLevel
    {
        get => _imageZoomLevel;
        private set
        {
            var clamped = Math.Clamp(value, ImageZoomMin, ImageZoomMax);
            if (Math.Abs(_imageZoomLevel - clamped) < 0.001) return;
            _imageZoomLevel = clamped;
            OnPropertyChanged();
            OnPropertyChanged(nameof(ImageZoomPercent));
            OnPropertyChanged(nameof(ImageZoomedWidth));
            OnPropertyChanged(nameof(ImageZoomedHeight));
        }
    }

    public string ImageZoomPercent => $"{(int)Math.Round(_imageZoomLevel * 100)}%";

    public double ImageZoomedWidth =>
        CurrentImagePreview is not null ? CurrentImagePreview.PixelSize.Width * _imageZoomLevel : 0;

    public double ImageZoomedHeight =>
        CurrentImagePreview is not null ? CurrentImagePreview.PixelSize.Height * _imageZoomLevel : 0;

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
        // Suppress all SaveSettings() calls for the entire constructor + OnOpened
        // startup sequence.  The flag is cleared (and a clean write is forced) in
        // OnOpened's finally block, so nothing is lost.  Setting it here — rather
        // than only at the top of OnOpened — closes the window between the
        // constructor start and the Opened event where incidental saves from
        // CollectionChanged, ActiveEditorTab, or any future constructor-path call
        // could overwrite the just-loaded settings with a partial snapshot.
        _suppressSettingsSave = true;

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
        EditorTextBox.TextArea.TextView.LineTransformers.Add(_htmlEmbeddedColorizer);
        EditorTextBox.TextArea.TextView.LineTransformers.Add(_markdownColorizer);
        EditorTextBox.TextArea.TextView.LineTransformers.Add(_emojiTypefaceColorizer);
        EditorTextBox.TextArea.TextView.LinkTextForegroundBrush = Brush.Parse("#5BA3D9");
        EditorTextBox.TextArea.TextView.LinkTextBackgroundBrush = Brushes.Transparent;
        // Replace AvaloniaEdit's default LinkElementGenerator with a stricter one that
        // trims trailing punctuation (e.g. ')' ')' ']') so prose text ending with those
        // characters isn't swept into a link span.
        var defaultLinkGen = EditorTextBox.TextArea.TextView.ElementGenerators.OfType<LinkElementGenerator>().FirstOrDefault();
        if (defaultLinkGen is not null)
            EditorTextBox.TextArea.TextView.ElementGenerators.Remove(defaultLinkGen);
        EditorTextBox.TextArea.TextView.ElementGenerators.Add(new StrictLinkElementGenerator());
        // Show a "Ctrl+click to open" tooltip whenever the pointer hovers over a URL in the editor.
        // We hit-test the document position under the pointer and scan the line text with the same
        // regex used by StrictLinkElementGenerator, so the tooltip only appears over actual links.
        EditorTextBox.TextArea.TextView.PointerMoved += EditorTextView_OnPointerMoved;
        EditorTextBox.TextArea.TextView.PointerExited += EditorTextView_OnPointerExited;
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
        // Register Ctrl+wheel zoom on the image scroll viewer in the tunnel phase so we
        // intercept before ScrollViewer processes the event. Without tunnel registration
        // the ScrollViewer consumes the wheel event for scrolling before our handler runs,
        // meaning Ctrl+wheel appears to stop working once the image overflows its container.
        ImageScrollViewer.AddHandler(
            InputElement.PointerWheelChangedEvent,
            ImageScrollViewer_OnPointerWheelChanged,
            RoutingStrategies.Tunnel);
        _isFirstLaunch = !File.Exists(SettingsFilePath);
        var settings = LoadSettings();
        _requestedThemeName = string.IsNullOrWhiteSpace(settings.ThemeName) ? "Dark" : settings.ThemeName;
        _isAutoSaveEnabled = settings.AutoSaveEnabled;
        _isDiscordRichPresenceEnabled = settings.DiscordRichPresenceEnabled;
        _isDiscordImprovedRpcEnabled  = settings.DiscordImprovedRpcEnabled;
        _isDeveloperOptionsVisible = settings.DeveloperOptionsVisible;
        _isVerboseLoggingEnabled = settings.VerboseLoggingEnabled;
        KodoDiagnostics.VerboseLoggingEnabled = _isVerboseLoggingEnabled;
        _isStatusBarFilePathVisible = settings.StatusBarFilePathVisible;
        _isWordWrapEnabled = settings.WordWrapEnabled;
        _isConfirmBeforeClosingUnsavedTabsEnabled = settings.ConfirmBeforeClosingUnsavedTabsEnabled;
        _isRestoreOpenTabsOnLaunchEnabled = settings.RestoreOpenTabsOnLaunchEnabled;
        _isAutoUpdateExtensionsEnabled = settings.AutoUpdateExtensionsEnabled;
        _isAutoUpdateExtensionsInBackgroundEnabled = settings.AutoUpdateExtensionsInBackgroundEnabled;
        _isAutoUpdateAppEnabled = settings.AutoUpdateAppEnabled;
        _isAutoUpdateAppInBackgroundEnabled = settings.AutoUpdateAppInBackgroundEnabled;
        _hasCompletedTutorial = settings.HasCompletedTutorial;
        _accentColorMode = settings.AccentColorMode is "kodo" or "windows" or "custom" or "theme"
            ? settings.AccentColorMode : "kodo";
        // Migration: before "theme" was a distinct mode, "kodo" with a theme active
        // meant "follow the theme accent". Detect this case and upgrade the saved value
        // so existing users don't silently lose their theme accent preference.
        // (The theme hasn't loaded yet here, so we defer the check to after ApplyThemeBrushes.)
        _customAccentHex = string.IsNullOrWhiteSpace(settings.CustomAccentHex)
            ? "#8C00FF" : settings.CustomAccentHex;
        _tabSize = NormalizeTabSize(settings.TabSize);
        _editorFontSize = settings.EditorFontSize is >= 8 and <= 32 ? settings.EditorFontSize : 14;
        _terminalPanelHeight = NormalizeTerminalPanelHeight(settings.TerminalPanelHeight);
        _userCountry = string.IsNullOrWhiteSpace(settings.UserCountry)
            ? DetectCountryCode()
            : settings.UserCountry.ToUpperInvariant();
        _userHemisphere     = settings.UserHemisphere is >= 0 and <= 2 ? settings.UserHemisphere : 0;
        _userTimezoneOffset = settings.UserTimezoneOffset ?? string.Empty;
        _userName           = settings.UserName ?? string.Empty;
        _lastSeenVersion    = settings.LastSeenVersion ?? string.Empty;
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
        _extensionAutoUpdateTimer.Tick += ExtensionAutoUpdateTimer_OnTick;
        _appUpdateScheduler = new AppUpdateScheduler(
            isEnabled: () => IsAutoUpdateAppEnabled,
            isManualCheckInProgress: () => IsCheckingForUpdatesManually,
            installInBackground: () => IsAutoUpdateAppInBackgroundEnabled);
        _marketplaceRefreshTimer.Tick += MarketplaceRefreshTimer_OnTick;

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

        // Migration: if the user saved "kodo" mode but a theme is active (legacy behaviour
        // where "kodo" followed the theme accent), silently upgrade to "theme" mode so
        // the Kodo purple option remains independently selectable going forward.
        if (_accentColorMode == "kodo" && _hasThemeAccent)
            _accentColorMode = "theme";

        DataContext = this;
        IsHomePageVisible = true;

        // Kick off the full async refresh (marketplace fetch, icon loading, etc.)
        // now that the UI is live. The synchronous LoadExtensions() above already
        // populated the theme; the async path below will update everything else
        // (marketplace data, update badges) without touching the brush values again
        // unless the user has changed settings while offline.
        UpdateDiscordRichPresenceLifecycle();
        UpdateExtensionAutoUpdateLifecycle();
        _appUpdateScheduler.UpdateLifecycle();
        _marketplaceRefreshTimer.Start();
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

            // Same approach for the System Default theme blob - poll the
            // light/dark registry value every 2 s so it stays live too.
            _lastSeenWindowsThemeName = ResolveSystemThemeName();
            _windowsThemePollTimer.Tick += WindowsThemePollTimer_OnTick;
            _windowsThemePollTimer.Start();
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

    private async Task RefreshExtensionsDataAsync(bool force = false, bool suppressWatchdog = false)
    {
        if (_isRefreshingExtensions)
            return;

        if (!force && DateTime.UtcNow - _lastExtensionsRefreshUtc < ExtensionsRefreshCooldown)
            return;

        IsRefreshingExtensions = true;
        ExtensionsStatusText = "Refreshing extensions...";

        // ── Refresh watchdog ────────────────────────────────────────────────
        // If the entire refresh (disk scan + marketplace fetch) does not complete
        // within GitHubOperationTimeout (7 s) we fire the standard Kodo warning
        // dialog so the user knows the panel is stuck - not just silently spinning.
        //
        // HOW IT WORKS:
        //   A CancellationTokenSource drives a parallel Task.Delay watchdog.
        //   The watchdog races the real work; whichever finishes first wins.
        //   • Real work finishes first  → watchdog CTS is cancelled, delay exits
        //     cleanly, no dialog is shown.
        //   • Watchdog fires first      → the delay elapses, the watchdog posts
        //     a TimeoutException to the UI thread (dialog + log), then exits.
        //     The real work is NOT hard-cancelled: network ops already carry their
        //     own 7-second CancellationToken (RunWithGitHubTimeoutAsync), so they
        //     will fail shortly after and clean up normally via the catch block.
        //
        // WHY NOT Task.WhenAny + hard cancel on the work task?
        //   The work awaits Dispatcher.UIThread.InvokeAsync in several places.
        //   Cancelling an InvokeAsync continuation from a background token is not
        //   safe - it can leave the ObservableCollections mid-mutation and corrupt
        //   the binding state.  The watchdog-only approach lets the work unwind
        //   itself cleanly while giving the user immediate feedback.
        //
        // suppressWatchdog = true when called as a sub-step of an install/uninstall.
        // In those flows the outer operation already owns error reporting (download
        // timeout, write failure, etc.) and has its own RunWithGitHubTimeoutAsync
        // guard on the network work.  Firing a second independent watchdog here
        // would race against the outer handler and produce a misleading
        // "Marketplace refresh" dialog for what is actually an install timeout.
        using var watchdogCts = new CancellationTokenSource();
        var watchdogToken = watchdogCts.Token;
        if (!suppressWatchdog)
        {
            _ = Task.Run(async () =>
            {
                try
                {
                    await Task.Delay(GitHubOperationTimeout, watchdogToken);
                }
                catch (OperationCanceledException)
                {
                    // Work finished in time - nothing to do.
                    return;
                }

                // Still refreshing after 7 s: build a descriptive TimeoutException
                // and surface it through the standard Kodo warning dialog + log.
                var timeoutEx = new TimeoutException(
                    $"Marketplace refresh did not complete within " +
                    $"{GitHubOperationTimeout.TotalSeconds:0} seconds. " +
                    "This may indicate a slow or stalled network connection, " +
                    "a slow disk scan, or a hung extension operation.");

                KodoDiagnostics.LogWarning(
                    source: "MainWindow.RefreshExtensionsDataAsync.Watchdog",
                    exception: timeoutEx,
                    operation: "Marketplace refresh watchdog");

                await Dispatcher.UIThread.InvokeAsync(async () =>
                {
                    // Update status bar so the panel itself reflects the stall.
                    ExtensionsStatusText = "Marketplace refresh is taking too long. Check your connection.";
                    await ShowWarningDialogAsync("Marketplace refresh", timeoutEx);
                });
            }, watchdogToken);
        }

        try
        {
            // ScanInstalledExtensions is pure I/O - run it off the UI thread.
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
                var installedCount = VisibleLoadedExtensions.Count();
                var marketplaceCount = MarketplaceExtensions.Count;
                var installedWord = installedCount == 1 ? "extension" : "extensions";
                var marketplaceWord = marketplaceCount == 1 ? "extension" : "extensions";
                var updateWord = updateCount == 1 ? "update" : "updates";
                ExtensionsStatusText = updateCount > 0
                    ? $"Found {installedCount} installed {installedWord} and {marketplaceCount} in the marketplace. {updateCount} {updateWord} available."
                    : $"Found {installedCount} installed {installedWord} and {marketplaceCount} in the marketplace.";
                _lastExtensionsRefreshUtc = DateTime.UtcNow;
            });
        }
        catch (Exception ex)
        {
            await Dispatcher.UIThread.InvokeAsync(() => ExtensionsStatusText = "Couldn't refresh extensions. Check your connection and try again.");
            await Dispatcher.UIThread.InvokeAsync(async () => await ShowWarningDialogAsync("Marketplace fetch", ex));
        }
        finally
        {
            // Cancel the watchdog whether we succeeded, timed out, or threw - it
            // must not fire after IsRefreshingExtensions has been cleared.
            await watchdogCts.CancelAsync();
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
                    Console.WriteLine($"[Extensions] Failed to load '{Path.GetFileName(koxFile)}': {ex.Message}");
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
                    Console.WriteLine($"[Extensions] Failed to load folder extension '{Path.GetFileName(dir)}': {ex.Message}");
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
        _compiledSyntaxProfileCache.Clear();
        _contentSniffCache.Clear();
        SyncObservableCollection(LoadedExtensions, result.Extensions, ext => ext.Id);
        SyncObservableCollection(ExtensionLoadErrors, result.LoadErrors, error => error);

        // Decode icon bitmaps here on the UI thread. The background scan stored raw
        // PNG/SVG bytes in IconBytes to avoid creating Avalonia Bitmaps off-thread (which
        // is unsafe and causes silent failures). Now that we're on the UI thread we
        // can safely decode them and clear the staging bytes to free the memory.
        foreach (var ext in LoadedExtensions)
        {
            if (ext.IconImage is null && ext.SvgData is null && ext.IconBytes is not null)
            {
                if (IsSvgContent(ext.IconBytes))
                {
                    try { ext.SvgData = System.Text.Encoding.UTF8.GetString(ext.IconBytes); }
                    catch { /* malformed SVG - leave icon absent */ }
                }
                else
                {
                    ext.IconImage = DecodeBitmapOnUiThread(ext.IconBytes);
                }
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
        OnPropertyChanged(nameof(GroupedThemeExtensions));
        OnPropertyChanged(nameof(HasGroupedThemeExtensions));
        RefreshExtensionTheme();
        SyncMarketplaceInstallStates();
    }

    private async Task LoadMarketplaceExtensionsAsync()
    {
        var marketplaceExtensions = new List<MarketplaceExtension>();
        var extensionLoadErrors = new List<string>();

        await Dispatcher.UIThread.InvokeAsync(() => RefreshMarketplaceConnectivityState());

        // -- Disk-cache fast path ---------------------------
        // Seed the collection from the on-disk cache so the marketplace appears
        // immediately even before the network round-trip completes. Subsequent
        // refreshes in the same session skip this (collection is already populated).
        var diskJson = TryReadMarketplaceIndexCache();
        if (diskJson is not null && !MarketplaceExtensions.Any())
            ParseAndApplyMarketplaceIndex(diskJson, marketplaceExtensions, extensionLoadErrors);

        try
        {
            // -- Conditional ETag fetch --------------------------
            // Send If-None-Match with the stored ETag so GitHub can reply 304 Not
            // Modified when the index hasn't changed.  304 responses do NOT count
            // against the 60 req/hr anonymous rate limit, so repeated refreshes
            // (startup, post-install, manual) are free as long as nothing changed.
            // Only a real 200 response burns one rate-limit slot.
            _marketplaceIndexETag ??= TryReadMarketplaceIndexETag();
            using var indexRequest = new HttpRequestMessage(HttpMethod.Get, DefaultMarketplaceIndexUrl);
            indexRequest.Headers.Accept.ParseAdd("application/vnd.github.raw+json");
            if (_marketplaceIndexETag is not null)
                indexRequest.Headers.TryAddWithoutValidation("If-None-Match", _marketplaceIndexETag);

            var (statusCode, remoteJson, newETag) = await RunWithGitHubTimeoutAsync(
                "Marketplace index fetch",
                async ct =>
                {
                    using var indexResponse = await MarketplaceHttpClient.SendAsync(indexRequest, ct);
                    if ((int)indexResponse.StatusCode == 304)
                        return (304, (string?)null, (string?)null);
                    indexResponse.EnsureSuccessStatusCode();
                    var body = await indexResponse.Content.ReadAsStringAsync(ct);
                    var etag = indexResponse.Headers.ETag?.Tag;
                    return (200, body, etag);
                });

            if (statusCode == 304)
            {
                // Index unchanged - reuse whatever is already in the collection
                // (disk-cache seed or previous refresh).  No parse, no disk write.
                KodoDiagnostics.LogDebug("Marketplace index: 304 Not Modified - reusing cached data.");
            }
            else if (remoteJson is not null)
            {
                // Fresh 200 - parse, update disk cache, store new ETag.
                marketplaceExtensions.Clear();
                extensionLoadErrors.Clear();
                ParseAndApplyMarketplaceIndex(remoteJson, marketplaceExtensions, extensionLoadErrors);
                TryWriteMarketplaceIndexCache(remoteJson);
                if (newETag is not null)
                {
                    _marketplaceIndexETag = newETag;
                    TryWriteMarketplaceIndexETag(newETag);
                }
            }
        }
        catch (Exception ex)
        {
            if (diskJson is not null)
            {
                // Network failed but a cached copy exists - the marketplace was
                // already seeded above, so stay usable.  Log without a dialog.
                extensionLoadErrors.Add($"Marketplace index fetch failed (using cached copy): {DescribeFetchFailure(ex)}");
                KodoDiagnostics.LogDebug("Marketplace index fetch failed; using disk cache.", ex);
                await Dispatcher.UIThread.InvokeAsync(() => RefreshMarketplaceConnectivityState("Marketplace fetch", ex));
            }
            else
            {
                // No cache at all - propagate so the caller shows the error dialog.
                extensionLoadErrors.Add($"Failed to load remote marketplace index: {DescribeFetchFailure(ex)}");
                await Dispatcher.UIThread.InvokeAsync(() => RefreshMarketplaceConnectivityState("Marketplace fetch", ex));
                throw;
            }
        }

        // All ObservableCollection mutations and PropertyChanged notifications must
        // run on the UI thread - Avalonia's binding engine requires it.
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
            OnPropertyChanged(nameof(IsMarketplaceUnavailableVisible));
            OnPropertyChanged(nameof(IsMarketplacePartialErrorVisible));
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
                            // Index icon fetched successfully - use it, replacing any kox icon.
                            ReplaceLoadedExtensionIcon(pair.ext, icon);
                        }
                        // else: fetch returned nothing (bad URL, corrupt bytes, etc.) -
                        // leave whatever the kox provided in place.
                    });
                }
                catch (Exception ex)
                {
                    // Network failure for this icon - leave the kox icon (or abbreviation) in place.
                    KodoDiagnostics.LogDebug($"Icon fetch failed for installed extension '{pair.ext.Id}': {pair.iconUrl}", ex);
                }
            });

        await Task.WhenAll(tasks);
    }
    private async Task FetchMarketplaceIconsAsync(IReadOnlyDictionary<string, string> marketplaceIconMap)
    {
        // Apply icons whose bytes are already cached synchronously on the UI thread -
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

        var iconFailures = 0;
        var iconAttempts = 0;
        Exception? lastIconException = null;

        var tasks = MarketplaceExtensions
            .Where(entry => entry.IconImage is null && entry.SvgData is null && marketplaceIconMap.TryGetValue(entry.Id, out _))
            .Select(async entry =>
            {
                Interlocked.Increment(ref iconAttempts);
                try
                {
                    var icon = await GetCachedIconAsync(marketplaceIconMap[entry.Id]);
                    if (!icon.HasValue)
                    {
                        KodoDiagnostics.LogDebug($"Icon fetch returned no data for marketplace extension '{entry.Id}': {marketplaceIconMap[entry.Id]}");
                        Interlocked.Increment(ref iconFailures);
                        return;
                    }

                    await Dispatcher.UIThread.InvokeAsync(() =>
                    {
                        ReplaceMarketplaceIcon(entry, icon);
                    });
                }
                catch (Exception ex)
                {
                    KodoDiagnostics.LogDebug($"Icon fetch failed for marketplace extension '{entry.Id}': {marketplaceIconMap[entry.Id]}", ex);
                    Interlocked.Increment(ref iconFailures);
                    Interlocked.Exchange(ref lastIconException, ex);
                }
            });

        await Task.WhenAll(tasks);

        // Log a warning if every icon fetch failed - suggests a network or rate-limit
        // problem - but don't surface a dialog: icons are decorative and a modal here
        // would block the marketplace from appearing while extensions loaded fine.
        if (iconAttempts > 0 && iconFailures == iconAttempts && lastIconException is not null)
        {
            KodoDiagnostics.LogDebug(
                $"All {iconAttempts} marketplace icon fetch(es) failed; icons will show abbreviations.",
                lastIconException);
        }
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

        // Cache miss - fetch under semaphore to avoid duplicate requests.
        await _iconFetchSemaphore.WaitAsync();
        try
        {
            if (!_marketplaceIconBytesCache.TryGetValue(iconUrl, out bytes))
            {
                // Per-request timeout so a single stalled icon fetch cannot hold
                // up the entire FetchMarketplaceIconsAsync Task.WhenAll.
                // Uses GitHubOperationTimeout (7 s) - the same ceiling applied to
                // every other GitHub operation - so icon fetches are consistent.
                using var cts = new CancellationTokenSource(GitHubOperationTimeout);

                // Kodo-hosted icon URLs use the GitHub Contents API
                // (api.github.com/repos/.../contents/...) so the raw+json Accept header
                // is required to get file bytes directly rather than a base64-wrapped
                // JSON response.  Third-party URLs (Wikipedia, etc.) are plain GETs.
                using var request = new HttpRequestMessage(HttpMethod.Get, iconUrl);
                if (IsGitHubContentsApiUrl(iconUrl))
                    request.Headers.Accept.ParseAdd("application/vnd.github.raw+json");

                using var response = await MarketplaceHttpClient.SendAsync(request, cts.Token);
                response.EnsureSuccessStatusCode();
                bytes = await response.Content.ReadAsByteArrayAsync(cts.Token);
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
            LatestReleaseStatusText = $"Could not load release info: {DescribeFetchFailure(ex)}";

            // Log and surface the Kodo warning dialog so the user knows why the
            // release panel is empty (timeout, rate-limit, no connectivity, etc.).
            KodoDiagnostics.LogDebug("Failed to fetch latest release info", ex);
            await ShowWarningDialogAsync("Latest release info fetch", ex);
        }
        finally
        {
            _isRefreshingLatestRelease = false;
            OnPropertyChanged(nameof(IsRefreshingLatestRelease));
            OnPropertyChanged(nameof(RefreshLatestReleaseButtonText));
        }
    }

    private async Task FetchAnnouncementsAsync()
    {
        IsNewsLoading = true;
        IsNewsError = false;
        NewsItems.Clear();

        try
        {
            // Use the GitHub Contents API with the raw+json Accept header so GitHub
            // returns the file bytes directly (no base64 wrapping) at the same
            // generous rate limits as the marketplace index fetch.
            // The entire operation - headers + body - must complete within
            // GitHubOperationTimeout (7 s); a stall past that point throws
            // TimeoutException and surfaces the standard Kodo error dialog.
            using var request = new HttpRequestMessage(HttpMethod.Get, AnnouncementsUrl);
            request.Headers.Accept.ParseAdd("application/vnd.github.raw+json");
            request.Headers.CacheControl = new System.Net.Http.Headers.CacheControlHeaderValue { NoCache = true, NoStore = true };

            var (response, md) = await RunWithGitHubTimeoutAsync(
                "News / Announcements fetch",
                async ct =>
                {
                    using var resp = await MarketplaceHttpClient.SendAsync(request, ct);
                    resp.EnsureSuccessStatusCode();
                    var body = await resp.Content.ReadAsStringAsync(ct);
                    // Return a tuple; ownership of HttpResponseMessage stays inside the
                    // lambda so the using-scope disposes it after we've read the body.
                    return (resp, body);
                });

            var items = ParseAnnouncementsMd(md);
            foreach (var item in items)
                NewsItems.Add(item);
        }
        catch (Exception ex)
        {
            KodoDiagnostics.LogDebug("Failed to fetch announcements", ex);
            IsNewsError = true;

            // Show the Kodo warning dialog so the user knows the news panel is stale
            // and can see the specific reason (timeout, HTTP error, parse failure, etc.).
            // Fire-and-forget is intentional: IsNewsLoading must be cleared immediately
            // in the finally block; we must not await the modal here.
            _ = ShowWarningDialogAsync("News / Announcements fetch", ex);
        }
        finally
        {
            IsNewsLoading = false;
            OnPropertyChanged(nameof(IsNewsContentVisible));
            OnPropertyChanged(nameof(IsNewsEmpty));
        }
    }

    // Parses the ANNOUNCEMENTS.md format:
    //   ## Title
    //   > 2024-06-01          ← optional date line (blockquote)
    //   body lines...
    //   ---   (separator between posts)
    //
    // Items are returned in reverse order so the last entry in the file
    // appears at the top of the news panel.
    private static List<NewsItem> ParseAnnouncementsMd(string md)
    {
        var items = new List<NewsItem>();
        // Split on the horizontal rule separator
        var sections = md.Split(["---"], StringSplitOptions.RemoveEmptyEntries);
        foreach (var section in sections)
        {
            var lines = section.Split('\n')
                               .Select(l => l.TrimEnd('\r').Trim())
                               .ToList();

            string? title = null;
            string? updatedAt = null;
            var bodyLines = new List<string>();

            foreach (var line in lines)
            {
                if (title is null && line.StartsWith("## "))
                {
                    title = line[3..].Trim();
                }
                else if (title is not null && updatedAt is null && line.StartsWith("> "))
                {
                    // A blockquote immediately after the heading is treated as the date.
                    // Try to parse yyyy-MM-dd and reformat to e.g. "June 12, 2026";
                    // fall back to the raw string if it doesn't match that pattern.
                    var raw = line[2..].Trim();
                    updatedAt = DateTime.TryParseExact(
                        raw,
                        "yyyy-MM-dd",
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None,
                        out var parsed)
                        ? parsed.ToString("MMMM d, yyyy", System.Globalization.CultureInfo.InvariantCulture)
                        : raw;
                }
                else if (title is not null && line.Length > 0)
                {
                    bodyLines.Add(line);
                }
            }

            if (title is null && bodyLines.Count == 0) continue;

            items.Add(new NewsItem
            {
                Title     = title ?? string.Empty,
                Body      = string.Join("\n", bodyLines).Trim(),
                UpdatedAt = updatedAt ?? string.Empty,
            });
        }

        // Reverse so the last-written entry in the file surfaces at the top.
        items.Reverse();
        return items;
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

        return await RunWithGitHubTimeoutAsync<ReleaseInfo?>(
            "Latest stable release fetch",
            async ct =>
            {
                using var response = await MarketplaceHttpClient.SendAsync(request, ct);
                if (!response.IsSuccessStatusCode)
                    return null;

                var json = await response.Content.ReadAsStringAsync(ct);
                using var doc = JsonDocument.Parse(json);
                return ParseReleaseInfo(doc.RootElement);
            });
    }

    private async Task<ReleaseInfo?> TryFetchLatestListedReleaseAsync()
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, ReleasesApiUrl);
        request.Headers.Accept.ParseAdd("application/vnd.github+json");

        return await RunWithGitHubTimeoutAsync<ReleaseInfo?>(
            "Releases list fetch",
            async ct =>
            {
                using var response = await MarketplaceHttpClient.SendAsync(request, ct);
                response.EnsureSuccessStatusCode();

                var json = await response.Content.ReadAsStringAsync(ct);
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
            });
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

    // Extracts marketplace entries from a raw JSON string and appends them to
    // the provided lists.  Shared by the disk-cache seed path and the live-200
    // parse path so both go through identical validation logic.
    private static void ParseAndApplyMarketplaceIndex(
        string json,
        List<MarketplaceExtension> marketplaceExtensions,
        List<string> extensionLoadErrors)
    {
        var jsonOptions = new JsonDocumentOptions
        {
            AllowTrailingCommas = true,
            CommentHandling = JsonCommentHandling.Skip
        };
        using var doc = JsonDocument.Parse(json, jsonOptions);
        if (!doc.RootElement.TryGetProperty("extensions", out var extensionsElement) ||
            extensionsElement.ValueKind != JsonValueKind.Array)
            return;

        foreach (var item in extensionsElement.EnumerateArray())
        {
            try
            {
                var entry = ParseMarketplaceExtension(item);
                if (string.IsNullOrWhiteSpace(entry.Id) || marketplaceExtensions.Any(e => e.Id == entry.Id))
                    continue;
                marketplaceExtensions.Add(entry);
            }
            catch (Exception itemEx)
            {
                var entryId = item.TryGetProperty("id", out var idEl) ? idEl.GetString() ?? "?" : "?";
                extensionLoadErrors.Add($"Skipped malformed marketplace entry '{entryId}': {itemEx.Message}");
                KodoDiagnostics.LogDebug($"Skipped malformed marketplace entry '{entryId}'", itemEx);
            }
        }
    }

    private string? TryReadMarketplaceIndexCache()
    {
        try { return File.Exists(MarketplaceIndexCachePath) ? File.ReadAllText(MarketplaceIndexCachePath, System.Text.Encoding.UTF8) : null; }
        catch (Exception ex) { KodoDiagnostics.LogDebug("Could not read marketplace index cache.", ex); return null; }
    }

    private void TryWriteMarketplaceIndexCache(string json)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(MarketplaceIndexCachePath)!);
            File.WriteAllText(MarketplaceIndexCachePath, json, System.Text.Encoding.UTF8);
        }
        catch (Exception ex) { KodoDiagnostics.LogDebug("Could not write marketplace index cache.", ex); }
    }

    private string? TryReadMarketplaceIndexETag()
    {
        try { return File.Exists(MarketplaceIndexETagPath) ? File.ReadAllText(MarketplaceIndexETagPath).Trim() : null; }
        catch { return null; }
    }

    private void TryWriteMarketplaceIndexETag(string etag)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(MarketplaceIndexETagPath)!);
            File.WriteAllText(MarketplaceIndexETagPath, etag);
        }
        catch (Exception ex) { KodoDiagnostics.LogDebug("Could not write marketplace index ETag.", ex); }
    }

    private static MarketplaceExtension ParseMarketplaceExtension(JsonElement item)
    {
        var id = item.TryGetProperty("id", out var idElement) ? idElement.GetString() ?? string.Empty : string.Empty;
        var declaredVersion = item.TryGetProperty("version", out var versionElement) ? versionElement.GetString() ?? string.Empty : string.Empty;
        var name = item.TryGetProperty("name", out var nameElement) ? nameElement.GetString() ?? string.Empty : string.Empty;
        var type = item.TryGetProperty("type", out var typeElement) ? typeElement.GetString() ?? string.Empty : string.Empty;
        var author = item.TryGetProperty("author", out var authorElement) ? authorElement.GetString() ?? string.Empty : string.Empty;
        var description = item.TryGetProperty("description", out var descriptionElement) ? descriptionElement.GetString() ?? string.Empty : string.Empty;
        var rawDownloadUrl = NormalizeGitHubBlobViewerUrl(
            item.TryGetProperty("downloadUrl", out var downloadUrlElement) ? downloadUrlElement.GetString() ?? string.Empty : string.Empty);
        var declaredFileName = item.TryGetProperty("fileName", out var fileNameElement) ? fileNameElement.GetString() ?? string.Empty : string.Empty;
        var iconUrl = NormalizeGitHubUrl(
            item.TryGetProperty("iconUrl", out var iconUrlElement) ? iconUrlElement.GetString() ?? string.Empty : string.Empty);
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
        OnPropertyChanged(nameof(AutoUpdateExtensionsStatusText));
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

            // Download the package under GitHubOperationTimeout (7 s) so a stalled
            // download cannot silently hold up the entire install pipeline and then
            // trigger the RefreshExtensionsDataAsync watchdog as a false positive.
            // Previously used GetByteArrayAsync with no CancellationToken, meaning
            // the download could block for the full 30-second HttpClient.Timeout.
            var bytes = await RunWithGitHubTimeoutAsync(
                $"Extension download - {marketplaceExtension.Name}",
                async ct =>
                {
                    using var downloadRequest = new HttpRequestMessage(HttpMethod.Get, marketplaceExtension.DownloadUrl);
                    // Contents API URLs require raw+json to receive file bytes directly
                    // instead of a base64-wrapped JSON envelope.
                    if (IsGitHubContentsApiUrl(marketplaceExtension.DownloadUrl))
                        downloadRequest.Headers.Accept.ParseAdd("application/vnd.github.raw+json");
                    using var downloadResponse = await MarketplaceHttpClient.SendAsync(
                        downloadRequest, HttpCompletionOption.ResponseContentRead, ct);
                    downloadResponse.EnsureSuccessStatusCode();
                    return await downloadResponse.Content.ReadAsByteArrayAsync(ct);
                });

            ValidateDownloadedExtensionPackage(marketplaceExtension, bytes);
            DeleteInstalledExtensionSources(marketplaceExtension.Id, outputPath);
            await File.WriteAllBytesAsync(outputPath, bytes);
            NormalizeKoxManifestVersion(outputPath);

            // suppressWatchdog: true - the download above already carried its own
            // RunWithGitHubTimeoutAsync guard. A second independent watchdog here
            // would race the outer install handler and emit a misleading
            // "Marketplace refresh" dialog for what was actually an install timeout.
            await RefreshExtensionsDataAsync(force: true, suppressWatchdog: true);
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

            // suppressWatchdog: true - uninstall is a local disk operation with its
            // own error handling above; the watchdog is not meaningful here.
            await RefreshExtensionsDataAsync(force: true, suppressWatchdog: true);
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
    /// Converts a GitHub blob viewer URL
    /// (https://github.com/{owner}/{repo}/blob/{branch}/{path}) to the
    /// Contents API (https://api.github.com/repos/{owner}/{repo}/contents/{path})
    /// so it can be fetched as raw bytes.  Raw CDN URLs
    /// (raw.githubusercontent.com), Contents API URLs, and all other URLs are
    /// returned unchanged - this rewrite is only needed for blob viewer links,
    /// which return an HTML page rather than file bytes when fetched directly.
    /// </summary>
    private static string NormalizeGitHubBlobViewerUrl(string url)
    {
        if (string.IsNullOrWhiteSpace(url))
            return url;

        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
            return url;

        // Only rewrite github.com /blob/ viewer URLs - everything else is already fetchable.
        if (!uri.Host.Equals("github.com", StringComparison.OrdinalIgnoreCase))
            return url;

        var segments = uri.AbsolutePath.TrimStart('/').Split('/');
        if (segments.Length >= 5 && segments[2].Equals("blob", StringComparison.OrdinalIgnoreCase))
        {
            var owner = segments[0];
            var repo  = segments[1];
            var path  = string.Join("/", segments, 4, segments.Length - 4);
            return $"https://api.github.com/repos/{owner}/{repo}/contents/{path}";
        }

        return url; // non-blob github.com URL (e.g. releases page) - leave alone
    }

    /// <summary>
    /// Normalises any GitHub file URL to the Contents API so it can be fetched
    /// as raw bytes using the <c>application/vnd.github.raw+json</c> Accept header.
    ///
    /// All three GitHub URL forms are accepted and rewritten:
    ///   https://github.com/{owner}/{repo}/blob/{branch}/{path}      (viewer)
    ///   https://raw.githubusercontent.com/{owner}/{repo}/{branch}/{path}  (raw CDN)
    ///   https://api.github.com/repos/{owner}/{repo}/contents/{path} (already correct)
    ///   -> https://api.github.com/repos/{owner}/{repo}/contents/{path}
    ///
    /// This means contributors can use any of the three forms in the extension
    /// index and the app will normalise them transparently at parse time.
    /// Third-party URLs (e.g. Wikipedia, Wikimedia) are returned unchanged.
    /// </summary>
    private static string NormalizeGitHubUrl(string url)
    {
        if (string.IsNullOrWhiteSpace(url))
            return url;

        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
            return url;

        // github.com/blob viewer URL
        // /{owner}/{repo}/blob/{branch}/{...path} -> Contents API
        if (uri.Host.Equals("github.com", StringComparison.OrdinalIgnoreCase))
        {
            var segments = uri.AbsolutePath.TrimStart('/').Split('/');
            if (segments.Length >= 5 &&
                segments[2].Equals("blob", StringComparison.OrdinalIgnoreCase))
            {
                var owner = segments[0];
                var repo  = segments[1];
                var path  = string.Join("/", segments, 4, segments.Length - 4);
                return $"https://api.github.com/repos/{owner}/{repo}/contents/{path}";
            }
            return url; // non-blob github.com URL (e.g. releases page) - leave alone
        }

        // raw.githubusercontent.com CDN URL
        // /{owner}/{repo}/{branch}/{...path} -> Contents API
        if (uri.Host.Equals("raw.githubusercontent.com", StringComparison.OrdinalIgnoreCase))
        {
            var segments = uri.AbsolutePath.TrimStart('/').Split('/');
            if (segments.Length >= 4)
            {
                var owner = segments[0];
                var repo  = segments[1];
                // segments[2] is the branch - omitted from the Contents API path
                var path  = string.Join("/", segments, 3, segments.Length - 3);
                return $"https://api.github.com/repos/{owner}/{repo}/contents/{path}";
            }
            return url;
        }

        // Already a Contents API URL or a third-party URL - leave unchanged.
        return url;
    }

    /// <summary>
    /// Returns true when <paramref name="url"/> is a GitHub Contents API endpoint
    /// (api.github.com/repos/.../contents/...).  Used to decide whether to add
    /// the <c>application/vnd.github.raw+json</c> Accept header to a request so
    /// GitHub returns file bytes directly instead of a base64-wrapped JSON object.
    /// </summary>
    private static bool IsGitHubContentsApiUrl(string url) =>
        !string.IsNullOrWhiteSpace(url) &&
        url.StartsWith("https://api.github.com/repos/", StringComparison.OrdinalIgnoreCase) &&
        url.Contains("/contents/", StringComparison.OrdinalIgnoreCase);

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
        else
        {
            var svgIconPath = Path.Combine(folderPath, "icon.svg");
            if (File.Exists(svgIconPath))
            {
                using var iconStream = File.OpenRead(svgIconPath);
                baseExt.IconBytes = ReadIconBytesFromStream(iconStream);
            }
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

        var iconEntry = archive.GetEntry("icon.png") ?? archive.GetEntry("icon.svg");
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
            // Result is cached per path so we only read the file once per session.
            if (!_contentSniffCache.TryGetValue(filePath, out var sniffed))
            {
                sniffed = TryDetectLanguageFromContent(filePath);
                _contentSniffCache[filePath] = sniffed;
            }
            if (sniffed is null)
                return null;

            // Content-sniffed match: use the base extension as-is (no profile narrowing).
            return sniffed;
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
        var syntaxProfile = ResolveCompiledSyntaxProfile(ext);
        if (!_highlightingCache.TryGetValue(ext, out var definition))
        {
            definition = new KodoHighlightingDefinition(ext, syntaxProfile);
            _highlightingCache[ext] = definition;
        }
        EditorTextBox.SyntaxHighlighting = definition;
        ConfigureRainbowBrackets(ext);
        ConfigureInterpolatedStrings(syntaxProfile);
        ConfigureHtmlEmbeddedHighlighting(ext);
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
        if (!ReferenceEquals(CurrentImagePreview, imagePreview))
            ImageZoomLevel = 1.0;
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
        ConfigureHtmlEmbeddedHighlighting(null);
        ConfigureMarkdownHighlighting(null);
    }

    private CompiledSyntaxProfile ResolveCompiledSyntaxProfile(LoadedExtension extension)
    {
        if (_compiledSyntaxProfileCache.TryGetValue(extension, out var cached))
            return cached;

        var profile = CompiledSyntaxProfile.Create(extension);
        _compiledSyntaxProfileCache[extension] = profile;
        return profile;
    }

    private void ConfigureHtmlEmbeddedHighlighting(LoadedExtension? extension)
    {
        _htmlEmbeddedColorizer.UpdateSyntax(extension, ResolveHtmlEmbeddedSyntaxProfile);
        EditorTextBox?.TextArea.TextView.InvalidateLayer(KnownLayer.Text);
    }

    private void ConfigureMarkdownHighlighting(LoadedExtension? extension)
    {
        _markdownColorizer.UpdateSyntax(extension, ResolveFenceLanguageSyntaxProfile, ResolveInlineCodeLanguageExtension);
        EditorTextBox?.TextArea.TextView.InvalidateLayer(KnownLayer.Text);
    }

    private CompiledSyntaxProfile? ResolveHtmlEmbeddedSyntaxProfile(string blockTag, string? typeAttribute)
    {
        static bool ContainsToken(string value, params string[] needles) =>
            needles.Any(needle => value.Contains(needle, StringComparison.OrdinalIgnoreCase));

        var normalizedTag = blockTag.Trim().ToLowerInvariant();
        var normalizedType = (typeAttribute ?? string.Empty).Trim();

        if (normalizedTag == "x:code")
            return FindLanguageSyntaxProfileForFileExtension(".cs");

        if (normalizedTag == "style")
            return FindLanguageSyntaxProfileForFileExtension(".css");

        if (normalizedTag != "script")
            return null;

        if (string.IsNullOrWhiteSpace(normalizedType) ||
            string.Equals(normalizedType, "module", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(normalizedType, "text/javascript", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(normalizedType, "application/javascript", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(normalizedType, "application/ecmascript", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(normalizedType, "text/ecmascript", StringComparison.OrdinalIgnoreCase))
        {
            return FindLanguageSyntaxProfileForFileExtension(".js");
        }

        if (ContainsToken(normalizedType, "typescript"))
            return FindLanguageSyntaxProfileForFileExtension(".ts");

        if (ContainsToken(normalizedType, "json", "importmap"))
            return FindLanguageSyntaxProfileForFileExtension(".json");

        if (ContainsToken(normalizedType, "javascript", "ecmascript", "jscript"))
            return FindLanguageSyntaxProfileForFileExtension(".js");

        if (ContainsToken(normalizedType, "css"))
            return FindLanguageSyntaxProfileForFileExtension(".css");

        return null;
    }

    private CompiledSyntaxProfile? FindLanguageSyntaxProfileForFileExtension(string extension)
    {
        var loadedExtension = LoadedExtensions.FirstOrDefault(loadedExtension =>
            string.Equals(loadedExtension.Type, "language", StringComparison.OrdinalIgnoreCase) &&
            loadedExtension.Extensions.Any(ext => string.Equals(ext, extension, StringComparison.OrdinalIgnoreCase)));

        return loadedExtension is null ? null : ResolveCompiledSyntaxProfile(loadedExtension);
    }

    private CompiledSyntaxProfile? ResolveFenceLanguageSyntaxProfile(string fenceLanguage)
    {
        var extension = ResolveFenceLanguageExtension(fenceLanguage);
        return extension is null ? null : ResolveCompiledSyntaxProfile(extension);
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
        {
            if (string.IsNullOrEmpty(alias))
                return null; // explicit plain-text marker - no syntax profile
            normalized = alias;
        }

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

    // Inline-code language detection now lives in SyntaxColorEngine.cs
    // (InlineCodeLanguageDetector) so all syntax-matching logic stays in the
    // engine rather than the window's code-behind. This wrapper just supplies
    // the currently loaded extensions.
    private LoadedExtension? ResolveInlineCodeLanguageExtension(string codeSnippet) =>
        InlineCodeLanguageDetector.Resolve(LoadedExtensions, codeSnippet);

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
        // Rainbow brackets have no meaning in plain text or markdown prose.
        // Markdown fenced code blocks are colourised by _markdownColorizer independently.
        var isMarkdown = string.Equals(ext?.Id, "markdown-kodo-extension", StringComparison.OrdinalIgnoreCase);
        _rainbowBracketColorizer.UpdateSyntax(isMarkdown ? null : ext);
        EditorTextBox?.TextArea.TextView.InvalidateLayer(KnownLayer.Text);
    }

    private void ConfigureInterpolatedStrings(CompiledSyntaxProfile? syntaxProfile)
    {
        _interpolatedStringColorizer.UpdateSyntax(syntaxProfile);
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

    public bool IsUpdateSplashVisible
    {
        get => _isUpdateSplashVisible;
        private set
        {
            if (_isUpdateSplashVisible == value) return;
            _isUpdateSplashVisible = value;
            OnPropertyChanged();
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

    /// <summary>
    /// ThemeExtensions grouped by extension name.  Multi-theme extensions
    /// (e.g. "Dark Themes" with 4 cards) become a single collapsible group;
    /// single-theme extensions appear as a group of one (always expanded,
    /// no collapse chrome).  Bind the Settings and tutorial theme lists to
    /// this instead of <see cref="ThemeExtensions"/>.
    /// </summary>
    public IEnumerable<ThemeExtensionGroup> GroupedThemeExtensions =>
        ThemeExtensions
            .GroupBy(e => e.Name, StringComparer.OrdinalIgnoreCase)
            .Select(g => new ThemeExtensionGroup(g.Key, g.ToList()));

    public bool HasGroupedThemeExtensions => ThemeExtensions.Any();

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
            OnPropertyChanged(nameof(IsMarketplaceUnavailableVisible));
            OnPropertyChanged(nameof(IsMarketplacePartialErrorVisible));
            OnPropertyChanged(nameof(IsMarketplaceEmptyVisible));
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
            OnPropertyChanged(nameof(LatestReleaseFormatted));
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

    private static bool IsCurrentNewerThanLastSeen(string lastSeen)
    {
        if (string.IsNullOrWhiteSpace(lastSeen)) return false;

        if (!Version.TryParse(StripPreRelease(lastSeen),          out var seen))    return false;
        if (!Version.TryParse(StripPreRelease(CurrentAppVersion), out var current)) return false;

        if (current != seen) return current > seen;

        // Same numeric version - current is "newer" only if its suffix has higher priority
        // (e.g. upgrading from v1.2.0-BETA to v1.2.0 stable should still show the splash).
        return VersionPriority(CurrentAppVersion) > VersionPriority(lastSeen);
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

    // Structured version of the release notes: a list of paragraphs, each containing
    // inline runs that carry bold/normal weight. Used by the formatted notes template
    // in both the Settings What's New card and the update splash screen.
    public IReadOnlyList<FormattedParagraph> LatestReleaseFormatted =>
        string.IsNullOrWhiteSpace(LatestRelease?.Notes)
            ? [new FormattedParagraph { Runs = [new FormattedRun { Text = "No release notes available." }] }]
            : ParseMarkdownParagraphs(LatestRelease.Notes);

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

    // Matches **bold** and __bold__ spans for inline run splitting.
    private static readonly Regex MdInlineBoldRegex = new(@"\*\*(.+?)\*\*|__(.+?)__", RegexOptions.Compiled | RegexOptions.Singleline);

    // Parses raw GitHub markdown into a flat list of FormattedParagraphs.
    // Each paragraph carries a list of FormattedRuns (bold / normal) so the
    // AXAML template can render them inline with correct FontWeight.
    // Handles: bullet lines (* / - / +), numbered lists, headings (as bold),
    // horizontal rules (skipped), code fences (stripped), and inline **bold**.
    private static IReadOnlyList<FormattedParagraph> ParseMarkdownParagraphs(string markdown)
    {
        if (string.IsNullOrWhiteSpace(markdown))
            return [];

        var raw = markdown.Replace("\r\n", "\n").Replace('\r', '\n');

        // Strip code fences entirely.
        raw = MdCodeFenceRegex.Replace(raw, string.Empty);
        raw = raw.Replace("```", string.Empty);

        var paragraphs = new List<FormattedParagraph>();

        foreach (var rawLine in raw.Split('\n'))
        {
            var line = rawLine.TrimEnd();

            // Skip horizontal rules and separator lines.
            if (MdHrRegex.IsMatch(line)) continue;

            // Blank line → skip (spacing is handled by TopMargin).
            if (string.IsNullOrWhiteSpace(line)) continue;

            var isBullet   = false;
            var isOrdered  = false;
            string? orderedPrefix = null;

            // Bullet list item.
            var bulletMatch = MdBulletRegex.Match(line);
            if (bulletMatch.Success)
            {
                isBullet = true;
                line = line[bulletMatch.Length..];
            }
            else
            {
                // Ordered list item.
                var orderedMatch = MdOrderedListRegex.Match(line);
                if (orderedMatch.Success)
                {
                    isOrdered     = true;
                    orderedPrefix = orderedMatch.Groups[1].Value + ".";
                    line          = line[orderedMatch.Length..];
                }
                else
                {
                    // Heading → strip markers, treat content as bold paragraph.
                    var headingMatch = MdHeadingRegex.Match(line);
                    if (headingMatch.Success)
                        line = line[headingMatch.Length..].Trim();
                }
            }

            // Strip blockquote markers.
            line = MdBlockquoteRegex.Replace(line, string.Empty);

            // Strip images, resolve links to display text only.
            line = MdImageRegex.Replace(line, "$1");
            line = MdLinkRegex.Replace(line, "$1");

            // Strip inline code backticks (keep the text).
            line = MdInlineCodeRegex.Replace(line, "$1");

            // Strip strikethrough.
            line = MdStrikethroughRegex.Replace(line, "$1");

            // Strip italic (single * or _) without consuming bold markers.
            line = MdItalicRegex.Replace(line, "$1");
            line = MdItalicUnderscoreRegex.Replace(line, "$1");

            line = line.Trim();
            if (string.IsNullOrEmpty(line)) continue;

            // Split the line into bold/normal runs. The list marker (if any)
            // is tracked separately rather than injected as a leading run, so
            // it can be rendered in its own fixed-width column and wrapped
            // lines hang-indent under the first word, not under the marker.
            var runs = new List<FormattedRun>();

            var marker = isBullet ? "•"
                : isOrdered && orderedPrefix is not null ? orderedPrefix
                : string.Empty;

            // Bullets need a narrow column; ordered markers ("1." .. "99.")
            // need a little more room so two-digit numbers don't clip.
            var markerColumnWidth = isBullet ? 18.0
                : isOrdered ? 28.0
                : 0.0;

            var pos = 0;
            foreach (Match m in MdInlineBoldRegex.Matches(line))
            {
                // Normal text before this bold span.
                if (m.Index > pos)
                    runs.Add(new FormattedRun { Text = line[pos..m.Index], IsBold = false });

                // The bold text itself (group 1 = **, group 2 = __).
                var boldText = m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value;
                if (!string.IsNullOrEmpty(boldText))
                    runs.Add(new FormattedRun { Text = boldText, IsBold = true });

                pos = m.Index + m.Length;
            }

            // Trailing normal text after the last bold span.
            if (pos < line.Length)
                runs.Add(new FormattedRun { Text = line[pos..], IsBold = false });

            if (runs.Count == 0) continue;

            paragraphs.Add(new FormattedParagraph
            {
                Runs              = runs,
                TopMargin         = isBullet || isOrdered ? new Thickness(0, 2, 0, 0) : new Thickness(0, 6, 0, 0),
                Marker            = marker,
                MarkerColumnWidth = markerColumnWidth,
            });
        }

        return paragraphs;
    }

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

    // True when the marketplace has NO entries and there was a connectivity error.
    public bool IsMarketplaceUnavailableVisible => MarketplaceExtensions.Count == 0 && IsMarketplaceConnectivityWarningVisible;
    // True when the marketplace has entries but some failed to load (partial error).
    public bool IsMarketplacePartialErrorVisible => MarketplaceExtensions.Count > 0 && IsMarketplaceConnectivityWarningVisible;
    public bool IsMarketplaceEmptyVisible => MarketplaceExtensions.Count == 0 && !IsRefreshingExtensions && !IsMarketplaceConnectivityWarningVisible;

    public bool IsMarketplaceConnectivityWarningVisible
    {
        get => _isMarketplaceConnectivityWarningVisible;
        private set
        {
            if (_isMarketplaceConnectivityWarningVisible == value) return;
            _isMarketplaceConnectivityWarningVisible = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsMarketplaceUnavailableVisible));
            OnPropertyChanged(nameof(IsMarketplacePartialErrorVisible));
            OnPropertyChanged(nameof(IsMarketplaceEmptyVisible));
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
            OnPropertyChanged(nameof(IsDiscordImprovedRpcEnabled));
        }
    }

    public bool IsDiscordImprovedRpcEnabled
    {
        get => _isDiscordImprovedRpcEnabled;
        set
        {
            if (_isDiscordImprovedRpcEnabled == value) return;
            _isDiscordImprovedRpcEnabled = value;
            OnPropertyChanged();
            SaveSettings();
            UpdateDiscordPresence();
            OnPropertyChanged(nameof(DiscordRichPresenceStatusText));
        }
    }

    // ── Developer Options ────────────────────────────────────────────────────

    public bool IsDeveloperOptionsVisible
    {
        get => _isDeveloperOptionsVisible;
        set
        {
            if (_isDeveloperOptionsVisible == value) return;
            _isDeveloperOptionsVisible = value;
            OnPropertyChanged();
            SaveSettings();
        }
    }

    /// <summary>
    /// When on, Debug-level traces are also appended to warnings.log (not just
    /// the Debug output), which is useful when diagnosing an issue but noisy
    /// enough that it stays off by default.
    /// </summary>
    public bool IsVerboseLoggingEnabled
    {
        get => _isVerboseLoggingEnabled;
        set
        {
            if (_isVerboseLoggingEnabled == value) return;
            _isVerboseLoggingEnabled = value;
            KodoDiagnostics.VerboseLoggingEnabled = value;
            OnPropertyChanged();
            SaveSettings();
        }
    }

    // Feedback shown under the Developer Options buttons after an action like
    // "Copy Diagnostic Info" or "Clear Logs" completes. Empty until used.
    public string DeveloperOptionsStatusText
    {
        get => _developerOptionsStatusText;
        private set
        {
            if (_developerOptionsStatusText == value) return;
            _developerOptionsStatusText = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(HasDeveloperOptionsStatus));
        }
    }

    public bool HasDeveloperOptionsStatus => !string.IsNullOrWhiteSpace(DeveloperOptionsStatusText);

    /// <summary>Path to crash.log, shown in the button tooltip.</summary>
    public string CrashLogFilePath => KodoDiagnostics.CrashLogFilePath;

    /// <summary>Path to warnings.log, shown in the button tooltip.</summary>
    public string WarningsLogFilePath => KodoDiagnostics.WarningsLogFilePath;

    /// <summary>Path to the folder that contains crash.log, shown in the button tooltip.</summary>
    public string CrashLogFolderPath => KodoDiagnostics.LogDirectoryPath;

    /// <summary>Path to the folder that contains kodosettings.json, shown in the button tooltip.</summary>
    public string SettingsFolderPath =>
        Path.GetDirectoryName(SettingsFilePath) ?? string.Empty;

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

    // Settings toggle: when on, Kodo silently installs marketplace updates for
    // installed extensions on launch and every few hours afterwards, instead of
    // requiring a manual click on "Update" / "Update All" in the Marketplace.
    public bool IsAutoUpdateExtensionsEnabled
    {
        get => _isAutoUpdateExtensionsEnabled;
        set
        {
            if (_isAutoUpdateExtensionsEnabled == value) return;
            _isAutoUpdateExtensionsEnabled = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(AutoUpdateExtensionsStatusText));
            SaveSettings();
            UpdateExtensionAutoUpdateLifecycle();
            if (value)
                _ = AutoUpdateExtensionsIfEnabledAsync();
        }
    }

    // Sub-setting: only takes effect while IsAutoUpdateExtensionsEnabled is on
    // (see the indented checkbox under it in Settings). Suppresses the
    // "Auto-updating X of Y..." progress text in AutoUpdateExtensionsIfEnabledAsync
    // so the silent sweep makes no visible change to the Extensions page at all.
    public bool IsAutoUpdateExtensionsInBackgroundEnabled
    {
        get => _isAutoUpdateExtensionsInBackgroundEnabled;
        set
        {
            if (_isAutoUpdateExtensionsInBackgroundEnabled == value) return;
            _isAutoUpdateExtensionsInBackgroundEnabled = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(AutoUpdateExtensionsStatusText));
            SaveSettings();
        }
    }

    // Settings toggle: when on, Kodo checks GitHub Releases for a newer build
    // a few seconds after launch and offers to download/install it. Turning
    // this off skips the background check entirely - the app stays on the
    // installed version until the user updates manually.
    public bool IsAutoUpdateAppEnabled
    {
        get => _isAutoUpdateAppEnabled;
        set
        {
            if (_isAutoUpdateAppEnabled == value) return;
            _isAutoUpdateAppEnabled = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(AutoUpdateAppStatusText));
            SaveSettings();
            _appUpdateScheduler.UpdateLifecycle();
        }
    }

    // Sub-setting: only takes effect while IsAutoUpdateAppEnabled is on (see
    // the indented checkbox under it in Settings). When on, a found update is
    // downloaded and installed immediately via UpdateService.SilentlyInstallAsync
    // instead of showing UpdateDialog's "Update Now" / "Later" prompt - across
    // the launch check, the periodic timer, and the manual "Check for Updates"
    // button alike.
    public bool IsAutoUpdateAppInBackgroundEnabled
    {
        get => _isAutoUpdateAppInBackgroundEnabled;
        set
        {
            if (_isAutoUpdateAppInBackgroundEnabled == value) return;
            _isAutoUpdateAppInBackgroundEnabled = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(AutoUpdateAppStatusText));
            SaveSettings();
        }
    }

    // Drives the "Check for Updates" button in Settings while a manual check
    // is in flight - disables the button and swaps its label so a slow/rate
    // limited GitHub response doesn't look like a dead click.
    public bool IsCheckingForUpdatesManually
    {
        get => _isCheckingForUpdatesManually;
        private set
        {
            if (_isCheckingForUpdatesManually == value) return;
            _isCheckingForUpdatesManually = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(CheckForUpdatesButtonText));
        }
    }

    public string CheckForUpdatesButtonText => IsCheckingForUpdatesManually ? "Checking…" : "Check for Updates";

    // Result of the most recent manual check ("You're up to date", "vX.Y.Z is
    // available", or a failure message). Empty until the button is clicked.
    public string CheckForUpdatesStatusText
    {
        get => _checkForUpdatesStatusText;
        private set
        {
            if (_checkForUpdatesStatusText == value) return;
            _checkForUpdatesStatusText = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(HasCheckForUpdatesStatus));
        }
    }

    public bool HasCheckForUpdatesStatus => !string.IsNullOrWhiteSpace(CheckForUpdatesStatusText);

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
            RefreshTerminalStatusBindings();
            SaveSettings();
            if (_isTerminalVisible)
                FocusActiveTerminal();
            else
                RefreshTerminalWindows();
        }
    }

    public double TerminalPanelHeight
    {
        get => _terminalPanelHeight;
        set
        {
            var normalized = NormalizeTerminalPanelHeight(value);
            if (_terminalPanelHeight == normalized) return;
            _terminalPanelHeight = normalized;
            OnPropertyChanged();
            SaveSettings();
        }
    }

    // ── Terminal panel resize splitter ───────────────────────────────────────
    // Manual drag handling (rather than Avalonia's GridSplitter) so the dragged
    // value flows straight through the TerminalPanelHeight property - which is
    // what gets persisted to settings.json and is the same value the AXAML
    // binds the panel's Height to. The actual ConPTY/grid resize underneath
    // needs no special handling here: ConsoleTerminal.ArrangeOverride
    // already reacts to any bounds change (this one included) by recalculating
    // rows/cols and calling Resize(), which resizes both the cell buffer and
    // the native pseudo console via ResizePseudoConsole.
    private void TerminalPanelSplitter_OnPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (sender is not InputElement element) return;
        if (!e.GetCurrentPoint(element).Properties.IsLeftButtonPressed) return;

        _isResizingTerminalPanel = true;
        _terminalPanelDragStartPointerY = e.GetPosition(this).Y;
        _terminalPanelDragStartHeight = TerminalPanelHeight;
        e.Pointer.Capture(element);
        e.Handled = true;
    }

    private void TerminalPanelSplitter_OnPointerMoved(object? sender, PointerEventArgs e)
    {
        if (!_isResizingTerminalPanel) return;

        // The splitter sits above the terminal panel, so dragging it up (toward
        // smaller Y) should grow the panel and dragging it down should shrink it -
        // the inverse of the pointer's own delta.
        var deltaY = _terminalPanelDragStartPointerY - e.GetPosition(this).Y;
        TerminalPanelHeight = _terminalPanelDragStartHeight + deltaY;
        e.Handled = true;
    }

    private void TerminalPanelSplitter_OnPointerReleased(object? sender, PointerReleasedEventArgs e)
    {
        if (!_isResizingTerminalPanel) return;

        _isResizingTerminalPanel = false;
        e.Pointer.Capture(null);
        // Flush the final height to disk right away instead of waiting on the
        // debounce timer, so a quick resize-then-close can't lose the change to
        // the same shutdown race the immediate/synchronous save guards against.
        SaveSettings(immediate: true);
        e.Handled = true;
    }

    // Defends against a stray PointerCaptureLost (e.g. dragging the splitter past
    // the window edge, or the window losing focus mid-drag) leaving the panel
    // permanently in "resizing" mode, which would make it keep tracking pointer
    // moves anywhere in the window after the drag should have ended.
    private void TerminalPanelSplitter_OnPointerCaptureLost(object? sender, PointerCaptureLostEventArgs e)
    {
        if (!_isResizingTerminalPanel) return;

        _isResizingTerminalPanel = false;
        SaveSettings(immediate: true);
    }


    public TerminalSession? ActiveTerminalSession
    {
        get => _activeTerminalSession;
        private set
        {
            if (ReferenceEquals(_activeTerminalSession, value))
                return;

            // ── Save the outgoing session's screen buffer ─────────────────────
            // ConsoleTerminal hosts exactly one ConPTY at a time. Switching
            // sessions stops the outgoing process, but the tab remains
            // restorable from its saved snapshot, so show it as paused instead
            // of exited.
            if (_activeTerminalSession is not null)
            {
                if (TerminalHostControl.HasLiveProcess)
                    _activeTerminalSession.Snapshot = TerminalHostControl.SaveSnapshot();
                if (_activeTerminalSession.StatusText is not "Exited" && !_activeTerminalSession.StatusText.StartsWith("Failed", StringComparison.Ordinal))
                    _activeTerminalSession.StatusText = "Paused";
                _activeTerminalSession.IsSelected = false;
            }

            _activeTerminalSession = value;

            if (_activeTerminalSession is not null)
                _activeTerminalSession.IsSelected = true;

            OnPropertyChanged();
            OnPropertyChanged(nameof(HasActiveTerminal));
            OnPropertyChanged(nameof(ActiveTerminalShellDisplayName));
            RefreshTerminalStatusBindings();
            if (_activeTerminalSession is not null)
            {
                // Cold-start the incoming session's process. Start() internally calls
                // ResizeCells which allocates a fresh empty cell grid, so we must
                // restore the snapshot AFTER Start() - not before - otherwise the
                // newly allocated grid would immediately overwrite the restored buffer.
                var shell = AvailableTerminalShells.FirstOrDefault(s =>
                    string.Equals(s.Id, _activeTerminalSession.ShellId, StringComparison.OrdinalIgnoreCase))
                    ?? GetSelectedTerminalShellOrFallback();
                if (shell is not null)
                {
                    try
                    {
                        // Unhook the previous handler BEFORE calling Start(), because
                        // Start() calls Stop() internally which signals the old process
                        // to exit. If the old handler is still subscribed when that
                        // WaitForSingleObject wake-up is posted to the UI thread, it
                        // will fire after the new OnExited is registered and - if the
                        // post races past the guard - will close the brand-new session.
                        if (_activeSessionExitedHandler is not null)
                        {
                            TerminalHostControl.SessionExited -= _activeSessionExitedHandler;
                            _activeSessionExitedHandler = null;
                        }

                        var hasSnapshot = _activeTerminalSession.Snapshot is not null;
                        TerminalHostControl.Start(shell.FileName, shell.Arguments,
                            _activeTerminalSession.WorkingDirectory,
                            suppressOutputUntilRestored: hasSnapshot);
                        _activeTerminalSession.IsRunning = true;
                        _activeTerminalSession.StatusText = "Ready";

                        // Capture the handle that Start() just launched. SessionExited
                        // fires with this same value as its argument, so we can reject
                        // any post whose handle doesn't match - meaning it's a stale
                        // wake-up from a process that Stop() already killed.
                        var expectedHandle = TerminalHostControl.CurrentProcessHandle;

                        var watchedSession = _activeTerminalSession;
                        void OnExited(object? s, IntPtr exitedHandle)
                        {
                            // Reject stale posts: Stop() inside a subsequent Start() kills
                            // the old process, which wakes WaitForSingleObject and queues
                            // a SessionExited post. By the time the UI thread drains it the
                            // new handler is already subscribed. Comparing handles ensures
                            // we only act on the exit of the process we actually started.
                            if (exitedHandle != expectedHandle)
                                return;

                            TerminalHostControl.SessionExited -= OnExited;
                            _activeSessionExitedHandler = null;
                            watchedSession.IsRunning = false;

                            CloseTerminalSession(watchedSession);
                            RefreshTerminalStatusBindings();
                        }
                        _activeSessionExitedHandler = OnExited;
                        TerminalHostControl.SessionExited += OnExited;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[Terminal] Failed to start shell for '{_activeTerminalSession.Title}': {ex.Message}");
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
            OnPropertyChanged(nameof(IsSystemThemeActive));
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
        : "Home";
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

    public string ExplorerHeaderTooltipText => IsFolderOpen
        ? Path.GetFileName(_currentFolderPath!.TrimEnd(Path.DirectorySeparatorChar))
        : "Explorer";

    public string ThemeStatusText => IsSystemThemeActive
        ? $"Current theme: System Default ({CurrentThemeName})"
        : $"Current theme: {CurrentThemeName}";

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
            _selectedWelcomeMessage = null;
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
            _selectedWelcomeMessage = null;
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
            _selectedWelcomeMessage = null;
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
            _selectedWelcomeMessage = null;
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

    // ── Welcome message ──────────────────────────────────────────────────────
    // Holiday/calendar detection, real-world sporting-event theming, and the
    // greeting-pool builder itself all live in WelcomeMessageBuilder.cs.

    // Populated by FetchSportingEventMessagesAsync; passed into
    // WelcomeMessageBuilder.BuildMessages so sporting-event lines can join
    // the greeting pool once they're available.
    private List<string>? _sportingEventMessages;

    /// <summary>
    /// Best-effort fetch of sporting-event-themed welcome messages via
    /// WelcomeMessageBuilder. Invalidates the welcome-message cache only if
    /// the Home screen hasn't shown a greeting yet, to avoid a visible
    /// flicker right after launch.
    /// </summary>
    private async Task FetchSportingEventMessagesAsync()
    {
        var messages = await WelcomeMessageBuilder.FetchSportingEventMessagesAsync(MarketplaceHttpClient);
        if (messages is null) return;

        _sportingEventMessages = messages;

        // Only rebuild the pool if the Home screen hasn't actually shown a
        // greeting yet. If it has, the greeting is already on screen and
        // resetting the cache here would make the WelcomeMessage binding
        // re-roll a *different* random message a moment after launch -
        // a visible flicker from one greeting to another. Leaving the
        // already-chosen message in place means the sporting-themed lines
        // simply join the pool the next time it's genuinely rebuilt
        // (e.g. a personalization setting change, or the next launch).
        if (_selectedWelcomeMessage is null)
        {
            _welcomeMessagesCache = null;
            OnPropertyChanged(nameof(WelcomeMessage));
        }
    }

    // Lazily constructed per-instance so it can incorporate the personalization settings
    // which are read from settings before DataContext is set.
    private string[]? _welcomeMessagesCache;

    // The single message picked for this "generation" of the pool. Cached so that
    // repeated reads of WelcomeMessage (multiple binding evaluations, layout passes,
    // etc.) all return the same text instead of each landing on a different random
    // pick - which is what caused the greeting to visibly flicker right after launch.
    // Reset alongside _welcomeMessagesCache whenever the pool genuinely needs to be
    // re-rolled (e.g. a personalization setting change).
    private string? _selectedWelcomeMessage;

    // Evaluated once per launch: true one in ten thousand times, showing the "Code fast. Stay light" tagline.
    private readonly bool _isTaglineGreeting = Random.Shared.Next(10_000) == 0;
    public bool IsTaglineGreeting => _isTaglineGreeting;

    // ── Birthday ──────────────────────────────────────────────────────────────

    private static readonly DateTime _kodoBirthDate = new(2026, 4, 18);

    /// <summary>
    /// True on April 18 every year- Kodo's birthday. Drives celebratory UI accents
    /// throughout the app (wordmark flourish, window title suffix, status bar note,
    /// and birthday messages in the welcome pool).
    /// </summary>
    public bool IsKodoBirthday
    {
        get
        {
            var today = DateTime.Now;
            return today.Month == _kodoBirthDate.Month && today.Day == _kodoBirthDate.Day;
        }
    }

    /// <summary>How old Kodo is today (whole years since April 18 2026).</summary>
    public int KodoBirthdayAge
    {
        get
        {
            var today = DateTime.Now;
            var age   = today.Year - _kodoBirthDate.Year;
            if (today.Month < _kodoBirthDate.Month ||
                (today.Month == _kodoBirthDate.Month && today.Day < _kodoBirthDate.Day))
                age--;
            return Math.Max(0, age);
        }
    }

    /// <summary>
    /// Short celebratory note shown in the status bar on Kodo's birthday, e.g.
    /// "Kodo turns 1 today! 🎂". Empty on every other day.
    /// </summary>
    public string StatusBarBirthdayText
    {
        get
        {
            if (!IsKodoBirthday) return string.Empty;
            var age = KodoBirthdayAge;
            return age == 1 ? "Kodo turns 1 today! 🎂" : $"Kodo turns {age} today! 🎂";
        }
    }

    /// <summary>True only when StatusBarBirthdayText is non-empty (i.e. on the birthday).</summary>
    public bool IsStatusBarBirthdayVisible => IsKodoBirthday;

    // Evaluated once per launch: true one in ten thousand times, swapping the
    // "Get Started" card's subtitle for a rare alternate line. Same odds as
    // IsTaglineGreeting above, by design - keeps the two easter eggs consistent.
    private readonly bool _isRareGetStartedMessage = Random.Shared.Next(10_000) == 0;
    public string GetStartedSubtitleText => _isRareGetStartedMessage
        ? "Legend has it the code writes itself if you wait long enough. (It doesn't - open a file.)"
        : "Open a file or folder to get started, or create something new.";

    public string WelcomeMessage
    {
        get
        {
            _welcomeMessagesCache ??= WelcomeMessageBuilder.BuildMessages(
                _userName,
                _userCountry,
                _userHemisphere,
                _userTimezoneOffset,
                IsKodoBirthday,
                KodoBirthdayAge,
                _sportingEventMessages);
            _selectedWelcomeMessage ??= _welcomeMessagesCache[Random.Shared.Next(_welcomeMessagesCache.Length)];
            return _selectedWelcomeMessage;
        }
    }

    public bool IsDarkThemeActive  => !IsSystemThemeActive && string.Equals(CurrentThemeName, "Dark",  StringComparison.OrdinalIgnoreCase);
    public bool IsLightThemeActive => !IsSystemThemeActive && string.Equals(CurrentThemeName, "Light", StringComparison.OrdinalIgnoreCase);

    // True when the user picked "follow Windows" rather than an explicit
    // Dark/Light/extension theme. Tracked off _requestedThemeName (not
    // CurrentThemeName) so the System Default blob - not the Dark or Light
    // blob - shows the active ring, even though CurrentThemeName still
    // resolves to a concrete "Dark"/"Light" for the colour-application logic.
    public bool IsSystemThemeActive => string.Equals(_requestedThemeName, "System", StringComparison.OrdinalIgnoreCase);

    // Live preview swatch for the System Default blob. Reflects whatever
    // Windows is currently reporting (Light or Dark) regardless of which
    // theme mode is actually active, refreshed by _windowsThemePollTimer -
    // mirrors how WindowsAccentPreviewBrush stays live for the accent blob.
    public IBrush SystemThemePreviewBackground { get; private set; } = Brush.Parse("#1E1E1E");
    public IBrush SystemThemePreviewBorder     { get; private set; } = Brush.Parse("#2B2B2B");

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
        : IsDiscordImprovedRpcEnabled
            ? "Rich Presence is on (Improved). Shows language, dirty state, page context, and open tab count."
            : "Discord Rich Presence is on when the Discord desktop app is running.";

    public string AutoSaveStatusText =>
        !IsAutoSaveEnabled
            ? "Autosave is turned off."
            : HasFileOpen
                ? "Changes are saved automatically a couple seconds after you stop typing."
                : "Autosave will start working after the file has been saved once.";

    public string AutoUpdateExtensionsStatusText =>
        !IsAutoUpdateExtensionsEnabled
            ? "Extensions only update when you click Update in the Marketplace."
            : AvailableExtensionUpdatesCount > 0
                ? $"Checking periodically and installing updates automatically. {AvailableExtensionUpdatesCount} update{(AvailableExtensionUpdatesCount == 1 ? string.Empty : "s")} pending."
                : IsAutoUpdateExtensionsInBackgroundEnabled
                    ? "Checking periodically and installing new extension updates automatically, without showing progress."
                    : "Checking periodically and installing new extension updates automatically.";

    public string AutoUpdateAppStatusText =>
        !IsAutoUpdateAppEnabled
            ? "Kodo only updates when you download a new installer yourself."
            : IsAutoUpdateAppInBackgroundEnabled
                ? "Checking for new Kodo versions on launch and every few hours, and installing them automatically without asking."
                : "Checking for new Kodo versions on launch and every few hours, and prompting to install them.";

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

    // Black or white - whichever contrasts better against the current AccentBrush.
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
            OnPropertyChanged(nameof(IsAccentTheme));
        }
    }
    public bool IsAccentKodo    => _accentColorMode == "kodo";
    public bool IsAccentWindows => _accentColorMode == "windows";
    public bool IsAccentCustom  => _accentColorMode == "custom";
    // True when the active theme supplies a preset accent colour.
    // When true the "Theme" blob is shown alongside "Kodo", "Windows", and "Custom".
    public bool HasThemeAccent  => _hasThemeAccent;
    // The "Theme" blob is active when the user has explicitly chosen to follow
    // the theme-supplied accent (mode == "theme").
    public bool IsAccentTheme   => _accentColorMode == "theme";
    // Solid-colour preview brush for the Theme blob, always reflects the
    // accent supplied by the currently active extension theme.
    public IBrush ThemeAccentPreviewBrush { get; private set; } = Brush.Parse("#8C00FF");

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
        if (e.PropertyName is nameof(TerminalSession.WorkingDirectory) or nameof(TerminalSession.StatusText) or nameof(TerminalSession.IsRunning) or nameof(TerminalSession.WindowHandle))
        {
            RefreshTerminalStatusBindings();
            RefreshTerminalWindows();
        }
    }

    private void RefreshTerminalStatusBindings()
    {
        OnPropertyChanged(nameof(ActiveTerminalWorkingDirectory));
        OnPropertyChanged(nameof(ActiveTerminalStatusText));
        OnPropertyChanged(nameof(ActiveTerminalFooterText));
        OnPropertyChanged(nameof(TerminalStatusBarText));
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
        Title = BuildWindowTitle();
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
        OnPropertyChanged(nameof(ExplorerHeaderTooltipText));
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
        catch (Exception ex)
        {
            Console.WriteLine($"[Discord] Failed to initialise Rich Presence: {ex.Message}");
            ResetDiscordPresenceForReconnect();
            _discordReconnectTimer.Start();
        }
    }

    private void UpdateDiscordPresence()
    {
        if (_discordRpcClient is null || !IsDiscordRichPresenceEnabled) return;

        // Guard against string allocation on every 75 ms tick by comparing the
        // primitive inputs that drive the presence strings first. Only when
        // something actually changed do we build strings and call SetPresence.
        var currentKey = GetDiscordPresenceKey();
        if (currentKey == _lastDiscordPresenceKey) return;
        _lastDiscordPresenceKey = currentKey;

        try
        {
            var details = GetDiscordPresenceDetails();
            var state   = GetDiscordPresenceState();

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
            _lastDiscordPresenceState   = state;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Discord] Failed to update Rich Presence: {ex.Message}");
            ResetDiscordPresenceForReconnect();
            _discordReconnectTimer.Start();
        }
    }

    // Cheap tuple key built from the primitive fields that drive presence strings.
    // Avoids allocating display strings on every 75 ms refresh tick.
    private (string? filePath, string? folderPath, int tabCount, bool dirty,
             string? language, bool settings, bool extensions, bool home,
             bool improved) GetDiscordPresenceKey() =>
        (_currentFilePath, _currentFolderPath, OpenTabs.Count, _isDirty,
         GetDiscordLanguageLabel(), _isSettingsPageVisible, _isExtensionsPageVisible,
         _isHomePageVisible, _isDiscordImprovedRpcEnabled);

    private string GetDiscordPresenceDetails() =>
        _isDiscordImprovedRpcEnabled
            ? GetDiscordPresenceDetailsImproved()
            : GetDiscordPresenceDetailsClassic();

    private string GetDiscordPresenceState() =>
        _isDiscordImprovedRpcEnabled
            ? GetDiscordPresenceStateImproved()
            : GetDiscordPresenceStateClassic();

    // ── Classic presence (original behaviour) ────────────────────────────────

    private string GetDiscordPresenceDetailsClassic()
    {
        if (HasDocumentOpen) return $"Editing {GetDocumentDisplayName()}";
        return IsFolderOpen ? "Browsing project files" : "Idle in Kodo";
    }

    private string GetDiscordPresenceStateClassic()
    {
        if (HasFileOpen)          return GetDiscordWorkspaceLabel();
        if (_hasUntitledDocument) return GetDiscordWorkspaceLabel("Editing an Unsaved file");
        if (IsFolderOpen)         return GetDiscordWorkspaceLabel();
        return "Waiting for a file";
    }

    // ── Improved presence (experimental) ─────────────────────────────────────

    private string GetDiscordPresenceDetailsImproved()
    {
        if (_isSettingsPageVisible)   return "Tweaking settings";
        if (_isExtensionsPageVisible) return "Browsing extensions";
        if (_isHomePageVisible)       return "On the home screen";
        if (_isTutorialPageVisible)   return "Following the tutorial";
        if (_isWhatsNewPageVisible)   return "Reading what's new";

        if (HasDocumentOpen)
        {
            var fileName = GetDocumentDisplayName();
            var lang     = GetDiscordLanguageLabel();
            var dirty    = _isDirty ? " \u25cf" : string.Empty;
            return string.IsNullOrWhiteSpace(lang)
                ? $"Editing {fileName}{dirty}"
                : $"Editing {fileName}{dirty}  \u00b7  {lang}";
        }

        return IsFolderOpen ? "Browsing project files" : "Idle in Kodo";
    }

    private string? GetDiscordLanguageLabel()
    {
        var extension = CurrentLanguageExtension;
        if (extension is null)
            return null;

        var name = extension.Name
            .Replace("Language Support", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("Support", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Trim();
        if (!string.IsNullOrWhiteSpace(name))
            return name;

        return extension.Id
            .Replace("-kodo-extension", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("-language-support", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("-", " ", StringComparison.OrdinalIgnoreCase)
            .Trim();
    }

    private string GetDiscordPresenceStateImproved()
    {
        if (HasFileOpen)          return GetDiscordWorkspaceLabelImproved();
        if (_hasUntitledDocument) return GetDiscordWorkspaceLabelImproved("Editing an unsaved file");
        if (IsFolderOpen)         return GetDiscordWorkspaceLabelImproved();
        return "Waiting for a file";
    }

    private string GetDiscordWorkspaceLabelImproved(string fallback = "Working in editor")
    {
        if (!IsFolderOpen) return fallback;
        var folderName = Path.GetFileName(_currentFolderPath!.TrimEnd(Path.DirectorySeparatorChar));
        if (string.IsNullOrWhiteSpace(folderName)) return fallback;
        var tabCount = OpenTabs.Count;
        return tabCount > 1
            ? $"Workspace: {folderName}  ({tabCount} files open)"
            : $"Workspace: {folderName}";
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
            _lastDiscordPresenceState   = string.Empty;
            _lastDiscordPresenceKey     = default;
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

            // Cap recursion depth so a deeply-nested or adversarial settings file
            // cannot cause a StackOverflowException inside the deserializer.
            var opts = new JsonSerializerOptions { MaxDepth = 32 };
            var settings = JsonSerializer.Deserialize<AppSettings>(json, opts);
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
            settings.TerminalPanelHeight = NormalizeTerminalPanelHeight(settings.TerminalPanelHeight);
            return settings;
        }
        catch (Exception ex)
        {
            KodoDiagnostics.LogWarning("MainWindow.LoadSettings", ex, operation: $"Failed to load settings from '{SettingsFilePath}'");
            return new AppSettings();
        }
    }

    private void SaveSettings(bool immediate = false, bool synchronous = false)
    {
        if (_suppressSettingsSave) return;

        if (!immediate)
        {
            _settingsSaveDebounceTimer.Stop();
            _settingsSaveDebounceTimer.Start();
            return;
        }

        _settingsSaveDebounceTimer.Stop();
        PersistSettingsSnapshot(BuildSettingsSnapshot(), synchronous);
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
            ThemeName                              = _requestedThemeName,
            AutoSaveEnabled                        = IsAutoSaveEnabled,
            DiscordRichPresenceEnabled             = IsDiscordRichPresenceEnabled,
            DiscordImprovedRpcEnabled              = IsDiscordImprovedRpcEnabled,
            DeveloperOptionsVisible                = IsDeveloperOptionsVisible,
            VerboseLoggingEnabled                   = IsVerboseLoggingEnabled,
            StatusBarFilePathVisible               = IsStatusBarFilePathVisible,
            WordWrapEnabled                        = IsWordWrapEnabled,
            TabSize                                = TabSize,
            EditorFontSize                         = EditorFontSize,
            ConfirmBeforeClosingUnsavedTabsEnabled  = IsConfirmBeforeClosingUnsavedTabsEnabled,
            RestoreOpenTabsOnLaunchEnabled          = IsRestoreOpenTabsOnLaunchEnabled,
            AutoUpdateExtensionsEnabled             = IsAutoUpdateExtensionsEnabled,
            AutoUpdateExtensionsInBackgroundEnabled = IsAutoUpdateExtensionsInBackgroundEnabled,
            AutoUpdateAppEnabled                    = IsAutoUpdateAppEnabled,
            AutoUpdateAppInBackgroundEnabled         = IsAutoUpdateAppInBackgroundEnabled,
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
            LastSeenVersion                         = CurrentAppVersion,
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

    private void PersistSettingsSnapshot(AppSettings snapshot, bool synchronous = false)
    {
        void WriteToDisk()
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
            catch (Exception ex) { KodoDiagnostics.LogWarning("MainWindow.PersistSettingsSnapshot", ex, operation: $"Failed to save settings to '{SettingsFilePath}'"); }
        }

        // The "synchronous" path exists for shutdown: PersistSettingsSnapshot used to
        // always hand the write off to Task.Run, which schedules it on a background
        // thread-pool thread. On window close, Main() returns right after
        // StartWithClassicDesktopLifetime, and the process exits as soon as it does -
        // .NET does not wait for background thread-pool work to finish. That created a
        // race where the very last settings save (the one that matters most) could be
        // dropped silently if the process tore down before the background write landed.
        // Writing synchronously on the UI thread during shutdown guarantees the file is
        // on disk before the app exits; it's a small, fast JSON write so this adds no
        // perceptible delay.
        if (synchronous)
        {
            WriteToDisk();
            return;
        }

        Task.Run(WriteToDisk);
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
        // "System" isn't a real palette - resolve it to whatever Windows is
        // currently reporting before doing the extension/built-in lookup below.
        // _requestedThemeName stays "System" so persistence and the blob's
        // active-ring binding still recognise the user's actual selection.
        var effectiveThemeName = string.Equals(themeName, "System", StringComparison.OrdinalIgnoreCase)
            ? ResolveSystemThemeName()
            : themeName;
        var extensionTheme = ThemeExtensions
            .Select(e => e.ThemeDefinition!)
            .FirstOrDefault(t => string.Equals(t.ThemeId, effectiveThemeName, StringComparison.OrdinalIgnoreCase));

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
            _hasThemeAccent       = true;
            ThemeAccentPreviewBrush = GetCachedBrush(extensionTheme.Accent);
        }
        else
        {
            CurrentThemeName = effectiveThemeName == "Light" ? "Light" : "Dark";
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
            _hasThemeAccent         = false;
            ThemeAccentPreviewBrush = GetCachedBrush("#8C00FF");
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
            "theme"   => _themeAccentHex,
            "windows" => windowsHex,
            "custom"  => _customAccentHex,
            _         => "#8C00FF"   // "kodo" - always the fixed Kodo purple
        };
        try { AccentBrush = GetCachedBrush(resolvedAccent); }
        catch { AccentBrush = GetCachedBrush("#8C00FF"); }
        AccentForegroundBrush = GetAccentForeground(AccentBrush);

        // Same reasoning as WindowsAccentPreviewBrush above: initialise the
        // System Default blob's preview from the live registry value now so
        // it doesn't flash a stale default before the poll timer's first tick.
        RefreshSystemThemePreview();
    }

    private void ApplyTheme(string themeName)
    {
        _requestedThemeName = themeName;
        // "System" isn't a real palette - resolve it to whatever Windows is
        // currently reporting before doing the extension/built-in lookup below.
        // _requestedThemeName stays "System" so persistence and the blob's
        // active-ring binding still recognise the user's actual selection.
        var effectiveThemeName = string.Equals(themeName, "System", StringComparison.OrdinalIgnoreCase)
            ? ResolveSystemThemeName()
            : themeName;
        var extensionTheme = ThemeExtensions
            .Select(e => e.ThemeDefinition!)
            .FirstOrDefault(t => string.Equals(t.ThemeId, effectiveThemeName, StringComparison.OrdinalIgnoreCase));

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
            _hasThemeAccent       = true;
            ThemeAccentPreviewBrush = GetCachedBrush(extensionTheme.Accent);
        }
        else
        {
            CurrentThemeName = effectiveThemeName == "Light" ? "Light" : "Dark";
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
            _hasThemeAccent         = false;
            ThemeAccentPreviewBrush = GetCachedBrush("#8C00FF");
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
        OnPropertyChanged(nameof(HasThemeAccent));
        OnPropertyChanged(nameof(IsAccentKodo));
        OnPropertyChanged(nameof(IsAccentTheme));
        OnPropertyChanged(nameof(ThemeAccentPreviewBrush));
        OnPropertyChanged(nameof(IsSystemThemeActive));
        OnPropertyChanged(nameof(IsDarkThemeActive));
        OnPropertyChanged(nameof(IsLightThemeActive));
        RefreshSystemThemePreview();
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

        // In "kodo" mode, always use the fixed Kodo purple regardless of any active theme.
        if (_accentColorMode == "kodo")
        {
            try { AccentBrush = GetCachedBrush("#8C00FF"); }
            catch { AccentBrush = GetCachedBrush("#8C00FF"); }
            AccentForegroundBrush = GetAccentForeground(AccentBrush);
            OnPropertyChanged(nameof(AccentBrush));
            OnPropertyChanged(nameof(AccentForegroundBrush));
            ApplyThemeToEditor();
            return;
        }

        var hex = _accentColorMode switch
        {
            "theme"   => _themeAccentHex,
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

    // Reads the same registry value Windows itself uses to decide whether
    // apps should render light or dark chrome. Null means "couldn't tell"
    // (non-Windows, locked-down system, etc.), not "light" or "dark".
    private static bool? GetWindowsAppsUseLightTheme()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return null;
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
            if (key?.GetValue("AppsUseLightTheme") is int raw)
                return raw != 0;
        }
        catch { /* Registry unavailable */ }
        return null;
    }

    // Resolves the "System" theme selection to the concrete built-in theme
    // ("Light" or "Dark") that Windows is currently reporting. Falls back to
    // Dark - Kodo's overall default - when the registry value can't be read.
    private static string ResolveSystemThemeName() =>
        GetWindowsAppsUseLightTheme() == true ? "Light" : "Dark";

    // Keeps the System Default blob's preview swatch live, independent of
    // which theme mode is actually active - mirrors how ApplyAccentOverride
    // keeps WindowsAccentPreviewBrush live for the accent blob.
    private void RefreshSystemThemePreview()
    {
        var isLight = GetWindowsAppsUseLightTheme() == true;
        SystemThemePreviewBackground = GetCachedBrush(isLight ? "#FFFFFF" : "#1E1E1E");
        SystemThemePreviewBorder     = GetCachedBrush(isLight ? "#D7DCE5" : "#2B2B2B");
        OnPropertyChanged(nameof(SystemThemePreviewBackground));
        OnPropertyChanged(nameof(SystemThemePreviewBorder));
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

    private void AccentThemeButton_OnClick(object? sender, RoutedEventArgs e)
    {
        AccentColorMode = "theme";
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
        var scrollOffset = EditorTextBox.TextArea.TextView.ScrollOffset;
        ActiveEditorTab.TopLineNumber = EditorTextBox.TextArea.TextView.GetDocumentLineByVisualTop(
            scrollOffset.Y)?.LineNumber ?? 1;
        // Also save the exact pixel offset so restoration is sub-line-accurate.
        ActiveEditorTab.ScrollOffsetY = scrollOffset.Y;
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
        // Restore scroll position. ScrollToLine is applied synchronously first (it works
        // immediately after SetEditorContent on line numbers without a layout pass) so the
        // viewport is already close to correct before the frame is painted. Then we post a
        // precise pixel-offset restore at Background priority - after AvaloniaEdit has
        // completed its own layout - to land exactly where the user left off.
        EditorTextBox.ScrollToLine(tab.TopLineNumber);
        var savedOffsetY = tab.ScrollOffsetY;
        if (savedOffsetY > 0.0)
        {
            // Post at Background priority so AvaloniaEdit finishes its own layout pass
            // before we reposition the viewport. ScrollViewer.Offset is a plain settable
            // Vector property - no interface cast needed, compiles against all Avalonia versions.
            Dispatcher.UIThread.Post(() =>
            {
                var sv = EditorTextBox.GetVisualDescendants().OfType<ScrollViewer>().FirstOrDefault();
                if (sv is not null)
                    sv.Offset = new Vector(sv.Offset.X, savedOffsetY);
            }, DispatcherPriority.Background);
        }
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
            else
            {
                // Offload both the binary-sniff read and the encoding-BOM read to a
                // thread-pool thread so the UI stays responsive on large or slow files.
                var (encoding, corrupted) = await Task.Run(() =>
                {
                    if (IsBinaryContent(path))
                        return (System.Text.Encoding.UTF8, true);
                    return (DetectFileEncoding(path), false);
                });

                isCorrupted = corrupted;
                _currentFileEncoding = encoding;
                content = isCorrupted ? string.Empty : await File.ReadAllTextAsync(path, encoding);
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

        // _suppressSettingsSave was set at the very top of the constructor and has
        // been blocking incidental saves throughout the entire startup sequence.
        // It stays true here while we restore tabs and open the startup file so
        // that each intermediate CollectionChanged / ActiveEditorTab notification
        // doesn't overwrite OpenTabPaths with a partial snapshot.
        // The flag is cleared in a finally block so an exception can never leave
        // it permanently set (which would silently disable all future saves).
        try
        {
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

        _ = RefreshExtensionsAndAutoUpdateAsync();
        _ = RefreshLatestReleaseAsync();
        _ = FetchAnnouncementsAsync();
        _ = FetchSportingEventMessagesAsync();

        // Show the What's New splash once per upgrade. Uses the same version parser
        // as IsNewerVersionAvailable so v-prefixes, -BETA, and -DEV suffixes are all
        // handled correctly. DEV builds suppress the splash (same as the update banner).
        // Not shown on a true first launch (the tutorial takes priority instead).
        if (!_isFirstLaunch && !IsDevBuild && IsCurrentNewerThanLastSeen(_lastSeenVersion))
            IsUpdateSplashVisible = true;

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

        // Note: we deliberately do NOT filter out files whose path happens to fall
        // under a recent folder's directory tree here. AddRecentFile already
        // collapses a file into its parent folder entry at write time, but only
        // when that file was actually opened while that folder was the active
        // project (_currentFolderPath). A file opened standalone (e.g. via
        // File > Open with no folder open, or with a different folder open) that
        // happens to physically live inside some other recent folder's tree is a
        // distinct, legitimate recent entry and must not be dropped on load -
        // doing so previously caused standalone-opened and newly-created files to
        // silently vanish from Recent Files after a restart.
        //
        // We also deliberately do NOT filter by File.Exists / Directory.Exists here.
        // A path that is currently unreachable (USB drive unplugged, network share
        // offline, file temporarily moved) is still a valid recent entry - it should
        // reappear as soon as the path becomes available again. Existence is checked
        // at open time in RecentFileButton_OnClick; entries are only removed when the
        // user explicitly clears them, not automatically on load.
        var entries = (recentFiles ?? [])
            .Where(entry => !string.IsNullOrWhiteSpace(entry.Path))
            .ToList();

        foreach (var entry in entries)
        {
            RecentFiles.Add(new RecentFileItem(entry.Path, entry.IsFolder, entry.LastOpened));
            if (RecentFiles.Count >= MaxRecentFiles)
                break;
        }
    }

    private void AddRecentFile(string? path)
    {
        if (!string.IsNullOrWhiteSpace(path) &&
            !string.IsNullOrWhiteSpace(_currentFolderPath) &&
            IsPathInsideDirectory(path, _currentFolderPath))
        {
            AddRecentFolder(_currentFolderPath);
            return;
        }

        AddRecentPath(path, isFolder: false);
    }

    private void AddRecentFolder(string? path) =>
        AddRecentPath(path, isFolder: true);

    private void AddRecentPath(string? path, bool isFolder)
    {
        if (string.IsNullOrWhiteSpace(path)) return;

        for (var i = RecentFiles.Count - 1; i >= 0; i--)
        {
            var item = RecentFiles[i];
            var isSamePath = string.Equals(item.Path, path, StringComparison.OrdinalIgnoreCase);
            var isChildFileOfFolder = isFolder && !item.IsFolder && IsPathInsideDirectory(item.Path, path);
            if (isSamePath || isChildFileOfFolder)
                RecentFiles.RemoveAt(i);
        }

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

    private void ZoomInButton_OnClick(object? sender, RoutedEventArgs e)  => ZoomImageIn();
    private void ZoomOutButton_OnClick(object? sender, RoutedEventArgs e) => ZoomImageOut();
    private void ZoomResetButton_OnClick(object? sender, RoutedEventArgs e) => ZoomImageReset();

    // ── Image zoom helpers ───────────────────────────────────────────────────

    private void ZoomImageIn()  => ImageZoomLevel = SnapToNiceZoom(_imageZoomLevel + ImageZoomStep);
    private void ZoomImageOut() => ImageZoomLevel = SnapToNiceZoom(_imageZoomLevel - ImageZoomStep);
    private void ZoomImageReset() => ImageZoomLevel = 1.0;

    // Snaps zoom to a clean percentage (0.25, 0.5, 0.75, 1.0, 1.25 …) to
    // avoid floating-point drift making levels like 0.9999999 appear.
    private static double SnapToNiceZoom(double zoom)
    {
        var snapped = Math.Round(zoom / ImageZoomStep) * ImageZoomStep;
        return Math.Clamp(snapped, ImageZoomMin, ImageZoomMax);
    }

    private void ImageScrollViewer_OnPointerWheelChanged(object? sender, PointerWheelEventArgs e)
    {
        if (!HasImagePreview) return;
        var hasControl = (e.KeyModifiers & KeyModifiers.Control) == KeyModifiers.Control;
        if (!hasControl) return;

        // Ctrl+wheel → zoom. Mark handled so the ScrollViewer does NOT also scroll.
        if (e.Delta.Y > 0)
            ZoomImageIn();
        else if (e.Delta.Y < 0)
            ZoomImageOut();

        e.Handled = true;
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

    // Builds the OS-level window title with context-aware page and file state.
    // Mirrors the logic of the improved Discord RPC but kept simpler:
    // page views are labelled plainly, and dirty files get a ● prefix.
    private string BuildWindowTitle()
    {
        var birthday = IsKodoBirthday ? " 🎂" : string.Empty;

        if (_isSettingsPageVisible)   return "Settings";
        if (_isExtensionsPageVisible) return "Extensions";
        if (_isTutorialPageVisible)   return "Tutorial";
        if (_isHomePageVisible)       return $"Kodo{birthday}";

        if (HasDocumentOpen)
        {
            var dirty = _isDirty ? "● " : string.Empty;
            var file  = GetDocumentDisplayName();
            if (IsFolderOpen)
            {
                var workspace = Path.GetFileName(_currentFolderPath!.TrimEnd(Path.DirectorySeparatorChar));
                if (!string.IsNullOrWhiteSpace(workspace))
                    return $"{dirty}{file} - {workspace}";
            }
            return $"{dirty}{file}";
        }

        return $"Kodo{birthday}";
    }

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
                "");
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

        // Do NOT auto-spawn a shell when the panel opens - the user must click
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
        // Mark the session as launching so the UI shows the right status text
        // while ConPTY is starting. The exit watcher is wired by the
        // ActiveTerminalSession setter on every Start() call - subscribing here
        // as well would add a duplicate handler that fires on the wrong session
        // after switch-away/switch-back cycles.
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
        // ConsoleTerminal is a native Avalonia control - it handles its own
        // layout and rendering. Showing / hiding is driven by IsVisible bindings on
        // the host Grid in AXAML, so there is nothing to manually synchronise here.
        // The method is kept so call-sites that still reference it compile cleanly.
    }

    private void FocusActiveTerminal()
    {
        // Post at DispatcherPriority.Background so our Focus() call fires after all
        // pending layout and visibility work has fully settled.
        //
        // When a session is created, HasActiveTerminal flips to true, which makes the
        // placeholder Border IsVisible=false. Avalonia's layout pass on that visibility
        // change moves focus to the window root at DispatcherPriority.Layout. A post
        // at DispatcherPriority.Loaded (lower than Layout) was intended to win that
        // race, but Loaded can still tie with residual layout work, causing the Focus()
        // call to be immediately overwritten. Background is lower than both Layout and
        // Loaded, guaranteeing all layout-driven focus resets have completed first.
        Dispatcher.UIThread.Post(() =>
        {
            if (IsTerminalVisible && ActiveTerminalSession is not null)
            {
                TerminalHostControl.Focus();
            }
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
        {
            if (_activeSessionExitedHandler is not null)
            {
                TerminalHostControl.SessionExited -= _activeSessionExitedHandler;
                _activeSessionExitedHandler = null;
            }
            TerminalHostControl.Stop();
        }

        session.IsRunning = false;
        session.StatusText = "Closed";
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

        // Any deliberate navigation dismisses the update splash.
        if (_isUpdateSplashVisible)
            IsUpdateSplashVisible = false;

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
        {
            CloseTerminalSession(session);
            // The × button has Focusable=False but the click still moves focus away
            // on some Avalonia versions. Explicitly return focus to the terminal so
            // the user can keep typing without having to click the terminal again.
            FocusActiveTerminal();
        }
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
        // Only toggle panel visibility - never clear the folder state.
        // Previously this called CloseFolder() when a project was open, which wiped
        // _currentFolderPath and FileTreeItems; reopening with Ctrl+B then showed an
        // empty sidebar because there was nothing left to repopulate from.
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

    // Used by the tab strip inside the Extensions page - switches the tab and
    // refreshes the marketplace listing (respects the normal cooldown, so rapid
    // tab-switching doesn't spam the GitHub API).
    private void MarketplaceTabButton_OnClick(object? sender, RoutedEventArgs e)
    {
        IsMarketplaceTabSelected = true;
        RefreshMarketplaceConnectivityState();
        _ = RefreshExtensionsDataAsync();
    }

    // Used by the "Visit Marketplace" button on the home screen -
    // opens the Extensions page AND switches to the Marketplace tab
    private void OpenMarketplaceButton_OnClick(object? sender, RoutedEventArgs e)
    {
        OpenExtensionsPage(showMarketplaceTab: true, forceRefresh: true);
    }

    private void RefreshNewsButton_OnClick(object? sender, RoutedEventArgs e) =>
        _ = FetchAnnouncementsAsync();

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

    private void DismissUpdateSplashButton_OnClick(object? sender, RoutedEventArgs e) =>
        IsUpdateSplashVisible = false;

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

    private async void CheckForUpdatesButton_OnClick(object? sender, RoutedEventArgs e) =>
        await CheckForUpdatesManuallyAsync();

    // Explicit, user-initiated update check from the Settings page. Unlike
    // the silent startup check (App.axaml.cs's CheckForUpdatesInBackground),
    // this always reports its result - found, not found, or failed - since
    // the user just asked for one. On finding an update it hands off to the
    // same UpdateDialog the auto-updater uses, so the download/install flow
    // is identical either way.
    private async Task CheckForUpdatesManuallyAsync()
    {
        if (IsCheckingForUpdatesManually) return;

        IsCheckingForUpdatesManually = true;
        CheckForUpdatesStatusText    = "Checking for updates…";

        try
        {
            var installInBackground = IsAutoUpdateAppInBackgroundEnabled;
            var update = await UpdateService.CheckAndHandleUpdateAsync(installInBackground, found =>
            {
                CheckForUpdatesStatusText = installInBackground
                    ? $"Kodo {found.Version} found - installing in the background…"
                    : $"Kodo {found.Version} is available.";
            });

            if (update is null)
                CheckForUpdatesStatusText = $"You're up to date - Kodo {KodoDiagnostics.AppVersion}.";
        }
        catch (Exception ex)
        {
            // CheckAndHandleUpdateAsync already swallows its own failures and
            // returns null, but guard here too so a manual click can never
            // crash the settings page.
            CheckForUpdatesStatusText = "Couldn't check for updates. Check your connection and try again.";
            KodoDiagnostics.LogDebug("Manual check-for-updates failed", ex);
        }
        finally
        {
            IsCheckingForUpdatesManually = false;
        }
    }

    private async void OpenReleasesPageButton_OnClick(object? sender, RoutedEventArgs e)
    {
        var button = sender as Button;
        var originalContent = button?.Content;

        if (button is not null)
        {
            button.IsEnabled = false;
            button.Content   = "Checking…";
        }

        try
        {
            // Mirror the silent startup auto-update flow (App.axaml.cs's
            // CheckForUpdatesInBackground): hit GitHub for the actual
            // downloadable installer asset, then either hand off to
            // UpdateDialog (which downloads it and launches the silent
            // install) or, if "Update automatically without asking" is on,
            // skip the dialog and install silently - instead of just sending
            // the user to the releases page in a browser.
            var update = await UpdateService.CheckAndHandleUpdateAsync(IsAutoUpdateAppInBackgroundEnabled);
            if (update is null)
                // No installer asset could be found (rate-limited, draft
                // release, no .exe attached, etc.) - fall back to the releases
                // page so the user isn't left stuck.
                OpenUrl(ReleasesPageUrl);
        }
        finally
        {
            if (button is not null)
            {
                button.Content   = originalContent;
                button.IsEnabled = true;
            }
        }
    }

    private void OpenDiscordButton_OnClick(object? sender, RoutedEventArgs e) =>
        OpenUrl(DiscordServerUrl);

    private void OpenWebsiteButton_OnClick(object? sender, RoutedEventArgs e) =>
        OpenUrl(WebsiteUrl);

    private void ViewShortcutsButton_OnClick(object? sender, RoutedEventArgs e)
    {
        // Shortcut rows: (gesture, description)
        var shortcuts = new (string Gesture, string Description)[]
        {
            // ── Navigation ────────────────────────────────────────────────────
            ("Ctrl+H",             "Go to Home"),
            ("Ctrl+Shift+E",       "Go to Editor"),
            ("Ctrl+,",             "Open Settings"),
            ("Ctrl+E",             "Open Extensions / Marketplace"),
            // ── Files & tabs ──────────────────────────────────────────────────
            ("Ctrl+N",             "New file"),
            ("Ctrl+O",             "Open file"),
            ("Ctrl+K",             "Open folder  (again to close)"),
            ("Ctrl+S",             "Save"),
            ("Ctrl+Shift+S",       "Save as"),
            ("Ctrl+W",             "Close tab"),
            // ── Editor ────────────────────────────────────────────────────────
            ("Ctrl+F",             "Find in file"),
            ("Ctrl+B",             "Toggle file explorer"),
            ("Ctrl+X / C / V",     "Cut / Copy / Paste"),
            // ── Terminal ──────────────────────────────────────────────────────
            ("Ctrl+`  or  Ctrl+J", "Toggle terminal panel"),
            ("Ctrl+Shift+`",       "New terminal session"),
            // ── Image viewer ──────────────────────────────────────────────────
            ("Ctrl++",             "Zoom in"),
            ("Ctrl+-",             "Zoom out"),
            ("Ctrl+0",             "Reset zoom"),
            // ── General ───────────────────────────────────────────────────────
            ("Escape",             "Close Settings / Extensions / Tutorial / What's New"),
        };

        // ── Header ──────────────────────────────────────────────────────────
        var titleText = new TextBlock
        {
            Text         = "Keyboard Shortcuts",
            FontSize     = 16,
            FontWeight   = FontWeight.SemiBold,
            Foreground   = PrimaryTextBrush,
            TextWrapping = TextWrapping.Wrap,
        };

        // ── Shortcut grid ────────────────────────────────────────────────────
        var grid = new Grid
        {
            ColumnDefinitions = new ColumnDefinitions("Auto,24,*"),
            RowDefinitions    = new RowDefinitions(string.Join(",", Enumerable.Repeat("Auto", shortcuts.Length))),
        };

        for (var i = 0; i < shortcuts.Length; i++)
        {
            var (gesture, description) = shortcuts[i];

            var gestureBorder = new Border
            {
                Background      = CardBrush,
                BorderBrush     = SurfaceBorderBrush,
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(5),
                Padding         = new Thickness(8, 3),
                Margin          = new Thickness(0, 0, 0, 6),
                Child           = new TextBlock
                {
                    Text       = gesture,
                    FontSize   = 12,
                    FontFamily = new FontFamily("Cascadia Code,Consolas,Menlo,monospace"),
                    Foreground = PrimaryTextBrush,
                },
            };

            var descText = new TextBlock
            {
                Text       = description,
                FontSize   = 13,
                Foreground = MutedTextBrush,
                VerticalAlignment = VerticalAlignment.Center,
                Margin     = new Thickness(0, 0, 0, 6),
            };

            Grid.SetRow(gestureBorder, i);
            Grid.SetColumn(gestureBorder, 0);
            Grid.SetRow(descText, i);
            Grid.SetColumn(descText, 2);
            grid.Children.Add(gestureBorder);
            grid.Children.Add(descText);
        }

        var scroll = new ScrollViewer
        {
            Content                       = grid,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
            MaxHeight                     = 460,
        };

        // ── Dismiss button ───────────────────────────────────────────────────
        var dismissButton = new Button
        {
            Content             = "Close",
            HorizontalAlignment = HorizontalAlignment.Right,
            Padding             = new Thickness(20, 8),
            Background          = AccentBrush,
            Foreground          = AccentForegroundBrush,
            BorderThickness     = new Thickness(0),
            CornerRadius        = new CornerRadius(8),
        };

        var content = new StackPanel
        {
            Spacing  = 16,
            Margin   = new Thickness(20),
            Children = { titleText, scroll, dismissButton },
        };

        Window? dialog = null;
        dialog = new Window
        {
            Title                 = "Kodo - Keyboard Shortcuts",
            Width                 = 460,
            SizeToContent         = SizeToContent.Height,
            MinWidth              = 340,
            MaxHeight             = 620,
            CanResize             = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Background            = CardBrush,
            Content               = content,
        };

        dismissButton.Click += (_, _) => dialog!.Close();
        _ = dialog.ShowDialog(this);
    }

    private async void OpenCrashLogFolderButton_OnClick(object? sender, RoutedEventArgs e)
    {
        try
        {
            var path = KodoDiagnostics.LogDirectoryPath;
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            Process.Start(new ProcessStartInfo
            {
                FileName        = path,
                UseShellExecute = true,
            });
        }
        catch (Exception ex)
        {
            await ShowWarningDialogAsync("Open crash logs folder", ex);
        }
    }

    private async void OpenSettingsFolderButton_OnClick(object? sender, RoutedEventArgs e)
    {
        try
        {
            var path = Path.GetDirectoryName(SettingsFilePath) ?? string.Empty;
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            Process.Start(new ProcessStartInfo
            {
                FileName        = path,
                UseShellExecute = true,
            });
        }
        catch (Exception ex)
        {
            await ShowWarningDialogAsync("Open settings folder", ex);
        }
    }

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

    private async void CopyDiagnosticInfoButton_OnClick(object? sender, RoutedEventArgs e)
    {
        try
        {
            var info = new StringBuilder()
                .Append("Kodo ").AppendLine(KodoDiagnostics.AppVersion)
                .Append("OS: ").AppendLine(KodoDiagnostics.OSDescription)
                .Append("Runtime: ").AppendLine(RuntimeInformation.FrameworkDescription)
                .Append("Architecture: ").Append(RuntimeInformation.ProcessArchitecture)
                .Append(" / ").AppendLine(Environment.Is64BitProcess ? "64-bit" : "32-bit")
                .Append("Theme: ").AppendLine(CurrentThemeName)
                .Append("Settings file: ").AppendLine(SettingsFilePath)
                .Append("Crash log: ").AppendLine(CrashLogFilePath)
                .Append("Warnings log: ").AppendLine(WarningsLogFilePath)
                .ToString();

            await (TopLevel.GetTopLevel(this)?.Clipboard?.SetTextAsync(info) ?? Task.CompletedTask);
            DeveloperOptionsStatusText = "Diagnostic info copied to clipboard.";
        }
        catch (Exception ex)
        {
            DeveloperOptionsStatusText = $"Could not copy diagnostic info: {ex.Message}";
        }
    }

    private void ClearLogsButton_OnClick(object? sender, RoutedEventArgs e)
    {
        var clearedAny = false;
        var failures = new List<string>();

        foreach (var path in new[] { KodoDiagnostics.CrashLogFilePath, KodoDiagnostics.WarningsLogFilePath })
        {
            try
            {
                if (!File.Exists(path)) continue;
                File.Delete(path);
                clearedAny = true;
            }
            catch (Exception ex)
            {
                failures.Add($"{Path.GetFileName(path)} ({ex.Message})");
            }
        }

        DeveloperOptionsStatusText = failures.Count > 0
            ? $"Couldn't clear: {string.Join(", ", failures)}"
            : clearedAny
                ? "Logs cleared."
                : "No logs to clear.";
    }

    // Builds a complete, human-readable snapshot of everything Kodo currently
    // knows about this install - settings, personalization, open/recent files,
    // and installed extensions. Backs the "Export Kodo Data" developer option.
    private string BuildKodoDataExport()
    {
        var sb = new StringBuilder();

        void Section(string title)
        {
            sb.AppendLine();
            sb.AppendLine($"── {title} ──");
        }

        sb.AppendLine("Kodo Data Export");
        sb.Append("Generated: ").Append(KodoDiagnostics.UtcNow().ToString("yyyy-MM-dd HH:mm:ss")).AppendLine(" UTC");
        sb.Append("Kodo ").AppendLine(KodoDiagnostics.AppVersion);
        sb.Append("OS: ").AppendLine(KodoDiagnostics.OSDescription);
        sb.Append("Runtime: ").AppendLine(RuntimeInformation.FrameworkDescription);
        sb.Append("Architecture: ").Append(RuntimeInformation.ProcessArchitecture)
          .Append(" / ").AppendLine(Environment.Is64BitProcess ? "64-bit" : "32-bit");

        Section("Appearance");
        sb.AppendLine($"Theme: {(IsSystemThemeActive ? $"System Default ({CurrentThemeName})" : CurrentThemeName)}");
        sb.AppendLine($"Accent mode: {_accentColorMode}");
        sb.AppendLine($"Custom accent colour: {_customAccentHex}");

        Section("Editor");
        sb.AppendLine($"Word wrap: {IsWordWrapEnabled}");
        sb.AppendLine($"Tab size: {TabSize}");
        sb.AppendLine($"Font size: {EditorFontSize}");
        sb.AppendLine($"Confirm before closing unsaved tabs: {IsConfirmBeforeClosingUnsavedTabsEnabled}");
        sb.AppendLine($"Restore open tabs on launch: {IsRestoreOpenTabsOnLaunchEnabled}");
        sb.AppendLine($"Show full file path in status bar: {IsStatusBarFilePathVisible}");
        sb.AppendLine($"Auto-save: {IsAutoSaveEnabled}");

        Section("Updates");
        sb.AppendLine($"Auto-update extensions: {IsAutoUpdateExtensionsEnabled} (silent install: {IsAutoUpdateExtensionsInBackgroundEnabled})");
        sb.AppendLine($"Auto-update Kodo: {IsAutoUpdateAppEnabled} (silent install: {IsAutoUpdateAppInBackgroundEnabled})");
        sb.AppendLine($"Last seen version: {(string.IsNullOrWhiteSpace(_lastSeenVersion) ? "(none)" : _lastSeenVersion)}");
        sb.AppendLine($"Completed tutorial: {_hasCompletedTutorial}");

        Section("Discord Rich Presence");
        sb.AppendLine($"Enabled: {IsDiscordRichPresenceEnabled}");
        sb.AppendLine($"Improved RPC: {IsDiscordImprovedRpcEnabled}");

        Section("Terminal");
        sb.AppendLine($"Preferred shell: {SelectedTerminalShell?.DisplayName ?? "(none)"}");
        sb.AppendLine($"Visible on launch: {IsTerminalVisible}");
        sb.AppendLine($"Panel height: {TerminalPanelHeight:0}px");

        Section("Personalization");
        sb.AppendLine($"Name: {(string.IsNullOrWhiteSpace(_userName) ? "(not set)" : _userName)}");
        sb.AppendLine($"Country: {(string.IsNullOrWhiteSpace(_userCountry) ? "(not set)" : _userCountry)}");
        sb.AppendLine($"Hemisphere: {_userHemisphere}");
        sb.AppendLine($"Timezone offset: {(string.IsNullOrWhiteSpace(_userTimezoneOffset) ? "(not set)" : _userTimezoneOffset)}");

        Section("Developer Options");
        sb.AppendLine($"Developer options visible: {IsDeveloperOptionsVisible}");
        sb.AppendLine($"Verbose logging: {IsVerboseLoggingEnabled}");

        Section("File Locations");
        sb.AppendLine($"Settings file: {SettingsFilePath}");
        sb.AppendLine($"Crash log: {CrashLogFilePath}");
        sb.AppendLine($"Warnings log: {WarningsLogFilePath}");
        sb.AppendLine($"Extensions folder: {ExtensionsFolderPath}");

        var openTabPaths = OpenTabs
            .Where(tab => !tab.IsUntitled && !string.IsNullOrWhiteSpace(tab.Path))
            .Select(tab => tab.Path)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        Section($"Open Tabs ({openTabPaths.Count})");
        if (openTabPaths.Count == 0)
            sb.AppendLine("(none)");
        else
            foreach (var path in openTabPaths)
                sb.AppendLine(path);

        Section($"Recent Files ({RecentFiles.Count})");
        if (RecentFiles.Count == 0)
            sb.AppendLine("(none)");
        else
            foreach (var entry in RecentFiles)
                sb.AppendLine($"{entry.Path}{(entry.IsFolder ? "  [folder]" : string.Empty)}  -  last opened {entry.LastOpened:yyyy-MM-dd HH:mm}");

        var extensions = VisibleLoadedExtensions.ToList();
        Section($"Installed Extensions ({extensions.Count})");
        if (extensions.Count == 0)
            sb.AppendLine("(none)");
        else
            foreach (var ext in extensions)
            {
                sb.AppendLine($"{ext.Name} (v{ext.Version}) - {ext.Type}, by {ext.Author}");
                sb.AppendLine($"    Source: {ext.SourcePath}");
            }

        return sb.ToString();
    }

    private async void ExportKodoDataButton_OnClick(object? sender, RoutedEventArgs e)
    {
        try
        {
            var export = BuildKodoDataExport();
            var suggestedFileName = $"Kodo-Data-Export-{KodoDiagnostics.UtcNow():yyyyMMdd-HHmmss}.txt";

            var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "Export Kodo Data",
                SuggestedFileName = suggestedFileName
            });

            var path = file?.TryGetLocalPath();
            if (string.IsNullOrWhiteSpace(path))
            {
                DeveloperOptionsStatusText = "Export cancelled.";
                return;
            }

            await File.WriteAllTextAsync(path, export);
            DeveloperOptionsStatusText = $"Kodo data exported to {path}.";
        }
        catch (Exception ex)
        {
            KodoDiagnostics.LogWarning("MainWindow.ExportKodoDataButton_OnClick", ex, operation: "Export Kodo data");
            DeveloperOptionsStatusText = $"Could not export Kodo data: {ex.Message}";
        }
    }

    private void ThemeButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Control { Tag: string themeName })
            ApplyTheme(themeName);
    }

    private void ThemeGroupHeader_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Control { DataContext: ThemeExtensionGroup group })
            group.IsExpanded = !group.IsExpanded;
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
            if (!Directory.Exists(item.Path))
            {
                await ShowNotFoundDialogAsync(item.Path, isFolder: true);
                return;
            }
            _currentFolderPath = item.Path;
            AddRecentFolder(item.Path);
            await PopulateFileTreeAsync(item.Path);
            IsFileExplorerVisible = true;
            RefreshState(fullRefresh: true);
        }
        else
        {
            if (!File.Exists(item.Path))
            {
                await ShowNotFoundDialogAsync(item.Path, isFolder: false);
                return;
            }
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
        if (IsUpdatingAllExtensions || _isAutoUpdatingExtensions)
            return;

        var pendingUpdateIds = MarketplaceExtensions
            .Where(extension => extension.IsUpdateAvailable && extension.IsInstallEnabled)
            .Select(extension => extension.Id)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (pendingUpdateIds.Count == 0)
        {
            ExtensionsStatusText = "All extensions are up to date.";
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
                : $"Updated {successfulUpdates} extension{(successfulUpdates == 1 ? string.Empty : "s")}. {failedUpdates} couldn't be updated.";
        }
        finally
        {
            IsUpdatingAllExtensions = false;
        }
    }

    // ── Extension auto-updater ───────────────────────────────────────────────
    // Starts/stops the periodic background check timer to match the current
    // value of IsAutoUpdateExtensionsEnabled. Called once at startup and again
    // every time the setting is toggled in Settings.
    private void UpdateExtensionAutoUpdateLifecycle()
    {
        _extensionAutoUpdateTimer.Stop();
        if (IsAutoUpdateExtensionsEnabled)
            _extensionAutoUpdateTimer.Start();
    }

    // ── App auto-updater ──────────────────────────────────────────────────────
    // The periodic timer, its start/stop lifecycle, and the tick handler that
    // checks-and-handles a found update all live in AppUpdateScheduler
    // (Updater.cs) now - see _appUpdateScheduler above.

    // Fires every few hours while Kodo is open and the setting is enabled, so
    // extensions published mid-session aren't only picked up on next launch.
    private async void ExtensionAutoUpdateTimer_OnTick(object? sender, EventArgs e)
    {
        if (!IsAutoUpdateExtensionsEnabled)
            return;

        // suppressWatchdog: true - this is a silent background check; a stalled
        // network shouldn't pop the "Marketplace refresh" timeout dialog while
        // the user is busy working on something unrelated.
        await RefreshExtensionsDataAsync(force: true, suppressWatchdog: true);
        await AutoUpdateExtensionsIfEnabledAsync();
    }

    // Fires once an hour while Kodo stays open so the Marketplace tab's listing
    // (new extensions, version bumps) stays current during long sessions, even
    // when "Automatically update extensions" is switched off. Unlike
    // ExtensionAutoUpdateTimer_OnTick above, this never installs anything - it
    // only refreshes the data shown in the marketplace.
    private async void MarketplaceRefreshTimer_OnTick(object? sender, EventArgs e)
    {
        // suppressWatchdog: true for the same reason as the extension auto-update
        // sweep - a slow network shouldn't pop a timeout dialog over a silent
        // hourly background refresh.
        await RefreshExtensionsDataAsync(force: true, suppressWatchdog: true);
    }

    // Used on startup: runs the normal extension/marketplace refresh first
    // (which computes IsUpdateAvailable for every installed extension), then
    // silently installs any pending updates if the user has opted in.
    private async Task RefreshExtensionsAndAutoUpdateAsync()
    {
        await RefreshExtensionsDataAsync();
        await AutoUpdateExtensionsIfEnabledAsync();
    }

    // Silently installs every pending marketplace update when the user has
    // opted into "Automatically update extensions" in Settings. Mirrors the
    // manual "Update All" flow above, but is meant to run unattended (startup,
    // periodic timer, or right after the setting is switched on) so it guards
    // against overlapping with itself or with a manual Update All in progress.
    private async Task AutoUpdateExtensionsIfEnabledAsync()
    {
        if (!IsAutoUpdateExtensionsEnabled || _isAutoUpdatingExtensions || IsUpdatingAllExtensions)
            return;

        var pendingUpdateIds = MarketplaceExtensions
            .Where(extension => extension.IsUpdateAvailable && extension.IsInstallEnabled)
            .Select(extension => extension.Id)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (pendingUpdateIds.Count == 0)
            return;

        _isAutoUpdatingExtensions = true;
        var successfulUpdates = 0;
        var failedUpdates = 0;

        try
        {
            for (var index = 0; index < pendingUpdateIds.Count; index++)
            {
                var extensionId = pendingUpdateIds[index];
                var marketplaceExtension = MarketplaceExtensions.FirstOrDefault(entry =>
                    entry.Id.Equals(extensionId, StringComparison.OrdinalIgnoreCase));

                if (marketplaceExtension is null || !marketplaceExtension.IsUpdateAvailable || marketplaceExtension.IsInstalling)
                    continue;

                if (!IsAutoUpdateExtensionsInBackgroundEnabled)
                    ExtensionsStatusText = $"Auto-updating {index + 1} of {pendingUpdateIds.Count}: {marketplaceExtension.Name}...";
                await InstallMarketplaceExtensionAsync(marketplaceExtension);

                var refreshedExtension = MarketplaceExtensions.FirstOrDefault(entry =>
                    entry.Id.Equals(extensionId, StringComparison.OrdinalIgnoreCase));

                if (refreshedExtension is not null && refreshedExtension.IsInstalled && !refreshedExtension.IsUpdateAvailable)
                    successfulUpdates++;
                else
                    failedUpdates++;
            }

            if ((successfulUpdates > 0 || failedUpdates > 0) && !IsAutoUpdateExtensionsInBackgroundEnabled)
            {
                ExtensionsStatusText = failedUpdates == 0
                    ? $"Automatically updated {successfulUpdates} extension{(successfulUpdates == 1 ? string.Empty : "s")}."
                    : $"Automatically updated {successfulUpdates} extension{(successfulUpdates == 1 ? string.Empty : "s")}. {failedUpdates} couldn't be updated.";
            }
        }
        finally
        {
            _isAutoUpdatingExtensions = false;
        }
    }

    private async void UpdateInstalledExtensionButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: LoadedExtension extension })
            return;

        var marketplaceExtension = GetMarketplaceExtensionForInstalled(extension);
        if (marketplaceExtension is null)
        {
            ExtensionsStatusText = $"Couldn't find an update source for {extension.Name}.";
            return;
        }

        await InstallMarketplaceExtensionAsync(marketplaceExtension);
    }

    private async void UninstallExtensionButton_OnClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: LoadedExtension extension })
            await UninstallExtensionAsync(extension);
    }

    // Shows a "Ctrl+click to open link" tooltip when the pointer hovers over a URL,
    // and switches the cursor to a Hand. Only fires tooltip/cursor changes on state
    // transitions (not → over link, over link → not) so Avalonia's tooltip system
    // is not disturbed on every pixel of movement.
    private void EditorTextView_OnPointerMoved(object? sender, PointerEventArgs e)
    {
        var textView = EditorTextBox.TextArea.TextView;
        var nowOverLink = IsPointerOverLink(e.GetPosition(textView), textView);

        if (nowOverLink == _isPointerOverEditorLink)
            return; // no state change - leave tooltip and cursor alone

        _isPointerOverEditorLink = nowOverLink;

        if (nowOverLink)
        {
            ToolTip.SetTip(textView, "Ctrl+click to open link");
            ToolTip.SetShowDelay(textView, 400);
            textView.Cursor = new Cursor(StandardCursorType.Hand);
        }
        else
        {
            ToolTip.SetTip(textView, null);
            textView.Cursor = new Cursor(StandardCursorType.Ibeam);
        }
    }

    private void EditorTextView_OnPointerExited(object? sender, PointerEventArgs e)
    {
        if (!_isPointerOverEditorLink) return;
        _isPointerOverEditorLink = false;
        var textView = EditorTextBox.TextArea.TextView;
        ToolTip.SetTip(textView, null);
        textView.Cursor = new Cursor(StandardCursorType.Ibeam);
    }

    // Returns true if the given pointer position (relative to the TextView, not
    // scroll-adjusted) falls within any URL span on the visible line.
    private bool IsPointerOverLink(Point pointerPosition, AvaloniaEdit.Rendering.TextView textView)
    {
        var pos = textView.GetPositionFloor(pointerPosition + textView.ScrollOffset);
        if (pos is null) return false;

        try
        {
            var line = EditorTextBox.Document.GetLineByNumber(pos.Value.Line);
            var lineText = EditorTextBox.Document.GetText(line.Offset, line.Length);
            var colOffset = pos.Value.Column - 1; // Column is 1-based

            foreach (Match m in StrictLinkElementGenerator.UrlRegex.Matches(lineText))
            {
                var trimmedLength = m.Value.TrimEnd(')', ']', '}', '.', ',', ':', ';', '!', '?', '\'', '"').Length;
                if (colOffset >= m.Index && colOffset < m.Index + trimmedLength)
                    return true;
            }
        }
        catch
        {
            // Document may be null or line out of range during rapid edits - treat as no link
        }

        return false;
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
        // Don't intercept keys destined for the terminal - it handles all input itself.
        // Without this guard the tunnel handler (registered with handledEventsToo: true)
        // fires before ConsoleTerminal.OnKeyDown and marks Enter / Tab / Back /
        // Ctrl+V as handled, so the control never receives them and the terminal freezes.
        //
        // NOTE: We use TopLevel FocusManager instead of e.Source. In a tunnel event
        // e.Source is the element that originated the route, which can still point to the
        // previously-focused control during the first frames after the terminal panel opens
        // (FocusActiveTerminal posts at DispatcherPriority.Input while the tunnel fires
        // synchronously). Reading the focus manager gives the true current focus owner.
        if (IsTerminalVisible && ActiveTerminalSession is not null)
        {
            var focused = TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement() as Visual;
            var isTerminalFocused = focused is not null &&
                (ReferenceEquals(focused, TerminalHostControl) ||
                 focused.GetSelfAndVisualAncestors().Any(v => ReferenceEquals(v, TerminalHostControl)));

            if (isTerminalFocused)
                return;
        }

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
        SaveSettings(immediate: true, synchronous: true);
        _autoSaveTimer.Stop();
        _autoSaveStatusTimer.Stop();
        _discordReconnectTimer.Stop();
        _extensionsRefreshDebounceTimer.Stop();
        _extensionAutoUpdateTimer.Stop();
        _appUpdateScheduler.Stop();
        _marketplaceRefreshTimer.Stop();
        _wordCountRefreshTimer.Stop();
        _settingsSaveDebounceTimer.Stop();
        _windowsAccentPollTimer.Stop();
        _windowsThemePollTimer.Stop();
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

    // Runs on the UI thread every 2 s; mirrors WindowsAccentPollTimer_OnTick.
    // Refreshes the System Default blob's preview whenever Windows' light/dark
    // setting changes, and re-applies the theme too when System mode is the
    // user's active selection so switching Windows' setting takes effect live.
    private void WindowsThemePollTimer_OnTick(object? sender, EventArgs e)
    {
        var current = ResolveSystemThemeName();
        if (current == _lastSeenWindowsThemeName) return;
        _lastSeenWindowsThemeName = current;
        RefreshSystemThemePreview();
        if (IsSystemThemeActive)
            ApplyTheme("System");
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

    // GitHub answers an over-quota request with 403 (anonymous "60 req/hr" limit)
    // or 429 (secondary rate limit, e.g. too many requests in a short burst).
    // EnsureSuccessStatusCode() turns either into an HttpRequestException whose
    // .Message is a generic, unhelpful "Response status code does not indicate
    // success: 403 (Forbidden)." Recognize those two status codes specifically
    // and swap in a tight, actionable message instead of showing that verbatim.
    private static bool IsGitHubRateLimitException(Exception exception) =>
        exception is HttpRequestException { StatusCode: HttpStatusCode.Forbidden or HttpStatusCode.TooManyRequests };

    private static string DescribeFetchFailure(Exception exception) =>
        IsGitHubRateLimitException(exception)
            ? "GitHub's API rate limit was hit. Wait a few minutes, then try again."
            : exception.Message;

    private void RefreshMarketplaceConnectivityState(string? operation = null, Exception? exception = null)
    {
        var hasWirelessConnection = HasActiveWirelessConnection();
        var hasInternetConnection = HasActiveInternetConnection();

        string message = string.Empty;

        if (!hasInternetConnection)
        {
            message = hasWirelessConnection
                ? "No internet connection. Marketplace installs and updates won't work until you're back online."
                : "No Wi-Fi or internet detected. Marketplace installs and updates won't work until you're back online.";
        }
        else if (exception is not null)
        {
            // Show a message for any exception when internet is available - not just
            // known connectivity failures. This covers rate-limit responses, JSON parse
            // errors, unexpected HTTP status codes, and other non-network failures that
            // would otherwise be silent.
            message = IsGitHubRateLimitException(exception)
                ? "GitHub's API rate limit was hit. Marketplace refreshes will resume once it resets."
                : hasWirelessConnection
                    ? "Couldn't reach the marketplace. Your connection may be unstable - try again in a moment."
                    : "No Wi-Fi detected. If you're expecting a connection, reconnect first. Marketplace downloads may fail while offline.";
        }

        if (!string.IsNullOrWhiteSpace(operation) && !string.IsNullOrWhiteSpace(message))
            message = $"{message} (Last issue: {operation})";

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
        // Don't swallow keys when the terminal is focused. ConsoleTerminal.OnKeyDown
        // marks Ctrl+letter and VT sequences as Handled=true, so this handler normally
        // won't see them via the bubble phase - but be explicit as a safety net, and to
        // stop Escape from stealing focus from the terminal.
        if (IsTerminalVisible && ActiveTerminalSession is not null)
        {
            var focused = TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement() as Visual;
            if (focused is not null &&
                (ReferenceEquals(focused, TerminalHostControl) ||
                 focused.GetSelfAndVisualAncestors().Any(v => ReferenceEquals(v, TerminalHostControl))))
                return;
        }

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

            // Image zoom: Ctrl++ / Ctrl+= / Ctrl+NumpadAdd  →  zoom in
            //             Ctrl+-  / Ctrl+NumpadSubtract      →  zoom out
            //             Ctrl+0  / Ctrl+Numpad0             →  reset to 100 %
            case Key.OemPlus:
            case Key.Add:
                if (HasImagePreview)
                {
                    ZoomImageIn();
                    e.Handled = true;
                }
                break;

            case Key.OemMinus:
            case Key.Subtract:
                if (HasImagePreview)
                {
                    ZoomImageOut();
                    e.Handled = true;
                }
                break;

            case Key.D0:
            case Key.NumPad0:
                if (HasImagePreview)
                {
                    ZoomImageReset();
                    e.Handled = true;
                }
                break;
        }
    }

    // ── Nested types ─────────────────────────────────────────────────────────

    private static int NormalizeTabSize(int value) => value is 2 or 4 or 8 ? value : 4;

    // Guards against NaN/Infinity (which Math.Clamp does not reject) coming from a
    // hand-edited or corrupted settings.json, in addition to clamping to the
    // draggable range.
    private static double NormalizeTerminalPanelHeight(double value) =>
        double.IsFinite(value)
            ? Math.Clamp(value, MinTerminalPanelHeight, MaxTerminalPanelHeight)
            : DefaultTerminalPanelHeight;

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
    // Shown when a recent file/folder path is unreachable at open time.
    // Unlike ShowWarningDialogAsync, this is not an error - the path may simply
    // be on a drive that isn't currently connected. The entry is kept in recents
    // so it reappears automatically when the path becomes available again.
    // A "Remove from recents" button is offered as an explicit opt-in to deletion.
    private async Task ShowNotFoundDialogAsync(string path, bool isFolder)
    {
        try
        {
            var kind = isFolder ? "Folder" : "File";

            var titleText = new TextBlock
            {
                Text         = $"{kind} Not Found",
                FontSize     = 16,
                FontWeight   = FontWeight.SemiBold,
                Foreground   = PrimaryTextBrush,
                TextWrapping = TextWrapping.Wrap,
            };

            var bodyText = new TextBlock
            {
                Text         = $"This {kind.ToLowerInvariant()} couldn't be opened because it isn't currently accessible. " +
                               $"It may be on a drive that isn't connected, or it may have been moved or deleted.\n\n{path}",
                FontSize     = 13,
                Foreground   = MutedTextBrush,
                TextWrapping = TextWrapping.Wrap,
            };

            var removeButton = new Button
            {
                Content             = "Remove from Recents",
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
                Content             = "OK",
                HorizontalAlignment = HorizontalAlignment.Right,
                Padding             = new Thickness(28, 8),
                Background          = AccentBrush,
                Foreground          = AccentForegroundBrush,
                BorderThickness     = new Thickness(0),
                CornerRadius        = new CornerRadius(8),
            };

            var buttonRow = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
            buttonRow.Children.Add(removeButton);
            Grid.SetColumn(dismissButton, 1);
            buttonRow.Children.Add(dismissButton);

            var content = new StackPanel
            {
                Spacing  = 12,
                Margin   = new Thickness(20),
                Children = { titleText, bodyText, buttonRow },
            };

            Window? dialog = null;
            dialog = new Window
            {
                Title                 = "Kodo - Not Found",
                Width                 = 480,
                SizeToContent         = SizeToContent.Height,
                MinWidth              = 360,
                MaxHeight             = 400,
                CanResize             = false,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Background            = CardBrush,
                Content               = content,
            };

            removeButton.Click += (_, _) => { RemoveRecentFile(path); dialog!.Close(); };
            dismissButton.Click += (_, _) => dialog!.Close();
            await dialog.ShowDialog(this);
        }
        catch (Exception dialogEx)
        {
            KodoDiagnostics.LogDebug("ShowNotFoundDialogAsync failed to display.", dialogEx);
        }
    }

    // ── Two-tier warning dialog ───────────────────────────────────────────────
    //
    // Critical (isCritical = true):  file-save failures and any operation where
    //   data may be at risk.  Shown with an amber warning banner, logged to
    //   warnings.log, and the title reads "Kodo - Warning".
    //
    // Non-critical (default):  network/marketplace/update failures and other
    //   recoverable errors.  No banner, softer subtitle, same log destination.
    //   Title reads "Kodo - Notice".
    //
    // Both tiers share the same visual structure as the crash dialog (source
    // badge, metadata line, scrollable stack trace, Copy button) so the UI
    // language is consistent across all error surfaces.
    private async Task ShowWarningDialogAsync(string context, Exception exception, bool isCritical = false)
    {
        // Classify automatically: file-save and auto-save failures always
        // get the critical tier since unsaved data may be at risk.
        isCritical = isCritical
            || context.StartsWith("File save", StringComparison.OrdinalIgnoreCase)
            || context.StartsWith("Auto-save", StringComparison.OrdinalIgnoreCase);

        var source = isCritical ? "MainWindow.Warning.Critical" : "MainWindow.Warning";
        KodoDiagnostics.LogWarning(source, exception, operation: context);

        if (ShouldSuppressWarningDialog(context, exception))
        {
            KodoDiagnostics.LogDebug($"Suppressed duplicate warning dialog for '{context}'.", exception);
            return;
        }

        try
        {
            var titleLabel   = isCritical ? "Action required" : "Something went wrong";
            var subtitleMessage = isCritical
                ? "Kodo could not complete this file operation. Your in-editor content is still intact - try saving again or use Save As to choose a different location."
                : "Kodo ran into a problem with this operation. No data was lost - you can try again.";
            var windowTitle  = isCritical ? "Kodo - Warning" : "Kodo - Notice";
            var logPath      = KodoDiagnostics.WarningsLogFilePath;

            // --- Header ---
            var titleText = new TextBlock
            {
                Text         = titleLabel,
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

            // Amber banner - only shown for critical tier so the visual weight
            // matches the severity (mirrors the terminating-crash amber banner).
            var criticalBanner = new Border
            {
                IsVisible       = isCritical,
                Background      = new SolidColorBrush(Color.Parse("#2D1F00")),
                BorderBrush     = new SolidColorBrush(Color.Parse("#6B4800")),
                BorderThickness = new Thickness(1),
                CornerRadius    = new CornerRadius(6),
                Padding         = new Thickness(10, 6),
                Child = new TextBlock
                {
                    Text        = "⚠ This operation affects file data. Check the log if the problem persists.",
                    FontSize    = 12,
                    Foreground  = new SolidColorBrush(Color.Parse("#FFA040")),
                    TextWrapping = TextWrapping.Wrap,
                },
            };

            // Context badge (e.g. "File save", "Extension install - MyLang")
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

            // Human-readable error message above the raw stack trace.
            var errorMessageText = new TextBlock
            {
                Text         = string.IsNullOrWhiteSpace(exception.Message)
                                   ? "An unexpected error occurred."
                                   : DescribeFetchFailure(exception),
                FontSize     = 13,
                Foreground   = PrimaryTextBrush,
                TextWrapping = TextWrapping.Wrap,
            };

            // Scrollable, selectable stack trace.
            var exceptionText = new SelectableTextBlock
            {
                Text         = KodoDiagnostics.BuildDiagnosticPayload(source, exception, false, KodoSeverity.Warning, context),
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
                Text         = $"Log written to: {logPath}",
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
                    criticalBanner,
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
                Title         = windowTitle,
                Width         = 520,
                SizeToContent = SizeToContent.Height,
                MinWidth      = 380,
                MinHeight     = 180,
                MaxHeight     = 660,
                CanResize     = true,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Background    = CardBrush,
                Content       = content,
            };

            copyButton.Click += async (_, _) =>
            {
                try
                {
                    var clip = TopLevel.GetTopLevel(dialog)?.Clipboard;
                    if (clip is not null)
                    {
                        var text = KodoDiagnostics.BuildDiagnosticPayload(source, exception, false, KodoSeverity.Warning, context);
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
            KodoDiagnostics.LogWarning(source, dialogEx, operation: $"Warning dialog failed to display for context '{context}'");
            KodoDiagnostics.LogDebug($"ShowWarningDialogAsync failed to display for context '{context}'.", dialogEx);
        }
    }

    // ── GitHub timeout helper ─────────────────────────────────────────────────
    //
    // Runs <paramref name="factory"/> with a CancellationToken that fires after
    // GitHubOperationTimeout (7 s).  On expiry the task is cancelled and a
    // TimeoutException (with the operation name embedded) is thrown so the
    // caller's existing catch block can route it straight to ShowWarningDialogAsync.
    //
    // Usage:
    //   var result = await RunWithGitHubTimeoutAsync("Marketplace index fetch",
    //       ct => MarketplaceHttpClient.SendAsync(request, ct));
    private static async Task<T> RunWithGitHubTimeoutAsync<T>(
        string operationName,
        Func<CancellationToken, Task<T>> factory)
    {
        using var cts = new CancellationTokenSource(GitHubOperationTimeout);
        try
        {
            return await factory(cts.Token).ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (cts.IsCancellationRequested)
        {
            // Re-raise as TimeoutException so callers can distinguish a
            // deliberate 7-second timeout from a user-initiated cancellation.
            throw new TimeoutException(
                $"GitHub operation '{operationName}' did not complete within " +
                $"{GitHubOperationTimeout.TotalSeconds:0} seconds and was cancelled.");
        }
    }

    // Overload for operations that return no value.
    private static async Task RunWithGitHubTimeoutAsync(
        string operationName,
        Func<CancellationToken, Task> factory)
    {
        await RunWithGitHubTimeoutAsync<bool>(
            operationName,
            async ct => { await factory(ct).ConfigureAwait(false); return true; })
            .ConfigureAwait(false);
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
        public bool DiscordImprovedRpcEnabled  { get; set; }
        public bool DeveloperOptionsVisible { get; set; }
        public bool VerboseLoggingEnabled { get; set; }
        public bool StatusBarFilePathVisible { get; set; } = true;
        public bool WordWrapEnabled { get; set; }
        public int TabSize { get; set; } = 4;
        public int EditorFontSize { get; set; } = 14;
        public bool ConfirmBeforeClosingUnsavedTabsEnabled { get; set; } = true;
        public bool RestoreOpenTabsOnLaunchEnabled { get; set; }
        public bool AutoUpdateExtensionsEnabled { get; set; }
        // Sub-setting under AutoUpdateExtensionsEnabled - see
        // IsAutoUpdateExtensionsInBackgroundEnabled for what it controls.
        public bool AutoUpdateExtensionsInBackgroundEnabled { get; set; }
        // Defaults to true: most users want Kodo to keep itself current
        // without thinking about it, mirroring how the auto-update dialog
        // already behaved before this setting existed.
        public bool AutoUpdateAppEnabled { get; set; } = true;
        // Sub-setting under AutoUpdateAppEnabled - see
        // IsAutoUpdateAppInBackgroundEnabled for what it controls. Defaults to
        // false so the "Update Now" / "Later" prompt still shows unless the
        // user explicitly opts into fully silent installs.
        public bool AutoUpdateAppInBackgroundEnabled { get; set; }
        public string? PreferredTerminalShellId { get; set; }
        public bool TerminalVisible { get; set; }
        public double TerminalPanelHeight { get; set; } = DefaultTerminalPanelHeight;
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
        public string? LastSeenVersion  { get; set; }
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

    public string LastOpenedLongText
    {
        get
        {
            var diff = DateTime.Now - LastOpened;
            if (diff.TotalMinutes < 1)  return "just now";
            if (diff.TotalMinutes < 2)  return "1 minute ago";
            if (diff.TotalHours   < 1)  return $"{(int)diff.TotalMinutes} minutes ago";
            if (diff.TotalHours   < 2)  return "1 hour ago";
            if (diff.TotalDays    < 1)  return $"{(int)diff.TotalHours} hours ago";
            if (diff.TotalDays    < 2)  return "yesterday";
            if (diff.TotalDays    < 7)  return $"{(int)diff.TotalDays} days ago";
            if (diff.TotalDays    < 14) return "1 week ago";
            if (diff.TotalDays    < 30) return $"{(int)(diff.TotalDays / 7)} weeks ago";
            if (diff.TotalDays    < 60) return "1 month ago";
            if (diff.TotalDays    < 365) return $"{(int)(diff.TotalDays / 30)} months ago";
            if (diff.TotalDays    < 730) return "1 year ago";
            return $"{(int)(diff.TotalDays / 365)} years ago";
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

public sealed class EmojiTypefaceColorizer : DocumentColorizingTransformer
{
    private static readonly Typeface EmojiTypeface = new(new FontFamily("Segoe UI Emoji"));
    private static readonly MethodInfo? SetTextRunPropertiesMethod =
        typeof(VisualLineElement).GetMethod("SetTextRunProperties", BindingFlags.Instance | BindingFlags.NonPublic);

    protected override void ColorizeLine(AvaloniaEdit.Document.DocumentLine line)
    {
        var document = CurrentContext.Document;
        if (document is null || line.Length <= 0)
            return;

        var text = document.GetText(line.Offset, line.Length);
        foreach (var (start, end) in EnumerateEmojiRanges(text))
        {
            ChangeLinePart(line.Offset + start, line.Offset + end, element =>
            {
                var properties = element.TextRunProperties.Clone();
                properties.SetTypeface(EmojiTypeface);
                SetTextRunPropertiesMethod?.Invoke(element, [properties]);
            });
        }
    }

    private static IEnumerable<(int Start, int End)> EnumerateEmojiRanges(string text)
    {
        var rangeStart = -1;
        var index = 0;

        while (index < text.Length)
        {
            var codePoint = char.ConvertToUtf32(text, index);
            var length = char.IsSurrogatePair(text, index) ? 2 : 1;
            var isEmoji = IsEmojiCodePoint(codePoint) ||
                          codePoint == 0x200D ||
                          codePoint == 0xFE0F ||
                          codePoint == 0x20E3 ||
                          IsRegionalIndicator(codePoint) ||
                          IsEmojiModifier(codePoint);

            if (isEmoji)
            {
                if (rangeStart < 0)
                    rangeStart = index;
            }
            else if (rangeStart >= 0)
            {
                yield return (rangeStart, index);
                rangeStart = -1;
            }

            index += length;
        }

        if (rangeStart >= 0)
            yield return (rangeStart, text.Length);
    }

    private static bool IsEmojiCodePoint(int codePoint) =>
        codePoint is >= 0x1F000 and <= 0x1FAFF ||
        codePoint is >= 0x2600 and <= 0x27BF ||
        codePoint is >= 0x2300 and <= 0x23FF ||
        codePoint is >= 0x2B00 and <= 0x2BFF;

    private static bool IsRegionalIndicator(int codePoint) =>
        codePoint is >= 0x1F1E6 and <= 0x1F1FF;

    private static bool IsEmojiModifier(int codePoint) =>
        codePoint is >= 0x1F3FB and <= 0x1F3FF;
}

// ── Syntax highlighting ──────────────────────────────────────────────────────
// Builds an AvaloniaEdit IHighlightingDefinition at runtime from the data
// declared in a LoadedExtension (keywords, types, comment markers, color tokens).
public sealed class InterpolatedStringColorizer : DocumentColorizingTransformer
{
    private static readonly MethodInfo? SetTextRunPropertiesMethod =
        typeof(VisualLineElement).GetMethod("SetTextRunProperties", BindingFlags.Instance | BindingFlags.NonPublic);
    private const string VariableIdentifierBodyPattern =
        "[\\p{L}_][\\p{L}\\p{Nd}_]*";
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

    public void UpdateSyntax(CompiledSyntaxProfile? syntaxProfile)
    {
        _rules.Clear();
        _snapshot = null;
        _support = default;

        if (syntaxProfile is null)
        {
            IsEnabled = false;
            _keywordBrush = Brushes.White;
            _punctuationBrush = Brushes.White;
            return;
        }

        IsEnabled = true;
        _keywordBrush = BrushFor(syntaxProfile.Extension, "keyword", "#569CD6");
        _punctuationBrush = BrushFor(syntaxProfile.Extension, "punctuation", "#D4D4D4");
        BuildRules(syntaxProfile);
        _support = BuildSupport(syntaxProfile.Extension);
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

    private void BuildRules(CompiledSyntaxProfile syntaxProfile)
    {
        foreach (var rule in syntaxProfile.TokenRules)
            _rules.Add(new SyntaxBrushRule(rule.Regex, BrushFor(syntaxProfile.Extension, rule.ColorTokenName, rule.FallbackHex)));
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
            if (ch is '(' or '[' or '{')
            {
                ApplyBrush(lineOffset + index, lineOffset + index + 1, GetRainbowBrush(stack.Count));
                stack.Push(ch);
                continue;
            }

            if (TryGetMatchingOpeningBracket(ch, out var opening) &&
                stack.Count > 0 &&
                stack.Peek() == opening)
            {
                ApplyBrush(lineOffset + index, lineOffset + index + 1, GetRainbowBrush(stack.Count - 1));
                stack.Pop();
            }
        }
    }

    private static bool TryGetMatchingOpeningBracket(char ch, out char opening)
    {
        switch (ch)
        {
            case ')':
                opening = '(';
                return true;
            case ']':
                opening = '[';
                return true;
            case '}':
                opening = '{';
                return true;
            default:
                opening = default;
                return false;
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
    private static readonly Regex OrderedListRegex = new(@"^\s{0,3}\d+[.)]\s+", RegexOptions.Compiled);
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
    private static readonly Regex InlineHtmlTagRegex = new(@"</?[\p{L}_][\p{L}\p{Nd}_:-]*(?:\s+[^>\r\n]*)?/?>", RegexOptions.Compiled);
    private static readonly Regex InlineHtmlAttributeStringRegex = new(@"\b[\p{L}_:][\p{L}\p{Nd}_:.-]*\s*=\s*(?<value>""[^""]*""|'[^']*')", RegexOptions.Compiled);
    private static readonly Regex HtmlEmbeddedOpenTagRegex =
        new(@"<(?<tag>script|style)\b(?<attrs>[^>]*)>", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex HtmlEmbeddedTypeAttributeRegex =
        new(@"\btype\s*=\s*(?:""(?<value>[^""]*)""|'(?<value>[^']*)'|(?<value>[^\s>]+))", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private MarkdownSnapshot? _snapshot;
    private Func<string, CompiledSyntaxProfile?>? _languageResolver;
    private Func<string, LoadedExtension?>? _inlineLanguageResolver;
    private readonly Dictionary<string, EmbeddedSyntaxProfile> _embeddedProfileCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, EmbeddedSyntaxProfile?> _htmlEmbeddedProfileCache = new(StringComparer.OrdinalIgnoreCase);
    private IBrush _keywordBrush = Brushes.White;
    private IBrush _typeBrush = Brushes.White;
    private IBrush _stringBrush = Brushes.White;
    private IBrush _commentBrush = Brushes.White;
    private IBrush _operatorBrush = Brushes.White;
    private IBrush _punctuationBrush = Brushes.White;
    private IBrush _variableBrush = Brushes.White;
    private IBrush _mutedBrush = Brushes.White;

    public bool IsEnabled { get; private set; }

    public void UpdateSyntax(LoadedExtension? extension, Func<string, CompiledSyntaxProfile?>? languageResolver, Func<string, LoadedExtension?>? inlineLanguageResolver)
    {
        _snapshot = null;
        _embeddedProfileCache.Clear();
        _htmlEmbeddedProfileCache.Clear();
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
            foreach (var segment in state.HtmlSegments)
                ColorizeHtmlEmbeddedSegment(text, line.Offset, segment);
            return;
        }

        if (state.HtmlSegments.Count > 0)
        {
            ColorizeMarkdownLine(text, line.Offset);
            foreach (var segment in state.HtmlSegments)
                ColorizeHtmlEmbeddedSegment(text, line.Offset, segment);
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
            foreach (var index in AllIndexesOf(text, ':'))
                ApplyBrush(lineOffset, index, index + 1, _punctuationBrush);
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

        foreach (Match tagMatch in InlineHtmlTagRegex.Matches(text))
        {
            foreach (Match attrMatch in InlineHtmlAttributeStringRegex.Matches(tagMatch.Value))
            {
                var valueGroup = attrMatch.Groups["value"];
                if (!valueGroup.Success)
                    continue;

                var start = tagMatch.Index + valueGroup.Index;
                var end = start + valueGroup.Length;
                if (!TryReserveRange(protectedRanges, start, end))
                    continue;

                ApplyBrush(lineOffset, start, end, _stringBrush);
            }
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
            // Unknown/plain language (e.g. ```text, ```plain, unlabelled fence):
            // paint the whole line white so no syntax rules bleed in from the
            // KodoHighlightingDefinition or any other transformer.
            ApplyBrush(lineOffset, 0, text.Length, Brushes.White);
            return;
        }

        fence.Profile.Colorize(
            text,
            lineOffset,
            fence.State,
            (start, end, brush) => ApplyBrush(0, start, end, brush),
            GetRainbowBrush);
    }

    private void ColorizeHtmlEmbeddedSegment(string text, int lineOffset, MarkdownHtmlSegment segment)
    {
        if (segment.Profile is null || segment.End <= segment.Start || segment.Start >= text.Length)
            return;

        var start = Math.Max(0, Math.Min(segment.Start, text.Length));
        var end = Math.Max(start, Math.Min(segment.End, text.Length));
        if (end <= start)
            return;

        segment.Profile.Colorize(
            text[start..end],
            lineOffset + start,
            segment.State,
            (tokenStart, tokenEnd, brush) => ApplyBrush(0, tokenStart, tokenEnd, brush),
            GetRainbowBrush);
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
        MarkdownHtmlBlock? activeHtmlBlock = null;
        var currentState = EmbeddedSyntaxState.Empty;

        foreach (var line in lines)
        {
            var trimmed = line.TrimStart();
            var htmlSegments = activeFence is null
                ? BuildMarkdownHtmlSegments(line, ref activeHtmlBlock)
                : [];

            if (activeFence is null)
            {
                if (TryParseFenceOpening(trimmed, out var opening))
                {
                    states.Add(new MarkdownLineState(null, opening, htmlSegments));
                    activeFence = new FenceState(
                        opening.MarkerChar,
                        opening.MarkerLength,
                        opening.LanguageLabel,
                        ResolveEmbeddedProfile(opening.LanguageLabel),
                        EmbeddedSyntaxState.Empty,
                        null);
                    currentState = EmbeddedSyntaxState.Empty;
                    continue;
                }

                states.Add(new MarkdownLineState(null, null, htmlSegments));
                continue;
            }

            if (TryParseFenceClosing(trimmed, activeFence, out var closing))
            {
                states.Add(new MarkdownLineState(null, closing, []));
                activeFence = null;
                currentState = EmbeddedSyntaxState.Empty;
                continue;
            }

            var fenceHtmlBlock = activeFence.HtmlBlock;
            var fenceHtmlSegments = SupportsMarkdownNestedHtml(activeFence.Profile)
                ? BuildMarkdownHtmlSegments(line, ref fenceHtmlBlock)
                : [];
            var lineFenceState = activeFence with { State = currentState, HtmlBlock = fenceHtmlBlock };
            states.Add(new MarkdownLineState(lineFenceState, null, fenceHtmlSegments));
            activeFence = activeFence with { HtmlBlock = fenceHtmlBlock };
            if (activeFence.Profile is not null)
                currentState = activeFence.Profile.Advance(line, currentState);
        }

        return new MarkdownSnapshot(text, states);
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

        // Try to recognise the snippet's language from installed .kox language
        // extensions and colourise it the same way that language's own files
        // are coloured. Falls back to the flat string colour (previous
        // behaviour) when no installed extension's keywords/types/etc. match.
        var profile = ResolveInlineEmbeddedProfile(content.Value);
        if (profile is null)
        {
            ApplyBrush(lineOffset, content.Index, content.Index + content.Length, _stringBrush);
            return;
        }

        profile.Colorize(
            content.Value,
            lineOffset + content.Index,
            EmbeddedSyntaxState.Empty,
            (start, end, brush) => ApplyBrush(0, start, end, brush),
            GetRainbowBrush);
    }

    private EmbeddedSyntaxProfile? ResolveEmbeddedProfile(string? languageLabel)
    {
        if (string.IsNullOrWhiteSpace(languageLabel) || _languageResolver is null)
            return null;

        var profile = _languageResolver(languageLabel);
        return ResolveEmbeddedProfile(profile);
    }

    private EmbeddedSyntaxProfile? ResolveEmbeddedProfile(CompiledSyntaxProfile? syntaxProfile)
    {
        if (syntaxProfile is null || IsMarkdownExtension(syntaxProfile.Extension))
            return null;

        var cacheKey = $"{syntaxProfile.Extension.Id}|{syntaxProfile.Extension.Version}";
        if (_embeddedProfileCache.TryGetValue(cacheKey, out var cached))
            return cached;

        var profile = EmbeddedSyntaxProfile.Create(syntaxProfile);
        _embeddedProfileCache[cacheKey] = profile;
        return profile;
    }

    // Detects the inline code snippet's language from actual installed .kox
    // language extensions (see InlineCodeLanguageDetector in
    // SyntaxColorEngine.cs for the scoring/heuristics) and resolves it to the
    // same kind of EmbeddedSyntaxProfile that fenced blocks use, so a one-liner
    // like `var x = 5;` is colourised with that language's actual
    // keyword/type/string/comment colours instead of one flat colour.
    private EmbeddedSyntaxProfile? ResolveInlineEmbeddedProfile(string content)
    {
        if (string.IsNullOrWhiteSpace(content) || _inlineLanguageResolver is null)
            return null;

        var extension = _inlineLanguageResolver(content);
        if (extension is null || IsMarkdownExtension(extension))
            return null;

        var cacheKey = $"{extension.Id}|{extension.Version}";
        if (_embeddedProfileCache.TryGetValue(cacheKey, out var cached))
            return cached;

        var profile = EmbeddedSyntaxProfile.Create(CompiledSyntaxProfile.Create(extension));
        _embeddedProfileCache[cacheKey] = profile;
        return profile;
    }

    private List<MarkdownHtmlSegment> BuildMarkdownHtmlSegments(string line, ref MarkdownHtmlBlock? active)
    {
        var segments = new List<MarkdownHtmlSegment>();
        var cursor = 0;

        while (cursor <= line.Length)
        {
            if (active is not null)
            {
                var closeMatch = BuildHtmlCloseTagRegex(active.TagName).Match(line, cursor);
                var segmentEnd = closeMatch.Success ? closeMatch.Index : line.Length;
                if (active.Profile is not null && segmentEnd > cursor)
                {
                    if (EmbeddedTagContent.TryExtract(line, cursor, segmentEnd, active.ContentMode, out var contentStart, out var contentEnd, out var nextContentMode))
                    {
                        var segmentText = line[contentStart..contentEnd];
                        segments.Add(new MarkdownHtmlSegment(contentStart, contentEnd, active.Profile, active.State));
                        active = active with
                        {
                            State = active.Profile.Advance(segmentText, active.State),
                            ContentMode = nextContentMode
                        };
                    }
                    else
                    {
                        active = active with { ContentMode = nextContentMode };
                    }
                }

                if (!closeMatch.Success)
                    break;

                active = null;
                cursor = closeMatch.Index + closeMatch.Length;
                continue;
            }

            var openMatch = HtmlEmbeddedOpenTagRegex.Match(line, cursor);
            if (!openMatch.Success)
                break;

            var tagName = openMatch.Groups["tag"].Value;
            var openEnd = openMatch.Index + openMatch.Length;
            var profile = ResolveMarkdownHtmlEmbeddedProfile(tagName, ExtractHtmlTypeAttribute(openMatch.Groups["attrs"].Value));
            var inlineCloseMatch = BuildHtmlCloseTagRegex(tagName).Match(line, openEnd);

            if (inlineCloseMatch.Success)
            {
                if (profile is not null && inlineCloseMatch.Index > openEnd &&
                    EmbeddedTagContent.TryExtract(line, openEnd, inlineCloseMatch.Index, EmbeddedBlockContentMode.AwaitingContent, out var inlineContentStart, out var inlineContentEnd, out _))
                {
                    segments.Add(new MarkdownHtmlSegment(inlineContentStart, inlineContentEnd, profile, EmbeddedSyntaxState.Empty));
                }

                cursor = inlineCloseMatch.Index + inlineCloseMatch.Length;
                continue;
            }

            active = new MarkdownHtmlBlock(tagName, profile, EmbeddedSyntaxState.Empty, EmbeddedBlockContentMode.AwaitingContent);
            if (profile is not null && openEnd < line.Length)
            {
                if (EmbeddedTagContent.TryExtract(line, openEnd, line.Length, active.ContentMode, out var contentStart, out var contentEnd, out var nextContentMode))
                {
                    var segmentText = line[contentStart..contentEnd];
                    segments.Add(new MarkdownHtmlSegment(contentStart, contentEnd, profile, EmbeddedSyntaxState.Empty));
                    active = active with
                    {
                        State = profile.Advance(segmentText, EmbeddedSyntaxState.Empty),
                        ContentMode = nextContentMode
                    };
                }
                else
                {
                    active = active with { ContentMode = nextContentMode };
                }
            }
            break;
        }

        return segments;
    }

    private EmbeddedSyntaxProfile? ResolveMarkdownHtmlEmbeddedProfile(string tagName, string? typeAttribute)
    {
        if (_languageResolver is null)
            return null;

        var cacheKey = $"{tagName}|{typeAttribute ?? string.Empty}";
        if (_htmlEmbeddedProfileCache.TryGetValue(cacheKey, out var cached))
            return cached;

        var normalizedTag = tagName.Trim().ToLowerInvariant();
        var normalizedType = (typeAttribute ?? string.Empty).Trim();
        string languageLabel;

        if (normalizedTag == "style")
        {
            languageLabel = "css";
        }
        else if (string.IsNullOrWhiteSpace(normalizedType) ||
                 string.Equals(normalizedType, "module", StringComparison.OrdinalIgnoreCase) ||
                 normalizedType.Contains("javascript", StringComparison.OrdinalIgnoreCase) ||
                 normalizedType.Contains("ecmascript", StringComparison.OrdinalIgnoreCase) ||
                 normalizedType.Contains("jscript", StringComparison.OrdinalIgnoreCase))
        {
            languageLabel = "js";
        }
        else if (normalizedType.Contains("typescript", StringComparison.OrdinalIgnoreCase))
        {
            languageLabel = "ts";
        }
        else if (normalizedType.Contains("json", StringComparison.OrdinalIgnoreCase) ||
                 normalizedType.Contains("importmap", StringComparison.OrdinalIgnoreCase))
        {
            languageLabel = "json";
        }
        else if (normalizedType.Contains("css", StringComparison.OrdinalIgnoreCase))
        {
            languageLabel = "css";
        }
        else
        {
            _htmlEmbeddedProfileCache[cacheKey] = null;
            return null;
        }

        var profile = ResolveEmbeddedProfile(_languageResolver(languageLabel));
        _htmlEmbeddedProfileCache[cacheKey] = profile;
        return profile;
    }

    private static string? ExtractHtmlTypeAttribute(string attrs)
    {
        var match = HtmlEmbeddedTypeAttributeRegex.Match(attrs);
        return match.Success ? match.Groups["value"].Value : null;
    }

    private static Regex BuildHtmlCloseTagRegex(string tagName) =>
        new($@"</\s*{Regex.Escape(tagName)}\s*>", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static bool SupportsMarkdownNestedHtml(EmbeddedSyntaxProfile? profile) =>
        profile?.Extension.Extensions.Any(ext =>
            string.Equals(ext, ".html", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(ext, ".htm", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(ext, ".xml", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(ext, ".svg", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(ext, ".xaml", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(ext, ".axaml", StringComparison.OrdinalIgnoreCase)) == true;

    private static IBrush GetRainbowBrush(int depth) => RainbowBrushes[Math.Abs(depth) % RainbowBrushes.Length];

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
        // CommonMark: info string for backtick fences cannot contain backticks.
        if (markerChar == '`' && info.Contains('`'))
            return false;
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
                : new MarkdownLineState(null, null, []);
    }

    private readonly record struct MarkdownLineState(FenceState? ActiveFence, FenceDelimiterInfo? Delimiter, IReadOnlyList<MarkdownHtmlSegment> HtmlSegments);
    private readonly record struct FenceDelimiterInfo(char MarkerChar, int MarkerLength, string? LanguageLabel);
    private sealed record FenceState(char MarkerChar, int MarkerLength, string? LanguageLabel, EmbeddedSyntaxProfile? Profile, EmbeddedSyntaxState State, MarkdownHtmlBlock? HtmlBlock);
    private sealed record MarkdownHtmlSegment(int Start, int End, EmbeddedSyntaxProfile? Profile, EmbeddedSyntaxState State);
    private sealed record MarkdownHtmlBlock(string TagName, EmbeddedSyntaxProfile? Profile, EmbeddedSyntaxState State, EmbeddedBlockContentMode ContentMode);
}

internal sealed class HtmlEmbeddedColorizer : DocumentColorizingTransformer
{
    private static readonly IBrush[] RainbowBrushes =
    [
        Brush.Parse("#FFD700"),
        Brush.Parse("#DA70D6"),
        Brush.Parse("#4FC1FF"),
        Brush.Parse("#C586C0"),
        Brush.Parse("#9CDCFE"),
        Brush.Parse("#D7BA7D")
    ];

    private static readonly Regex OpenTagRegex =
        new(@"<(?<tag>script|style|x:code)\b(?<attrs>[^>]*)>", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex TypeAttributeRegex =
        new(@"\btype\s*=\s*(?:""(?<value>[^""]*)""|'(?<value>[^']*)'|(?<value>[^\s>]+))", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly MethodInfo? SetTextRunPropertiesMethod =
        typeof(VisualLineElement).GetMethod("set_TextRunProperties", BindingFlags.Instance | BindingFlags.NonPublic);

    private HtmlSnapshot? _snapshot;
    private Func<string, string?, CompiledSyntaxProfile?>? _languageResolver;
    private readonly Dictionary<string, EmbeddedSyntaxProfile?> _profileCache = new(StringComparer.OrdinalIgnoreCase);

    public bool IsEnabled { get; private set; }

    public void UpdateSyntax(LoadedExtension? extension, Func<string, string?, CompiledSyntaxProfile?>? languageResolver)
    {
        _snapshot = null;
        _profileCache.Clear();
        _languageResolver = languageResolver;
        IsEnabled = extension is not null && SupportsEmbeddedLanguageTags(extension);
    }

    protected override void ColorizeLine(AvaloniaEdit.Document.DocumentLine line)
    {
        if (!IsEnabled)
            return;

        var document = CurrentContext.Document;
        if (document is null || line.Length <= 0)
            return;

        var text = document.GetText(line.Offset, line.Length);
        var state = EnsureSnapshot(document.Text ?? string.Empty).GetLineState(line.LineNumber);

        foreach (var segment in state.Segments)
        {
            if (segment.Profile is null || segment.End <= segment.Start || segment.Start >= text.Length)
                continue;

            var start = Math.Max(0, Math.Min(segment.Start, text.Length));
            var end = Math.Max(start, Math.Min(segment.End, text.Length));
            if (end <= start)
                continue;

            ColorizeEmbeddedSegment(
                text[start..end],
                line.Offset + start,
                segment.Profile,
                segment.State);
        }
    }

    private static bool TryGetMatchingOpeningBracket(char ch, out char opening)
    {
        switch (ch)
        {
            case ')':
                opening = '(';
                return true;
            case ']':
                opening = '[';
                return true;
            case '}':
                opening = '{';
                return true;
            default:
                opening = default;
                return false;
        }
    }

    private HtmlSnapshot EnsureSnapshot(string text)
    {
        if (_snapshot is { } snapshot && string.Equals(snapshot.Text, text, StringComparison.Ordinal))
            return snapshot;

        _snapshot = BuildSnapshot(text);
        return _snapshot;
    }

    private HtmlSnapshot BuildSnapshot(string text)
    {
        var lines = text.Split(["\r\n", "\n", "\r"], StringSplitOptions.None);
        var states = new List<HtmlLineState>(lines.Length);
        ActiveHtmlBlock? active = null;

        foreach (var line in lines)
            states.Add(BuildLineState(line, ref active));

        return new HtmlSnapshot(text, states);
    }

    private HtmlLineState BuildLineState(string line, ref ActiveHtmlBlock? active)
    {
        var segments = new List<HtmlEmbeddedSegment>();
        var cursor = 0;

        while (cursor <= line.Length)
        {
            if (active is not null)
            {
                var closeMatch = BuildCloseTagRegex(active.TagName).Match(line, cursor);
                var segmentEnd = closeMatch.Success ? closeMatch.Index : line.Length;

                if (active.Profile is not null && segmentEnd > cursor)
                {
                    if (EmbeddedTagContent.TryExtract(line, cursor, segmentEnd, active.ContentMode, out var contentStart, out var contentEnd, out var nextContentMode))
                    {
                        var segmentText = line[contentStart..contentEnd];
                        segments.Add(new HtmlEmbeddedSegment(
                            contentStart,
                            contentEnd,
                            active.Profile,
                            active.State));
                        active = active with
                        {
                            State = active.Profile.Advance(segmentText, active.State),
                            ContentMode = nextContentMode
                        };
                    }
                    else
                    {
                        active = active with { ContentMode = nextContentMode };
                    }
                }

                if (!closeMatch.Success)
                    break;

                active = null;
                cursor = closeMatch.Index + closeMatch.Length;
                continue;
            }

            var openMatch = OpenTagRegex.Match(line, cursor);
            if (!openMatch.Success)
                break;

            var tagName = openMatch.Groups["tag"].Value;
            var attrs = openMatch.Groups["attrs"].Value;
            var openEnd = openMatch.Index + openMatch.Length;
            var profile = ResolveEmbeddedProfile(tagName, ExtractTypeAttribute(attrs));
            var inlineCloseMatch = BuildCloseTagRegex(tagName).Match(line, openEnd);

            if (inlineCloseMatch.Success)
            {
                if (profile is not null && inlineCloseMatch.Index > openEnd &&
                    EmbeddedTagContent.TryExtract(line, openEnd, inlineCloseMatch.Index, EmbeddedBlockContentMode.AwaitingContent, out var inlineContentStart, out var inlineContentEnd, out _))
                {
                    segments.Add(new HtmlEmbeddedSegment(
                        inlineContentStart,
                        inlineContentEnd,
                        profile,
                        EmbeddedSyntaxState.Empty));
                }

                cursor = inlineCloseMatch.Index + inlineCloseMatch.Length;
                continue;
            }

            active = new ActiveHtmlBlock(tagName, profile, EmbeddedSyntaxState.Empty, EmbeddedBlockContentMode.AwaitingContent);
            if (profile is not null && openEnd < line.Length)
            {
                if (EmbeddedTagContent.TryExtract(line, openEnd, line.Length, active.ContentMode, out var contentStart, out var contentEnd, out var nextContentMode))
                {
                    var segmentText = line[contentStart..contentEnd];
                    segments.Add(new HtmlEmbeddedSegment(
                        contentStart,
                        contentEnd,
                        profile,
                        EmbeddedSyntaxState.Empty));
                    active = active with
                    {
                        State = profile.Advance(segmentText, EmbeddedSyntaxState.Empty),
                        ContentMode = nextContentMode
                    };
                }
                else
                {
                    active = active with { ContentMode = nextContentMode };
                }
            }
            break;
        }

        return new HtmlLineState(segments);
    }

    private EmbeddedSyntaxProfile? ResolveEmbeddedProfile(string tagName, string? typeAttribute)
    {
        if (_languageResolver is null)
            return null;

        var cacheKey = $"{tagName}|{typeAttribute ?? string.Empty}";
        if (_profileCache.TryGetValue(cacheKey, out var cached))
            return cached;

        var syntaxProfile = _languageResolver(tagName, typeAttribute);
        if (syntaxProfile is null || SupportsEmbeddedLanguageTags(syntaxProfile.Extension))
        {
            _profileCache[cacheKey] = null;
            return null;
        }

        var profile = EmbeddedSyntaxProfile.Create(syntaxProfile);
        _profileCache[cacheKey] = profile;
        return profile;
    }

    private void ColorizeEmbeddedSegment(string text, int lineOffset, EmbeddedSyntaxProfile profile, EmbeddedSyntaxState state)
    {
        if (string.IsNullOrEmpty(text))
            return;

        profile.Colorize(
            text,
            lineOffset,
            state,
            (start, end, brush) => ApplyBrush(0, start, end, brush),
            GetRainbowBrush);
    }

    private static bool SupportsEmbeddedLanguageTags(LoadedExtension extension) =>
        extension.Extensions.Any(ext =>
            string.Equals(ext, ".html", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(ext, ".htm", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(ext, ".xml", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(ext, ".svg", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(ext, ".xaml", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(ext, ".axaml", StringComparison.OrdinalIgnoreCase));

    private static string? ExtractTypeAttribute(string attrs)
    {
        var match = TypeAttributeRegex.Match(attrs);
        return match.Success ? match.Groups["value"].Value : null;
    }

    private static Regex BuildCloseTagRegex(string tagName) =>
        new($@"</\s*{Regex.Escape(tagName)}\s*>", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static IBrush GetRainbowBrush(int depth) => RainbowBrushes[Math.Abs(depth) % RainbowBrushes.Length];

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

    private sealed record HtmlSnapshot(string Text, IReadOnlyList<HtmlLineState> LineStates)
    {
        public HtmlLineState GetLineState(int lineNumber) =>
            lineNumber > 0 && lineNumber <= LineStates.Count
                ? LineStates[lineNumber - 1]
                : new HtmlLineState([]);
    }

    private sealed record HtmlLineState(IReadOnlyList<HtmlEmbeddedSegment> Segments);
    private sealed record HtmlEmbeddedSegment(int Start, int End, EmbeddedSyntaxProfile? Profile, EmbeddedSyntaxState State);
    private sealed record ActiveHtmlBlock(string TagName, EmbeddedSyntaxProfile? Profile, EmbeddedSyntaxState State, EmbeddedBlockContentMode ContentMode);
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

    public KodoHighlightingDefinition(LoadedExtension ext, CompiledSyntaxProfile syntaxProfile)
    {
        Name = ext.Name;
        _mainRuleSet = BuildRuleSet(ext, syntaxProfile);
    }

    private static HighlightingColor ColorFor(LoadedExtension ext, string tokenName, string fallback)
    {
        var hex = ext.ColorTokens.TryGetValue(tokenName, out var h) ? h : fallback;
        return new HighlightingColor { Foreground = new SimpleHighlightingBrush(Color.Parse(hex)) };
    }

    private static HighlightingRuleSet BuildRuleSet(LoadedExtension ext, CompiledSyntaxProfile syntaxProfile)
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

        if (!isMarkdown)
        {
            foreach (var rule in syntaxProfile.TokenRules)
            {
                codeRuleSet.Rules.Add(new HighlightingRule
                {
                    Regex = rule.Regex,
                    Color = ColorFor(ext, rule.ColorTokenName, rule.FallbackHex)
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
            // Link/image opening bracket only - closing ] ) are left as default text colour
            codeRuleSet.Rules.Add(new HighlightingRule
            {
                Regex = new Regex(@"!?\[", RegexOptions.Compiled),
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
            // Fenced code blocks (``` or ~~~) - added first so they take priority over
            // all inline rules. The emptyRuleSet means no bold/italic/link rules fire
            // inside the fence body, matching what the MarkdownColorizer handles itself.
            mainRuleSet.Spans.Add(new HighlightingSpan
            {
                StartExpression        = new Regex(@"^(?:`{3,}|~{3,})[^\r\n]*$", RegexOptions.Compiled | RegexOptions.Multiline),
                EndExpression          = new Regex(@"^(?:`{3,}|~{3,})\s*$",      RegexOptions.Compiled | RegexOptions.Multiline),
                SpanColor              = new HighlightingColor { Foreground = new SimpleHighlightingBrush(Color.Parse("#F4F4F4")) },
                SpanColorIncludesStart = true,
                SpanColorIncludesEnd   = true,
                RuleSet                = emptyRuleSet
            });

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
        if (!isMarkdown && !string.IsNullOrEmpty(ext.CommentLine))
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
            // $@"..."/@$"...": interpolated verbatim. Escaping is via a
            // doubled "" (handled by the (?!"") lookahead below), never via
            // backslash - isVerbatim:true so a literal trailing backslash
            // (e.g. $@"C:\Users\") doesn't get mistaken for an escape and
            // swallow the rest of the file.
            mainRuleSet.Spans.Add(CreateRegexStringSpan(@"(?:\$@|@\$)""", @"""(?!"")", stringColor, emptyRuleSet, allowEndOfLineFallback: false, isVerbatim: true));
            mainRuleSet.Spans.Add(CreateRegexStringSpan(@"\$""", @"""", stringColor, emptyRuleSet, allowEndOfLineFallback: true));

            // Bare verbatim string (@"..."): no backslash escaping at all, can
            // span multiple lines, and a literal quote is written as a
            // doubled "" rather than \". There was previously no standalone-
            // file equivalent of this at all - a verbatim string in a .cs
            // file fell through to the ordinary backslash-escaped '"' span
            // below, so e.g. @"C:\Users\" (a literal trailing backslash) would
            // be misread as an escaped delimiter and the string would never
            // close. This mirrors TryMatchCSharpStringPrefix's "@\"" case in
            // EmbeddedSyntaxProfile so a verbatim string looks the same in a
            // .cs file as it does inside a fenced ```cs block.
            mainRuleSet.Spans.Add(CreateRegexStringSpan(@"@""", @"""(?!"")", stringColor, emptyRuleSet, allowEndOfLineFallback: false, isVerbatim: true));
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
        bool allowEndOfLineFallback,
        bool isVerbatim = false)
    {
        // For backslash-escaped strings, a closing delimiter only actually
        // closes the string when it's preceded by an *even* number of
        // backslashes - an odd run means the last backslash escapes the
        // delimiter itself (e.g. \"), not the other way around. The old
        // "(?<!\\)" guard only ever looked one character back, so it got this
        // wrong for any string ending in an escaped backslash followed by the
        // delimiter - e.g. a perfectly ordinary "C:\\Users\\" path literal -
        // and would keep scanning past the real end of the string. This
        // mirrors the correct counting done by IsEscaped/IsStringTerminator in
        // EmbeddedSyntaxProfile (SyntaxColorEngine.cs) so a string literal
        // looks the same whether it's in a standalone file or inside a fenced
        // code block / embedded <script>/<style> tag.
        //
        // Verbatim-style strings (e.g. C#'s @"...") don't use backslash
        // escapes at all - a doubled delimiter is the escape, which the
        // caller already encodes as a lookahead in endDelimiterPattern - so
        // no backslash guard is applied for those; a backslash right before
        // the closing delimiter is just a literal character.
        var unescapedDelimiterGuard = isVerbatim ? string.Empty : @"(?<=(?:^|[^\\])(?:\\\\)*)";
        var endPattern = allowEndOfLineFallback
            ? $@"{unescapedDelimiterGuard}{endDelimiterPattern}|$"
            : $@"{unescapedDelimiterGuard}{endDelimiterPattern}";

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

// Replaces AvaloniaEdit's default LinkElementGenerator with one that:
// 1. Only matches genuine http:// / https:// URLs (not bare words or www. paths)
// 2. Trims trailing punctuation so ')' etc. at the end of a URL stay white
// 3. Requires Ctrl+click to open (RequireControlModifierForClick = true)
public sealed class StrictLinkElementGenerator : LinkElementGenerator
{
    private static readonly char[] TrailingPunctuation = [')', ']', '}', '.', ',', ':', ';', '!', '?', '\'', '"'];
    private static readonly Regex StrictUrlRegex = new(
        @"https?://[^\s<>""'\[\](){}\|\\^`]+",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // Exposed so the hover-tooltip handler can reuse the same regex without
    // duplicating the pattern.
    internal static Regex UrlRegex => StrictUrlRegex;

    public StrictLinkElementGenerator()
    {
        RequireControlModifierForClick = true;
    }

    public override VisualLineElement? ConstructElement(int offset)
    {
        var line = CurrentContext.VisualLine;
        var document = CurrentContext.Document;
        var lineText = document.GetText(line.FirstDocumentLine.Offset, line.FirstDocumentLine.Length);
        var relativeOffset = offset - line.FirstDocumentLine.Offset;

        var match = StrictUrlRegex.Match(lineText, relativeOffset);
        if (!match.Success || match.Index != relativeOffset)
            return null;

        var url = match.Value.TrimEnd(TrailingPunctuation);
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri))
            return null;

        var linkText = new VisualLineLinkText(line, url.Length);
        linkText.NavigateUri = uri;
        linkText.RequireControlModifierForClick = RequireControlModifierForClick;
        return linkText;
    }
}

// Converts bool → FontWeight for the release-notes inline run template.
// True  → SemiBold (bold spans from **…** or __…__)
// False → Regular  (normal text)
public sealed class BoolToFontWeightConverter : Avalonia.Data.Converters.IValueConverter
{
    public static readonly BoolToFontWeightConverter Instance = new();

    public object Convert(object? value, Type targetType, object? parameter, System.Globalization.CultureInfo culture) =>
        value is true ? FontWeight.SemiBold : FontWeight.Regular;

    public object ConvertBack(object? value, Type targetType, object? parameter, System.Globalization.CultureInfo culture) =>
        throw new NotSupportedException();
}
public sealed class NewsItem
{
    public string Title     { get; init; } = string.Empty;
    public string Body      { get; init; } = string.Empty;
    public string UpdatedAt { get; init; } = string.Empty;
    public bool HasTitle     => !string.IsNullOrWhiteSpace(Title);
    public bool HasBody      => !string.IsNullOrWhiteSpace(Body);
    public bool HasUpdatedAt => !string.IsNullOrWhiteSpace(UpdatedAt);
}