// Licensed under GPL-v3.0
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;

namespace Kodo;

// Result of an update check: nothing newer, or a downloadable release found.
internal sealed record UpdateInfo(
    string Version,        // e.g. "v1.2.0" (raw tag_name from GitHub)
    string ReleaseNotesUrl,
    string AssetDownloadUrl,
    string AssetName,
    long AssetSizeBytes);

// Reports download progress back to the UI (0.0 - 1.0, plus a human label).
internal sealed record UpdateDownloadProgress(double Fraction, string Label);

// Checks GitHub Releases, downloads the installer, and launches it silently. Best-effort only.
internal static class UpdateService
{
    // Repo that publishes Kodo releases. Update if the repo ever moves.
    private const string RepoOwner = "KerbalMissile";
    private const string RepoName  = "Kodo";

    private const string LatestReleaseUrl =
        $"https://api.github.com/repos/{RepoOwner}/{RepoName}/releases/latest";

    private static readonly HttpClient Http = CreateHttpClient();

    private static HttpClient CreateHttpClient()
    {
        var client = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(15),
        };
        // GitHub's API requires a User-Agent header on every request.
        client.DefaultRequestHeaders.UserAgent.ParseAdd("Kodo-Updater");
        client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
        return client;
    }

    // Update check

    // Hits GitHub's "latest release" endpoint and compares tags; null if none, unreachable, or no installer asset.
    public static async Task<UpdateInfo?> CheckForUpdateAsync(CancellationToken ct = default)
    {
        try
        {
            using var response = await Http.GetAsync(LatestReleaseUrl, ct).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
                return null;

            await using var stream = await response.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
            var release = await JsonSerializer.DeserializeAsync<GitHubRelease>(stream, JsonOptions, ct)
                .ConfigureAwait(false);

            if (release is null || string.IsNullOrWhiteSpace(release.TagName))
                return null;

            if (release.Draft || release.Prerelease)
                return null;

            if (!IsNewerVersion(release.TagName, KodoDiagnostics.AppVersion))
                return null;

            // Inno Setup output is a plain .exe - grab the first .exe asset.
            var asset = release.Assets?.FirstOrDefault(a =>
                a.Name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase));

            if (asset is null)
                return null;

            return new UpdateInfo(
                Version: release.TagName,
                ReleaseNotesUrl: release.HtmlUrl ?? $"https://github.com/{RepoOwner}/{RepoName}/releases",
                AssetDownloadUrl: asset.BrowserDownloadUrl,
                AssetName: asset.Name,
                AssetSizeBytes: asset.Size);
        }
        catch
        {
            // Non-critical - swallow failures and report "no update".
            return null;
        }
    }

    // Compares two vX.Y.Z tags; true if remote is newer. Falls back to string inequality for unparseable formats.
    internal static bool IsNewerVersion(string remote, string local)
    {
        var remoteParts = ParseVersionParts(remote);
        var localParts  = ParseVersionParts(local);

        if (remoteParts is null || localParts is null)
            return !string.Equals(remote, local, StringComparison.OrdinalIgnoreCase);

        for (var i = 0; i < Math.Max(remoteParts.Length, localParts.Length); i++)
        {
            var r = i < remoteParts.Length ? remoteParts[i] : 0;
            var l = i < localParts.Length ? localParts[i] : 0;
            if (r != l) return r > l;
        }

        return false;
    }

    private static int[]? ParseVersionParts(string tag)
    {
        // Strips leading "v" and trailing "-DEV"/"-beta"/build metadata.
        var core = tag.Trim();
        if (core.Length > 0 && (core[0] == 'v' || core[0] == 'V'))
            core = core[1..];

        var dashIndex = core.IndexOf('-');
        if (dashIndex >= 0) core = core[..dashIndex];
        var plusIndex = core.IndexOf('+');
        if (plusIndex >= 0) core = core[..plusIndex];

        var segments = core.Split('.');
        var parts = new int[segments.Length];
        for (var i = 0; i < segments.Length; i++)
        {
            if (!int.TryParse(segments[i], out parts[i]))
                return null;
        }

        return parts.Length > 0 ? parts : null;
    }

    // Settings

    // Reads the auto-update toggle directly from kodosettings.json, standalone, before a MainWindow exists. Defaults to true.
    public static bool IsAutoUpdateEnabledInSettings()
    {
        try
        {
            var path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Kodo",
                "kodosettings.json");

            if (!File.Exists(path)) return true;

            var json = File.ReadAllText(path);
            if (string.IsNullOrWhiteSpace(json)) return true;

            var settings = JsonSerializer.Deserialize<AutoUpdateSettings>(json);
            return settings?.AutoUpdateAppEnabled ?? true;
        }
        catch
        {
            return true;
        }
    }

    // Same standalone read, for "Update Kodo in the background"; defaults to false.
    public static bool IsAutoUpdateInBackgroundEnabledInSettings()
    {
        try
        {
            var path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Kodo",
                "kodosettings.json");

            if (!File.Exists(path)) return false;

            var json = File.ReadAllText(path);
            if (string.IsNullOrWhiteSpace(json)) return false;

            var settings = JsonSerializer.Deserialize<AutoUpdateSettings>(json);
            return settings?.AutoUpdateAppInBackgroundEnabled ?? false;
        }
        catch
        {
            return false;
        }
    }

    // Minimal subset of MainWindow's AppSettings needed to read this one flag.
    private sealed class AutoUpdateSettings
    {
        public bool AutoUpdateAppEnabled { get; set; } = true;
        public bool AutoUpdateAppInBackgroundEnabled { get; set; }
    }

    // Download

    // Downloads the installer to a temp path with progress reporting.
    public static async Task<string> DownloadInstallerAsync(
        UpdateInfo update,
        IProgress<UpdateDownloadProgress>? progress,
        CancellationToken ct = default)
    {
        var destinationDir = Path.Combine(Path.GetTempPath(), "Kodo-Update");
        Directory.CreateDirectory(destinationDir);

        var destinationPath = Path.Combine(destinationDir, update.AssetName);

        using var response = await Http.GetAsync(
            update.AssetDownloadUrl,
            HttpCompletionOption.ResponseHeadersRead,
            ct).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();

        var totalBytes = response.Content.Headers.ContentLength ?? update.AssetSizeBytes;

        await using var httpStream = await response.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
        await using var fileStream = new FileStream(
            destinationPath, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize: 81920, useAsync: true);

        var buffer = new byte[81920];
        long readTotal = 0;
        int read;
        while ((read = await httpStream.ReadAsync(buffer, ct).ConfigureAwait(false)) > 0)
        {
            await fileStream.WriteAsync(buffer.AsMemory(0, read), ct).ConfigureAwait(false);
            readTotal += read;

            if (progress is not null && totalBytes > 0)
            {
                var fraction = Math.Clamp((double)readTotal / totalBytes, 0, 1);
                var label = $"{FormatBytes(readTotal)} / {FormatBytes(totalBytes)}";
                progress.Report(new UpdateDownloadProgress(fraction, label));
            }
        }

        return destinationPath;
    }

    private static string FormatBytes(long bytes)
    {
        const double mb = 1024 * 1024;
        return bytes >= mb
            ? $"{bytes / mb:0.#} MB"
            : $"{bytes / 1024.0:0} KB";
    }

    // Install / restart

    // Launches the installer silently, then exits so it can overwrite Kodo.exe.
    // Process.Kill() below is a fallback if CloseApplicationsFilter isn't in the .iss.
    public static void LaunchInstallerAndExit(string installerPath)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName        = installerPath,
            // No /SKIPIFSILENT - the .iss's "Launch Kodo" entry needs a silent run to fire.
            Arguments       = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /CLOSEAPPLICATIONS /RESTARTAPPLICATIONS",
            UseShellExecute = true,
        };

        Process.Start(startInfo);

        // Gives the installer time to start and Restart Manager to register this process - 500ms was too tight.
        Thread.Sleep(1500);

        Environment.Exit(0);
    }

    // Downloads and installs with no UI, used for "Update Kodo in the background"; mirrors UpdateDialog's download-then-launch.
    public static async Task SilentlyInstallAsync(UpdateInfo update, CancellationToken ct = default)
    {
        try
        {
            var installerPath = await DownloadInstallerAsync(update, progress: null, ct).ConfigureAwait(false);
            LaunchInstallerAndExit(installerPath);
        }
        catch (Exception ex)
        {
            // Best-effort like the rest of the pipeline - log and move on; the user is picked up on the next check.
            KodoDiagnostics.WriteDiagnosticLog(
                source: "UpdateService.SilentlyInstallAsync",
                exception: ex,
                isTerminating: false,
                severity: "Warning",
                operation: "AutoUpdate");
        }
    }

    // Consolidated check-then-act flow, replacing duplicated check/install/dialog branches at each call site.
    public static async Task<UpdateInfo?> CheckAndHandleUpdateAsync(
        bool installInBackground,
        Action<UpdateInfo>? onUpdateFound = null,
        CancellationToken ct = default)
    {
        var update = await CheckForUpdateAsync(ct).ConfigureAwait(false);
        if (update is null) return null;

        onUpdateFound?.Invoke(update);

        if (installInBackground)
            await SilentlyInstallAsync(update, ct).ConfigureAwait(false);
        else
            UpdateDialog.ShowFor(update);

        return update;
    }

    // GitHub API DTOs

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    private sealed class GitHubRelease
    {
        [JsonPropertyName("tag_name")]
        public string TagName { get; set; } = "";

        [JsonPropertyName("html_url")]
        public string? HtmlUrl { get; set; }

        [JsonPropertyName("draft")]
        public bool Draft { get; set; }

        [JsonPropertyName("prerelease")]
        public bool Prerelease { get; set; }

        [JsonPropertyName("assets")]
        public GitHubAsset[]? Assets { get; set; }
    }

    private sealed class GitHubAsset
    {
        [JsonPropertyName("name")]
        public string Name { get; set; } = "";

        [JsonPropertyName("browser_download_url")]
        public string BrowserDownloadUrl { get; set; } = "";

        [JsonPropertyName("size")]
        public long Size { get; set; }
    }
}

// Update dialog UI: self-contained, built in code. Flow: Update available -> Update Now -> progress -> installer launch.
internal sealed class UpdateDialog : Window
{
    // Same dark-surface palette as every other Kodo dialog, for visual consistency.
    private static readonly Color SurfaceColor   = DialogPalette.Surface;
    private static readonly Color SurfaceDeep    = DialogPalette.SurfaceDeep;
    private static readonly Color BorderColor    = DialogPalette.Border;
    private static readonly Color BadgeBgColor   = DialogPalette.BadgeBg;
    private static readonly Color TextMutedColor = DialogPalette.TextMuted;
    private static readonly Color TextDimColor   = DialogPalette.TextDim;

    // Resolved from the user's active accent setting, same as MainWindow.
    private readonly Color _accentColor;
    private readonly Color _accentForeground;

    private readonly UpdateInfo _update;
    private readonly string? _preDownloadedInstallerPath;
    private readonly TextBlock _statusText;
    private readonly ProgressBar _progressBar;
    private readonly Button _primaryButton;
    private readonly Button _laterButton;
    private readonly StackPanel _content;

    // preDownloadedInstallerPath means KodoUpdater already fetched the installer; Update Now skips straight to launching it.
    public UpdateDialog(UpdateInfo update, string? preDownloadedInstallerPath = null)
    {
        _update = update;
        _preDownloadedInstallerPath = preDownloadedInstallerPath;
        (_accentColor, _accentForeground) = AccentResolver.GetCurrentAccent();

        Title  = "Kodo - Update Available";
        Width  = 460;
        SizeToContent = SizeToContent.Height;
        CanResize  = false;
        Background = new SolidColorBrush(SurfaceColor);
        WindowStartupLocation = WindowStartupLocation.CenterScreen;

        var iconBadge = new Border
        {
            Background      = new SolidColorBrush(_accentColor),
            CornerRadius    = new CornerRadius(8),
            Width           = 40,
            Height          = 40,
            Child = new TextBlock
            {
                Text                = "↑",
                FontSize            = 20,
                Foreground          = new SolidColorBrush(_accentForeground),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment   = VerticalAlignment.Center,
            },
        };

        var titleText = new TextBlock
        {
            Text       = $"Kodo {update.Version} is available",
            FontSize   = 16,
            FontWeight = Avalonia.Media.FontWeight.SemiBold,
            Foreground = Brushes.White,
            TextWrapping = TextWrapping.Wrap,
            VerticalAlignment = VerticalAlignment.Center,
        };

        var headerRow = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Spacing     = 12,
            Children    = { iconBadge, titleText },
        };

        _statusText = new TextBlock
        {
            Text         = preDownloadedInstallerPath is not null
                ? "A new version of Kodo has already been downloaded and is ready to install."
                : "A new version of Kodo has been published. Update now to get the latest fixes and features.",
            FontSize     = 13,
            Foreground   = new SolidColorBrush(TextMutedColor),
            TextWrapping = TextWrapping.Wrap,
        };

        var notesLink = new TextBlock
        {
            Text         = "View release notes",
            FontSize     = 12,
            Foreground   = new SolidColorBrush(Color.Parse("#9CDCFE")),
            Cursor       = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand),
        };
        notesLink.PointerPressed += (_, _) => OpenUrl(update.ReleaseNotesUrl);

        _progressBar = new ProgressBar
        {
            Minimum    = 0,
            Maximum    = 1,
            Value      = 0,
            Height     = 6,
            IsVisible  = false,
            Foreground = new SolidColorBrush(_accentColor),
            Background = new SolidColorBrush(BadgeBgColor),
            CornerRadius = new CornerRadius(3),
        };

        _laterButton = new Button
        {
            Content             = "Later",
            HorizontalAlignment = HorizontalAlignment.Left,
            Padding             = new Thickness(16, 8),
            Background          = new SolidColorBrush(BadgeBgColor),
            Foreground          = new SolidColorBrush(TextMutedColor),
            BorderBrush         = new SolidColorBrush(BorderColor),
            BorderThickness     = new Thickness(1),
            CornerRadius        = new CornerRadius(8),
        };
        _laterButton.Click += (_, _) => Close();

        _primaryButton = new Button
        {
            Content             = "Update Now",
            HorizontalAlignment = HorizontalAlignment.Right,
            Padding             = new Thickness(20, 8),
            Background          = new SolidColorBrush(_accentColor),
            Foreground          = new SolidColorBrush(_accentForeground),
            BorderThickness     = new Thickness(0),
            CornerRadius        = new CornerRadius(8),
        };
        _primaryButton.Click += async (_, _) => await BeginUpdateAsync();

        var buttonRow = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
        buttonRow.Children.Add(_laterButton);
        Grid.SetColumn(_primaryButton, 1);
        buttonRow.Children.Add(_primaryButton);

        _content = new StackPanel
        {
            Spacing = 14,
            Margin  = new Thickness(22),
            Children =
            {
                headerRow,
                _statusText,
                notesLink,
                _progressBar,
                buttonRow,
            },
        };

        Content = _content;
    }

    // Shows non-modally if no owner can be safely used, modal otherwise - mirrors the crash dialog's owner-safety check.
    public static void ShowFor(UpdateInfo update, string? installerPath = null)
    {
        Dispatcher.UIThread.Post(() =>
        {
            var dialog = new UpdateDialog(update, installerPath);

            Window? owner = null;
            if (Application.Current?.ApplicationLifetime
                is Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop)
            {
                var main = desktop.MainWindow;
                if (main is { IsVisible: true })
                    owner = main;
            }

            if (owner is not null)
                dialog.Show(owner);
            else
                dialog.Show();
        });
    }

    private async Task BeginUpdateAsync()
    {
        _primaryButton.IsEnabled = false;
        _laterButton.IsEnabled   = false;

        // Fast path: sentinel file means it's already downloaded.
        if (_preDownloadedInstallerPath is not null && File.Exists(_preDownloadedInstallerPath))
        {
            _primaryButton.Content = "Restarting…";
            _statusText.Text       = "Restarting Kodo to finish installing…";
            await Task.Delay(400);
            UpdateService.LaunchInstallerAndExit(_preDownloadedInstallerPath);
            return;
        }

        _primaryButton.Content   = "Downloading…";
        _progressBar.IsVisible   = true;
        _statusText.Text         = "Downloading the update…";

        var progress = new Progress<UpdateDownloadProgress>(p =>
        {
            _progressBar.Value = p.Fraction;
            _statusText.Text   = $"Downloading… {p.Label}";
        });

        try
        {
            var installerPath = await UpdateService.DownloadInstallerAsync(_update, progress);

            _statusText.Text       = "Update downloaded. Restarting Kodo to finish installing…";
            _primaryButton.Content = "Restarting…";

            // Brief pause so the "downloaded" message is actually visible.
            await Task.Delay(600);

            UpdateService.LaunchInstallerAndExit(installerPath);
        }
        catch (Exception ex)
        {
            _statusText.Text         = "The update couldn't be downloaded. Check your connection and try again.";
            _primaryButton.Content   = "Retry";
            _primaryButton.IsEnabled = true;
            _laterButton.IsEnabled   = true;
            _progressBar.IsVisible   = false;

            KodoDiagnostics.WriteDiagnosticLog(
                source: "UpdateDialog.BeginUpdateAsync",
                exception: ex,
                isTerminating: false,
                severity: "Warning",
                operation: "AutoUpdate");
        }
    }

    private static void OpenUrl(string url)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName        = url,
                UseShellExecute = true,
            });
        }
        catch
        {
            // Opening the browser is a convenience action; never let it crash the dialog.
        }
    }
}
// Single source of truth for the dark-surface colours used by every code-built dialog.
internal static class DialogPalette
{
    public static readonly Color Surface     = Color.Parse("#1E1E1E");
    public static readonly Color SurfaceDeep = Color.Parse("#1A1A1A");
    public static readonly Color Border      = Color.Parse("#3A3A3A");
    public static readonly Color BadgeBg     = Color.Parse("#2B2B2B");
    public static readonly Color TextMuted   = Color.Parse("#A0A0A0");
    public static readonly Color TextDim     = Color.Parse("#606060");
    public static readonly Color TokenBlue   = Color.Parse("#9CDCFE");  // source badge
    public static readonly Color TokenOrange = Color.Parse("#CE9178");  // stack trace
}

// Resolves accent colour for dialogs shown before/independently of MainWindow.
// "Theme" mode reads the last-cached theme accent hex from settings, since standalone
// dialogs can't load the extension system to resolve it themselves.
internal static class AccentResolver
{
    private const string DefaultAccentHex = "#8C00FF";
    private const string SettingsFileName = "kodosettings.json";

    // Cached for the process lifetime - dialogs are short-lived and the accent rarely changes mid-session.
    public static (Color Accent, Color Foreground) GetCurrentAccent()
    {
        var hex = ResolveAccentHex();
        Color accent;
        try { accent = Color.Parse(hex); }
        catch { accent = Color.Parse(DefaultAccentHex); }

        return (accent, GetAccentForeground(accent));
    }

    private static string ResolveAccentHex()
    {
        var settings = LoadAccentSettings();

        return settings.AccentColorMode switch
        {
            // Uses the accent hex cached from the last time MainWindow resolved the
            // active extension theme; falls back to Kodo purple if none was ever cached.
            "theme"   => string.IsNullOrWhiteSpace(settings.CachedThemeAccentHex)
                ? DefaultAccentHex : settings.CachedThemeAccentHex,
            "windows" => GetWindowsAccentColor() ?? "#0078D4",
            "custom"  => string.IsNullOrWhiteSpace(settings.CustomAccentHex)
                ? DefaultAccentHex : settings.CustomAccentHex,
            _         => DefaultAccentHex, // "kodo" (and any unrecognised value)
        };
    }

    private static AccentSettings LoadAccentSettings()
    {
        try
        {
            var path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Kodo",
                SettingsFileName);

            if (!File.Exists(path)) return new AccentSettings();

            var json = File.ReadAllText(path);
            if (string.IsNullOrWhiteSpace(json)) return new AccentSettings();

            return JsonSerializer.Deserialize<AccentSettings>(json) ?? new AccentSettings();
        }
        catch
        {
            // Failure just falls back to default Kodo purple.
            return new AccentSettings();
        }
    }

    [SupportedOSPlatform("windows")]
    private static string? GetWindowsAccentColorWindows()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Explorer\Accent");
            if (key?.GetValue("AccentColorMenu") is int raw)
            {
                // AccentColorMenu is stored as AABBGGRR.
                var r = raw & 0xFF;
                var g = (raw >> 8) & 0xFF;
                var b = (raw >> 16) & 0xFF;
                return $"#{r:X2}{g:X2}{b:X2}";
            }
        }
        catch { /* Registry unavailable */ }
        return null;
    }

    private static string? GetWindowsAccentColor()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return null;
        return GetWindowsAccentColorWindows();
    }

    // Returns White or Black, whichever contrasts better against the accent (WCAG luminance).
    private static Color GetAccentForeground(Color accent)
    {
        static double Lin(byte channel)
        {
            var s = channel / 255.0;
            return s <= 0.04045 ? s / 12.92 : Math.Pow((s + 0.055) / 1.055, 2.4);
        }

        var luminance = 0.2126 * Lin(accent.R) + 0.7152 * Lin(accent.G) + 0.0722 * Lin(accent.B);
        var whiteContrast = 1.05 / (luminance + 0.05);
        var blackContrast = (luminance + 0.05) / 0.05;

        return whiteContrast >= blackContrast ? Colors.White : Colors.Black;
    }

    // Minimal AppSettings subset needed to resolve the accent, kept separate from MainWindow.
    private sealed class AccentSettings
    {
        public string AccentColorMode { get; set; } = "kodo";
        public string CustomAccentHex { get; set; } = DefaultAccentHex;
        public string? CachedThemeAccentHex { get; set; }
    }
}

// Reads KodoUpdater's sentinel file, left after its independent 6-hour GitHub poll.
internal static class PendingUpdateService
{
    private static string FilePath => Path.Combine(Path.GetTempPath(), "Kodo-Update", "pending.json");

    // Returns the pending update only if still newer and the installer file still exists; otherwise cleans up the stale sentinel.
    public static (string Version, string InstallerPath)? TryGetPendingUpdate()
    {
        try
        {
            if (!File.Exists(FilePath)) return null;

            var json = File.ReadAllText(FilePath);
            var record = JsonSerializer.Deserialize<PendingUpdateRecord>(json);
            if (record is null) { Clear(); return null; }

            if (!File.Exists(record.InstallerPath))
            {
                Clear();
                return null;
            }

            if (!UpdateService.IsNewerVersion(record.Version, KodoDiagnostics.AppVersion))
            {
                // Already on this version or newer - the download is stale.
                Clear();
                return null;
            }

            return (record.Version, record.InstallerPath);
        }
        catch
        {
            return null;
        }
    }

    public static void Clear()
    {
        try { File.Delete(FilePath); } catch { /* best-effort cleanup */ }
    }

    private sealed record PendingUpdateRecord(string Version, string InstallerPath, DateTime DownloadedAtUtc);
}
// Owns the timer that checks for a new build every six hours while the app is open.
internal sealed class AppUpdateScheduler
{
    private readonly DispatcherTimer _timer = new() { Interval = TimeSpan.FromHours(6) };
    private readonly Func<bool> _isEnabled;
    private readonly Func<bool> _isManualCheckInProgress;
    private readonly Func<bool> _installInBackground;
    public AppUpdateScheduler(Func<bool> isEnabled, Func<bool> isManualCheckInProgress, Func<bool> installInBackground)
    {
        _isEnabled = isEnabled;
        _isManualCheckInProgress = isManualCheckInProgress;
        _installInBackground = installInBackground;
        _timer.Tick += async (_, _) => await OnTickAsync().ConfigureAwait(true);
    }
    public void UpdateLifecycle()
    {
        _timer.Stop();
        if (_isEnabled())
            _timer.Start();
    }
    public void Stop() => _timer.Stop();

    // Fires every six hours while enabled; skips while a manual check is in flight.
    private async Task OnTickAsync()
    {
        if (!_isEnabled() || _isManualCheckInProgress())
            return;

        try
        {
            await UpdateService.CheckAndHandleUpdateAsync(_installInBackground()).ConfigureAwait(true);
        }
        catch (Exception ex)
        {
            // Silent background check - this must never surface as a crash.
            KodoDiagnostics.LogDebug("Periodic app update check failed", ex);
        }
    }
}