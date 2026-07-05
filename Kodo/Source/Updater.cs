// Licensed under the Kodo Public License v1.1
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

// Result of an update check: either nothing newer was found, or a downloadable
// release was found with everything needed to fetch + launch the installer.
internal sealed record UpdateInfo(
    string Version,        // e.g. "v1.2.0" (raw tag_name from GitHub)
    string ReleaseNotesUrl,
    string AssetDownloadUrl,
    string AssetName,
    long AssetSizeBytes);

// Reports download progress back to the UI (0.0 - 1.0, plus a human label).
internal sealed record UpdateDownloadProgress(double Fraction, string Label);

// Checks GitHub Releases for a newer Kodo build, downloads the installer asset,
// and launches it with silent/auto-restart flags so the update applies with
// minimal friction. The whole flow is best-effort: any failure here must never
// crash or block the app, since update-checking is a convenience feature.
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
        // GitHub's API requires a User-Agent header on every request, and rejecting
        // requests without one is its way of throttling anonymous/unidentified callers.
        client.DefaultRequestHeaders.UserAgent.ParseAdd("Kodo-Updater");
        client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
        return client;
    }

    // ── Update check ──────────────────────────────────────────────────────────

    // Hits the GitHub "latest release" endpoint and compares its tag against the
    // currently running version. Returns null if there's no update, the check
    // failed (offline, rate-limited, etc.), or no installer asset could be found.
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

            // Look for the Windows installer asset. Inno Setup output is a plain
            // .exe, so we just grab the first .exe asset attached to the release.
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
            // Network failures, rate limiting, malformed JSON, etc. - update
            // checking is non-critical, so swallow and report "no update".
            return null;
        }
    }

    // Compares two "vX.Y.Z"-style tags. Returns true if `remote` is strictly
    // newer than `local`. Falls back to a simple string-inequality check if
    // either string doesn't parse as a dotted version, so unexpected tag
    // formats (e.g. "v1.1.0-DEV") never crash the comparison - they're just
    // treated as "different, but not necessarily newer" unless segments parse.
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
        // Strip a leading "v"/"V" and any trailing "-DEV"/"-beta"/build metadata
        // after the numeric core, e.g. "v1.2.0-DEV" -> "1.2.0".
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

    // ── Settings ──────────────────────────────────────────────────────────────

    // Reads the "Automatically check for Kodo updates" toggle
    // directly from kodosettings.json. Mirrors AccentResolver's approach below:
    // a tiny, best-effort, standalone read so App.axaml.cs's startup check can
    // honour the setting without needing a MainWindow/ViewModel instance to
    // exist yet. Defaults to true (matching AppSettings.AutoUpdateAppEnabled's
    // own default) if the file is missing, unreadable, or doesn't have the key.
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

    // Same best-effort standalone read as IsAutoUpdateEnabledInSettings, for the
    // "Update Kodo in the background" sub-setting. Defaults to false
    // (matching AppSettings.AutoUpdateAppInBackgroundEnabled's own default) so
    // App.axaml.cs's launch-time check still shows the UpdateDialog prompt
    // unless the user has explicitly opted into fully silent installs.
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

    // ── Download ──────────────────────────────────────────────────────────────

    // Downloads the installer asset to a temp path, reporting progress along
    // the way. Returns the full path to the downloaded file.
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

    // ── Install / restart ────────────────────────────────────────────────────

    // Launches the downloaded installer with Inno Setup's silent + auto-restart
    // flags, then immediately exits the current process so the installer can
    // overwrite the running Kodo.exe.
    //
    //   /VERYSILENT          - no UI at all (no wizard pages, no progress window)
    //   /SUPPRESSMSGBOXES    - suppress any message boxes Inno would otherwise show
    //   /NORESTART           - never reboot the machine, even if a file is locked
    //   /CLOSEAPPLICATIONS   - automatically close apps using files the installer
    //                          needs to replace (requires AppMutex/CloseApplications
    //                          to be set in the .iss script - if it isn't, this flag
    //                          is simply ignored rather than failing)
    //   /RESTARTAPPLICATIONS - relaunch Kodo automatically once setup finishes
    //
    // If the .iss script doesn't have CloseApplicationsFilter configured, the
    // explicit Process.GetCurrentProcess().Kill() fallback below still ensures the
    // running exe is unlocked before the installer tries to write to it.
    public static void LaunchInstallerAndExit(string installerPath)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName        = installerPath,
            // NOTE: no /SKIPIFSILENT. KodoInstaller.iss's [Run] "Launch Kodo"
            // entry is what actually relaunches the app post-install; that
            // entry is itself conditioned on a silent run, so /SKIPIFSILENT
            // was suppressing the one thing that reliably restarts Kodo.
            // /RESTARTAPPLICATIONS alone isn't sufficient because it depends
            // on Restart Manager having already recorded this process, which
            // needs more lead time than we were giving it (see Thread.Sleep
            // below) before Environment.Exit(0) tears it down.
            Arguments       = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /CLOSEAPPLICATIONS /RESTARTAPPLICATIONS",
            UseShellExecute = true,
        };

        Process.Start(startInfo);

        // Give the installer enough time to spin up AND for Restart Manager
        // to register this process before we exit. 500ms was too tight - the
        // installer's RM session would sometimes query process state after
        // we'd already vanished, so it never saw Kodo as "needs restarting"
        // and the post-update relaunch silently failed.
        Thread.Sleep(1500);

        Environment.Exit(0);
    }

    // Downloads and installs an update with no UI at all - no dialog, no
    // "Update Now" prompt, no progress bar. Used in place of UpdateDialog.ShowFor
    // when the user has opted into "Update Kodo in the background"
    // (MainWindow's IsAutoUpdateAppInBackgroundEnabled, or the equivalent flag
    // read via IsAutoUpdateInBackgroundEnabledInSettings for the launch-time
    // check). Mirrors UpdateDialog.BeginUpdateAsync's download -> launch
    // sequence, minus anything that needs a window to report progress to.
    public static async Task SilentlyInstallAsync(UpdateInfo update, CancellationToken ct = default)
    {
        try
        {
            var installerPath = await DownloadInstallerAsync(update, progress: null, ct).ConfigureAwait(false);
            LaunchInstallerAndExit(installerPath);
        }
        catch (Exception ex)
        {
            // Best-effort, like every other step of the update pipeline - log
            // and move on. The user simply isn't bumped to this version until
            // the next check picks it up again.
            KodoDiagnostics.WriteDiagnosticLog(
                source: "UpdateService.SilentlyInstallAsync",
                exception: ex,
                isTerminating: false,
                severity: "Warning",
                operation: "AutoUpdate");
        }
    }

    // ── Consolidated check-then-act flow ─────────────────────────────────────
    //
    // Every call site that wants to "check for an update and do something
    // about it" (the manual Settings button, the releases-page fallback, and
    // the periodic background timer) used to duplicate this exact branch:
    // check → if found, either silently install or show UpdateDialog. This
    // is the single place that logic lives now; callers just decide what to
    // do with the optional pre-install status callback and the final result.

    /// <summary>
    /// Checks for an update and, if one is found, either installs it silently
    /// or shows <see cref="UpdateDialog"/>, depending on <paramref name="installInBackground"/>.
    /// Returns the <see cref="UpdateInfo"/> that was found (after handling it),
    /// or <c>null</c> if no update was available. Never throws - failures are
    /// already swallowed by <see cref="CheckForUpdateAsync"/> and <see cref="SilentlyInstallAsync"/>.
    /// </summary>
    /// <param name="onUpdateFound">
    /// Optional callback invoked the moment an update is found, before the
    /// install/dialog branch runs - lets a caller update status text (e.g.
    /// "Kodo vX is available.") without waiting for a silent install to finish.
    /// </param>
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

    // ── GitHub API DTOs ──────────────────────────────────────────────────────

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

// ── Update dialog UI ─────────────────────────────────────────────────────────
// Sleek, self-contained update dialog. Built entirely in code (no AXAML) so it
// has zero dependency on App.axaml/MainWindow.axaml resources, mirroring the
// crash dialog's approach in App.axaml.cs. Walks the user through:
//   "Update available" -> [Update Now] -> progress bar -> installer launch.
internal sealed class UpdateDialog : Window
{
    // Same dark-surface palette used by every other Kodo dialog (see
    // DialogPalette below), so every dialog looks consistent regardless of
    // which theme/accent the user has selected.
    private static readonly Color SurfaceColor   = DialogPalette.Surface;
    private static readonly Color SurfaceDeep    = DialogPalette.SurfaceDeep;
    private static readonly Color BorderColor    = DialogPalette.Border;
    private static readonly Color BadgeBgColor   = DialogPalette.BadgeBg;
    private static readonly Color TextMutedColor = DialogPalette.TextMuted;
    private static readonly Color TextDimColor   = DialogPalette.TextDim;

    // Resolved once per dialog instance from the user's active accent setting
    // (Kodo purple / Windows accent / custom / theme), same as MainWindow.
    private readonly Color _accentColor;
    private readonly Color _accentForeground;

    private readonly UpdateInfo _update;
    private readonly string? _preDownloadedInstallerPath;
    private readonly TextBlock _statusText;
    private readonly ProgressBar _progressBar;
    private readonly Button _primaryButton;
    private readonly Button _laterButton;
    private readonly StackPanel _content;

    // preDownloadedInstallerPath is set when KodoUpdater (the standalone
    // background process) already fetched the installer and wrote the
    // pending-update sentinel - in that case "Update Now" skips straight to
    // launching it instead of re-downloading.
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

    // Shows the dialog non-modally if no owner can be safely used, or as a
    // modal dialog over the main window otherwise. Mirrors the crash dialog's
    // owner-safety check (a closing/invisible window throws on ShowDialog).
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

        // Fast path: KodoUpdater already downloaded this installer (sentinel
        // file case). Skip straight to the restart-and-install step.
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

            // Brief pause so the "downloaded" message is actually visible before
            // the app vanishes and the installer takes over.
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
// ── Shared dialog palette ────────────────────────────────────────────────────
// Single source of truth for the dark-surface colours used by every
// code-built dialog in the app (crash dialog in App.axaml.cs, update dialog
// below, and the various in-app dialogs in MainWindow.axaml.cs). Previously
// each dialog kept its own duplicate copy of these hex values; centralising
// them here means a single change actually propagates everywhere, rather than
// only where a comment claimed it would.
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

// ── Accent colour resolution ─────────────────────────────────────────────────
// Resolves the user's chosen accent colour outside of MainWindow, for dialogs
// (crash dialog, update dialog, etc.) that may need to appear before a
// MainWindow exists or independently of it. Mirrors MainWindow's
// ApplyAccentColorMode resolution switch so every dialog in the app agrees on
// what "the accent colour" currently means.
//
// NOTE on "theme" mode: MainWindow derives the theme accent at runtime from
// whichever extension theme is actively loaded, and that value isn't persisted
// to kodosettings.json. Since standalone dialogs can't safely load the extension
// system themselves, "theme" mode falls back to the fixed Kodo purple here -
// the same fallback MainWindow itself uses before any theme has loaded.
internal static class AccentResolver
{
    private const string DefaultAccentHex = "#8C00FF";
    private const string SettingsFileName = "kodosettings.json";

    // Cached for the lifetime of the process - dialogs are short-lived, and the
    // accent rarely changes mid-session, so re-reading the registry/disk on
    // every dialog open isn't worth the (tiny) risk of a stale read mattering.
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
            // "theme" can't be resolved without the extension system loaded -
            // fall back to the fixed Kodo purple, matching MainWindow's own
            // pre-theme-load default.
            "theme"   => DefaultAccentHex,
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
            // Settings are a convenience read here; any failure just means the
            // dialog falls back to the default Kodo purple.
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

    // Returns Brushes-equivalent White or Black colour depending on which gives
    // better contrast against the supplied accent colour, using the WCAG
    // relative-luminance formula. Identical logic to MainWindow's
    // GetAccentForeground, duplicated here so dialogs don't need a MainWindow
    // instance to call into.
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

    // Minimal subset of MainWindow's AppSettings needed to resolve the accent.
    // Kept separate (rather than reusing MainWindow's private AppSettings) so
    // this file has zero compile-time dependency on MainWindow.
    private sealed class AccentSettings
    {
        public string AccentColorMode { get; set; } = "kodo";
        public string CustomAccentHex { get; set; } = DefaultAccentHex;
    }
}

// ── Pending-update sentinel ──────────────────────────────────────────────────
// Reads the sentinel file written by the standalone KodoUpdater.exe background
// process (see KodoUpdater/Program.cs). KodoUpdater runs independently of
// Kodo - including while Kodo is closed - polling GitHub every 6 hours. When
// it finds and downloads an update it either installs silently (if the user
// has "install in background" enabled and Kodo isn't running) or, more
// commonly, leaves the installer on disk and writes this file so Kodo's own
// launch-time check can show UpdateDialog pre-loaded with the installer path,
// skipping the download step entirely.
internal static class PendingUpdateService
{
    private static string FilePath => Path.Combine(Path.GetTempPath(), "Kodo-Update", "pending.json");

    // Returns the pending update only if it's still newer than the running
    // version and the installer file it points to still exists on disk -
    // otherwise treats it the same as "no pending update" (and cleans up a
    // stale/invalid sentinel so it isn't re-checked every launch).
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
                // Already on this version or newer (e.g. user updated
                // manually) - the downloaded installer is now stale.
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
// ── Periodic background app-update scheduler ────────────────────────────────
//
// Owns the DispatcherTimer that drives "check for a new Kodo build every six
// hours while the app stays open" - the launch-time check itself still lives
// in App.axaml.cs (CheckForUpdatesInBackground); this only covers everything
// after that. MainWindow owns one instance and feeds it small accessors for
// the settings it needs to read, rather than the scheduler reaching back into
// MainWindow's full surface area.
internal sealed class AppUpdateScheduler
{
    private readonly DispatcherTimer _timer = new() { Interval = TimeSpan.FromHours(6) };
    private readonly Func<bool> _isEnabled;
    private readonly Func<bool> _isManualCheckInProgress;
    private readonly Func<bool> _installInBackground;

    /// <param name="isEnabled">Mirrors the "Automatically check for and install Kodo updates" setting.</param>
    /// <param name="isManualCheckInProgress">
    /// True while a manual "Check for Updates" click (Settings page) is in
    /// flight, so the periodic tick never races it or steps on the same
    /// status text.
    /// </param>
    /// <param name="installInBackground">Mirrors the "Update automatically without asking" sub-setting.</param>
    public AppUpdateScheduler(Func<bool> isEnabled, Func<bool> isManualCheckInProgress, Func<bool> installInBackground)
    {
        _isEnabled = isEnabled;
        _isManualCheckInProgress = isManualCheckInProgress;
        _installInBackground = installInBackground;
        _timer.Tick += async (_, _) => await OnTickAsync().ConfigureAwait(true);
    }

    /// <summary>
    /// Starts or stops the timer to match the current value of <c>isEnabled</c>.
    /// Call once at startup and again every time the setting is toggled.
    /// </summary>
    public void UpdateLifecycle()
    {
        _timer.Stop();
        if (_isEnabled())
            _timer.Start();
    }

    /// <summary>Stops the timer outright - call when the window is closing.</summary>
    public void Stop() => _timer.Stop();

    // Fires every six hours while Kodo stays open and the setting is enabled,
    // so a release published mid-session isn't only picked up on next launch.
    // Skips while a manual check from the Settings page is already in flight.
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