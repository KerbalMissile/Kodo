using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;

namespace KodoUpdater;

// KodoUpdater is a tiny, headless, long-lived background process. Its only job
// is to poll GitHub Releases for a newer Kodo build every 6 hours - even while
// Kodo itself is closed - and either:
//   a) silently download + install the update directly (when the user has
//      opted into "Automatically install updates in the background" AND Kodo
//      isn't currently running), or
//   b) download the installer and write a "pending update" sentinel file that
//      Kodo picks up and shows a normal Update dialog for on its next launch.
//
// It deliberately has zero dependency on Avalonia or any Kodo.exe types -
// it's a standalone exe so it can run detached from Kodo's own process tree.
// Anything it can't do safely (parse settings, reach GitHub, write a file) is
// swallowed and retried on the next cycle; this process must never crash loud
// enough to show a Windows error dialog, since nobody is watching it.
internal static class Program
{
    private const string RepoOwner = "KerbalMissile";
    private const string RepoName = "Kodo";
    private const string LatestReleaseUrl = $"https://api.github.com/repos/{RepoOwner}/{RepoName}/releases/latest";
    private static readonly TimeSpan PollInterval = TimeSpan.FromHours(6);

    private static readonly HttpClient Http = CreateHttpClient();

    private static async Task Main(string[] args)
    {
        // Single-instance guard: if a previous KodoUpdater is already resident
        // (e.g. Kodo was relaunched and spawned another one), just exit. Using
        // a named mutex rather than a process-name check avoids false positives
        // from unrelated exes that happen to share the name during debugging.
        //
        // Wrapped in its own try/catch - if mutex creation itself fails for some
        // reason (rare, but possible in locked-down environments), we'd rather
        // risk a duplicate resident process than have the updater die silently
        // before it ever reaches the poll loop. Nobody is watching this process,
        // so an unhandled exception here would just mean updates quietly stop
        // working forever with zero indication why.
        Mutex? singleInstance = null;
        try
        {
            singleInstance = new Mutex(initiallyOwned: true, "Kodo-KodoUpdater-SingleInstance", out var createdNew);
            if (!createdNew) return;
        }
        catch
        {
            // Fall through and run anyway.
        }

        try
        {
            while (true)
            {
                try
                {
                    await RunOneCycleAsync();
                }
                catch
                {
                    // Never let one bad cycle kill the whole resident process.
                }

                await Task.Delay(PollInterval);
            }
        }
        finally
        {
            singleInstance?.Dispose();
        }
    }

    private static async Task RunOneCycleAsync()
    {
        var settings = ReadSettings();
        if (!settings.AutoUpdateAppEnabled)
            return;

        // Already have a pending, not-yet-installed update from a previous
        // cycle - don't re-download, just leave the sentinel for Kodo (or
        // retry the silent install if Kodo still isn't running).
        var pending = PendingUpdate.TryRead();
        if (pending is not null && File.Exists(pending.InstallerPath))
        {
            if (settings.AutoUpdateAppInBackgroundEnabled && !IsKodoRunning())
                LaunchInstallerSilently(pending.InstallerPath);
            return;
        }

        var localVersion = ReadInstalledKodoVersion();
        var update = await CheckForUpdateAsync(localVersion);
        if (update is null)
            return;

        string installerPath;
        try
        {
            installerPath = await DownloadInstallerAsync(update);
        }
        catch
        {
            return; // Network hiccup - try again next cycle.
        }

        if (settings.AutoUpdateAppInBackgroundEnabled && !IsKodoRunning())
        {
            // Fully silent path: install right away, no sentinel needed since
            // there's nothing left for Kodo to prompt about.
            LaunchInstallerSilently(installerPath);
            return;
        }

        // Otherwise, leave the installer downloaded and write the sentinel so
        // Kodo's own launch-time check can show UpdateDialog pre-loaded with
        // installerPath, skipping the download step entirely.
        PendingUpdate.Write(update.Version, installerPath);
    }

    // ── Settings ──────────────────────────────────────────────────────────

    private static UpdaterSettings ReadSettings()
    {
        try
        {
            var path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Kodo", "kodosettings.json");

            if (!File.Exists(path)) return new UpdaterSettings();

            var json = File.ReadAllText(path);
            if (string.IsNullOrWhiteSpace(json)) return new UpdaterSettings();

            return JsonSerializer.Deserialize<UpdaterSettings>(json) ?? new UpdaterSettings();
        }
        catch
        {
            return new UpdaterSettings();
        }
    }

    private sealed class UpdaterSettings
    {
        public bool AutoUpdateAppEnabled { get; set; } = true;
        public bool AutoUpdateAppInBackgroundEnabled { get; set; }
    }

    // ── Version / process detection ──────────────────────────────────────

    // KodoUpdater.exe is installed side-by-side with Kodo.exe, so its own
    // directory is the install directory - no registry lookup needed.
    private static string ReadInstalledKodoVersion()
    {
        try
        {
            var exeDir = AppContext.BaseDirectory;
            var kodoExePath = Path.Combine(exeDir, "Kodo.exe");
            if (!File.Exists(kodoExePath)) return "v0.0.0";

            var info = FileVersionInfo.GetVersionInfo(kodoExePath);
            var raw = info.ProductVersion ?? info.FileVersion ?? "v0.0.0";
            var plusIndex = raw.IndexOf('+');
            return plusIndex >= 0 ? raw[..plusIndex] : raw;
        }
        catch
        {
            return "v0.0.0";
        }
    }

    private static bool IsKodoRunning()
    {
        try
        {
            return Process.GetProcessesByName("Kodo").Length > 0;
        }
        catch
        {
            // If we can't tell, assume it's running - safer to fall back to
            // the sentinel-prompt path than to silently kill/replace a file
            // a running Kodo might be touching.
            return true;
        }
    }

    // ── GitHub release check / download (mirrors Kodo's own UpdateService) ─

    private static HttpClient CreateHttpClient()
    {
        var client = new HttpClient { Timeout = TimeSpan.FromSeconds(15) };
        client.DefaultRequestHeaders.UserAgent.ParseAdd("Kodo-Updater");
        client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
        return client;
    }

    private static async Task<UpdateInfo?> CheckForUpdateAsync(string localVersion)
    {
        try
        {
            using var response = await Http.GetAsync(LatestReleaseUrl);
            if (!response.IsSuccessStatusCode) return null;

            await using var stream = await response.Content.ReadAsStreamAsync();
            var release = await JsonSerializer.DeserializeAsync<GitHubRelease>(stream, JsonOptions);
            if (release is null || string.IsNullOrWhiteSpace(release.TagName)) return null;
            if (release.Draft || release.Prerelease) return null;
            if (!IsNewerVersion(release.TagName, localVersion)) return null;

            var asset = release.Assets?.FirstOrDefault(a =>
                a.Name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase));
            if (asset is null) return null;

            return new UpdateInfo(release.TagName, asset.BrowserDownloadUrl, asset.Name);
        }
        catch
        {
            return null;
        }
    }

    private static bool IsNewerVersion(string remote, string local)
    {
        var r = ParseVersionParts(remote);
        var l = ParseVersionParts(local);
        if (r is null || l is null) return !string.Equals(remote, local, StringComparison.OrdinalIgnoreCase);

        for (var i = 0; i < Math.Max(r.Length, l.Length); i++)
        {
            var rv = i < r.Length ? r[i] : 0;
            var lv = i < l.Length ? l[i] : 0;
            if (rv != lv) return rv > lv;
        }
        return false;
    }

    private static int[]? ParseVersionParts(string tag)
    {
        var core = tag.Trim();
        if (core.Length > 0 && (core[0] == 'v' || core[0] == 'V')) core = core[1..];
        var dash = core.IndexOf('-'); if (dash >= 0) core = core[..dash];
        var plus = core.IndexOf('+'); if (plus >= 0) core = core[..plus];

        var segments = core.Split('.');
        var parts = new int[segments.Length];
        for (var i = 0; i < segments.Length; i++)
            if (!int.TryParse(segments[i], out parts[i])) return null;

        return parts.Length > 0 ? parts : null;
    }

    private static async Task<string> DownloadInstallerAsync(UpdateInfo update)
    {
        var dir = Path.Combine(Path.GetTempPath(), "Kodo-Update");
        Directory.CreateDirectory(dir);
        var destPath = Path.Combine(dir, update.AssetName);

        using var response = await Http.GetAsync(update.AssetDownloadUrl, HttpCompletionOption.ResponseHeadersRead);
        response.EnsureSuccessStatusCode();

        await using var httpStream = await response.Content.ReadAsStreamAsync();
        await using var fileStream = new FileStream(destPath, FileMode.Create, FileAccess.Write, FileShare.None);
        await httpStream.CopyToAsync(fileStream);

        return destPath;
    }

    // Launches the installer with the same silent flags Kodo's UpdateService
    // uses, but with no foreground process to exit afterward - this updater
    // process simply keeps running its poll loop.
    private static void LaunchInstallerSilently(string installerPath)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = installerPath,
                Arguments = "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /CLOSEAPPLICATIONS /RESTARTAPPLICATIONS",
                UseShellExecute = true,
            });
            PendingUpdate.Clear();
        }
        catch
        {
            // Leave the downloaded installer in place; next cycle retries.
        }
    }

    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    private sealed record UpdateInfo(string Version, string AssetDownloadUrl, string AssetName);

    private sealed class GitHubRelease
    {
        [JsonPropertyName("tag_name")] public string TagName { get; set; } = "";
        [JsonPropertyName("draft")] public bool Draft { get; set; }
        [JsonPropertyName("prerelease")] public bool Prerelease { get; set; }
        [JsonPropertyName("assets")] public GitHubAsset[]? Assets { get; set; }
    }

    private sealed class GitHubAsset
    {
        [JsonPropertyName("name")] public string Name { get; set; } = "";
        [JsonPropertyName("browser_download_url")] public string BrowserDownloadUrl { get; set; } = "";
    }
}

// Shared sentinel-file format. Kept identical in shape to the one Kodo.exe
// reads in App.axaml.cs (PendingUpdateService) - see Updater.cs - so either
// side can write/read it without a shared assembly reference.
internal static class PendingUpdate
{
    private static string FilePath => Path.Combine(Path.GetTempPath(), "Kodo-Update", "pending.json");

    public static void Write(string version, string installerPath)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            var payload = JsonSerializer.Serialize(new PendingUpdateRecord(version, installerPath, DateTime.UtcNow));
            File.WriteAllText(FilePath, payload);
        }
        catch
        {
            // Best-effort - if this fails, Kodo's next 6h in-app check just
            // re-discovers the same update and prompts normally.
        }
    }

    public static PendingUpdateRecord? TryRead()
    {
        try
        {
            if (!File.Exists(FilePath)) return null;
            var json = File.ReadAllText(FilePath);
            return JsonSerializer.Deserialize<PendingUpdateRecord>(json);
        }
        catch
        {
            return null;
        }
    }

    public static void Clear()
    {
        try { File.Delete(FilePath); } catch { /* ignore */ }
    }
}

internal sealed record PendingUpdateRecord(string Version, string InstallerPath, DateTime DownloadedAtUtc);