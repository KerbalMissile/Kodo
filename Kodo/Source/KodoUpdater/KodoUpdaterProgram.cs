using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using Kodo;

namespace KodoUpdater;

internal static class Program
{
    private static readonly TimeSpan PollInterval = TimeSpan.FromMinutes(30);

    private static readonly HttpClient Http = CreateHttpClient();

    private static async Task Main(string[] args)
    {
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

        var localVersion = ReadInstalledKodoVersion();
        var pending = PendingUpdate.TryRead();
        if (pending is not null && File.Exists(pending.InstallerPath))
        {
            if (!IsNewerVersion(pending.Version, localVersion))
            {
                PendingUpdate.Clear();
                try { File.Delete(pending.InstallerPath); } catch { /* best-effort cleanup */ }
            }
            else
            {
                if (settings.AutoUpdateAppInBackgroundEnabled && !IsKodoRunning())
                    LaunchInstallerSilently(pending.InstallerPath);
                return;
            }
        }

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
            return true;
        }
    }

    // GitHub release check / download (mirrors Kodo's own UpdateService) 

    private static HttpClient CreateHttpClient()
    {
        var client = new HttpClient { Timeout = TimeSpan.FromSeconds(15) };
        client.DefaultRequestHeaders.UserAgent.ParseAdd("Kodo-Updater");
        client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
        return client;
    }

    private static async Task<UpdateInfo?> CheckForUpdateAsync(string localVersion)
    {
        foreach (var owner in GitHubRepoInfo.Owners)
        {
            try
            {
                var url = GitHubRepoInfo.GetLatestReleaseApiUrl(owner);
                using var response = await Http.GetAsync(url);
                if (!response.IsSuccessStatusCode) continue;

                await using var stream = await response.Content.ReadAsStreamAsync();
                var release = await JsonSerializer.DeserializeAsync<GitHubRelease>(stream, JsonOptions);
                if (release is null || string.IsNullOrWhiteSpace(release.TagName)) continue;
                if (release.Draft || release.Prerelease) continue;
                if (!IsNewerVersion(release.TagName, localVersion)) continue;

                var asset = release.Assets?.FirstOrDefault(a =>
                    a.Name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase));
                if (asset is null) continue;

                return new UpdateInfo(release.TagName, asset.BrowserDownloadUrl, asset.Name);
            }
            catch
            {
                // Try the other owner.
            }
        }

        return null;
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
