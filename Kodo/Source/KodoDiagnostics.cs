using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;
using Microsoft.Win32;

namespace Kodo;

// Critical: crashes, shows dialog. Warning: recoverable failures. Debug: internal diagnostics.
public enum KodoSeverity { Critical, Warning, Debug }

internal static class KodoDiagnostics
{
    public static string AppVersion { get; } = ResolveAppVersion();

    // CurrentBuildNumber is the only reliable Windows-11 signal; OSDescription still says 10.
    public static string OSDescription { get; } = ResolveOSDescription();

    // Debug traces also append to kodo.log when Developer Options is enabled.
    public static bool VerboseLoggingEnabled { get; set; }

    // Rolling window of recent log lines, flushed into crash.log on a Critical event.

    private const int BreadcrumbCapacity = 50;
    private static readonly Queue<string> _breadcrumbs = new();
    private static readonly object _breadcrumbLock = new();

    private static void PushBreadcrumb(string line)
    {
        lock (_breadcrumbLock)
        {
            if (_breadcrumbs.Count >= BreadcrumbCapacity)
                _breadcrumbs.Dequeue();
            _breadcrumbs.Enqueue(line);
        }
    }

    private static IReadOnlyList<string> DrainBreadcrumbs()
    {
        lock (_breadcrumbLock)
        {
            var snapshot = new List<string>(_breadcrumbs);
            // Keeps the buffer intact for a possible second crash dialog.
            return snapshot;
        }
    }

    // ── Version / OS helpers ──────────────────────────────────────────────────

    private static string ResolveAppVersion()
    {
        var raw = Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion ?? "v0.0.0";

        // Strips the +<git-hash> suffix so the logged version matches the csproj.
        var plusIndex = raw.IndexOf('+');
        return plusIndex >= 0 ? raw[..plusIndex] : raw;
    }

    private static string ResolveOSDescription()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return RuntimeInformation.OSDescription;

        return TryGetWindowsProductName() ?? RuntimeInformation.OSDescription;
    }

    [SupportedOSPlatform("windows")]
    private static string? TryGetWindowsProductName()
    {
        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(
                @"SOFTWARE\Microsoft\Windows NT\CurrentVersion");
            if (key is null) return null;

            var productName = key.GetValue("ProductName")        as string ?? string.Empty;
            var buildStr    = key.GetValue("CurrentBuildNumber") as string ?? string.Empty;
            var displayVer  = key.GetValue("DisplayVersion")     as string ?? string.Empty;

            if (!int.TryParse(buildStr, out var build)) return null;

            // Reattach the correct "Windows 10/11" prefix based on build number.
            var edition = productName
                .Replace("Windows 10", string.Empty, StringComparison.OrdinalIgnoreCase)
                .Replace("Windows 11", string.Empty, StringComparison.OrdinalIgnoreCase)
                .Trim();

            var winLabel = build >= 22000 ? "Windows 11" : "Windows 10";
            var fullName = string.IsNullOrWhiteSpace(edition) ? winLabel : $"{winLabel} {edition}";

            // e.g. "Windows 11 Pro 24H2 (build 26200)"
            return string.IsNullOrWhiteSpace(displayVer)
                ? $"{fullName} (build {build})"
                : $"{fullName} {displayVer} (build {build})";
        }
        catch
        {
            return null;
        }
    }

    // ── Log paths ─────────────────────────────────────────────────────────────

    public static string LogDirectoryPath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Kodo");

    // Main log - every severity tier appends here.
    public static string MainLogFilePath => Path.Combine(LogDirectoryPath, "kodo.log");

    // Crash-only log: breadcrumb buffer plus crash payload, easier to share than kodo.log.
    public static string CrashLogFilePath => Path.Combine(LogDirectoryPath, "crash.log");

    // Legacy alias - points at the main log so existing callers still compile.
    public static string LogFilePath => MainLogFilePath;

    public static DateTime UtcNow() => DateTime.UtcNow;

    // ── Payload builders ──────────────────────────────────────────────────────

    public static string BuildDiagnosticPayload(
        string source,
        Exception exception,
        bool isTerminating,
        KodoSeverity severity,
        string? operation = null)
    {
        var timestamp = UtcNow();
        var sb = new StringBuilder();
        sb.Append('[').Append(timestamp.ToString("yyyy-MM-dd HH:mm:ss")).Append(" UTC]")
          .Append(' ').Append(SeverityLabel(severity));

        if (!string.IsNullOrWhiteSpace(operation))
            sb.Append(" (").Append(operation).Append(')');

        sb.AppendLine();
        sb.Append("Source: ").AppendLine(source);
        sb.Append("Terminating: ").AppendLine(isTerminating ? "Yes" : "No");
        sb.Append("Version: ").AppendLine(AppVersion);
        sb.Append("OS: ").AppendLine(OSDescription);
        sb.Append("Runtime: ").AppendLine(RuntimeInformation.FrameworkDescription);
        sb.Append("Architecture: ").Append(RuntimeInformation.ProcessArchitecture)
          .Append(" / ").AppendLine(Environment.Is64BitProcess ? "64-bit" : "32-bit");
        sb.Append("Log: ").AppendLine(MainLogFilePath);
        sb.AppendLine();
        sb.AppendLine(exception.ToString());
        return sb.ToString();
    }

    public static string BuildDiagnosticSummary(
        string source,
        bool isTerminating,
        string? operation = null)
    {
        var timestamp = UtcNow().ToString("yyyy-MM-dd HH:mm:ss") + " UTC";
        var summary = new StringBuilder();
        summary.Append("Time: ").Append(timestamp)
               .Append("  |  Source: ").Append(source)
               .Append("  |  Version: ").Append(AppVersion);

        if (!string.IsNullOrWhiteSpace(operation))
            summary.Append("  |  Operation: ").Append(operation);

        summary.Append("  |  ").Append(isTerminating ? "Terminating" : "Recoverable");
        return summary.ToString();
    }

    // ── Typed logging API ─────────────────────────────────────────────────────

    /// <summary>Logs a Critical event to kodo.log and generates crash.log.</summary>
    public static void LogCritical(
        string source,
        Exception exception,
        bool isTerminating,
        string? operation = null)
    {
        WriteToLog(source, exception, isTerminating, KodoSeverity.Critical, operation);
        WriteCrashLog(source, exception, isTerminating, operation);
    }

    /// <summary>Logs a Warning event to kodo.log.</summary>
    public static void LogWarning(
        string source,
        Exception exception,
        string? operation = null) =>
        WriteToLog(source, exception, isTerminating: false, KodoSeverity.Warning, operation);

    /// <summary>Emits a Debug trace; also appends to kodo.log if an exception is attached.</summary>
    public static void LogDebug(string message, Exception? exception = null)
    {
        try
        {
            Debug.WriteLine(exception is null
                ? $"[Kodo] {message}"
                : $"[Kodo] {message}{Environment.NewLine}{exception}");

            if (exception is not null)
                WriteToLog("KodoDiagnostics.Debug", exception, false, KodoSeverity.Debug, message);
            else if (VerboseLoggingEnabled)
                WriteVerboseTrace(message);
        }
        catch { /* never throw from a debug trace */ }
    }

    // Appends a plain trace line to kodo.log; only called when VerboseLoggingEnabled is on.
    private static void WriteVerboseTrace(string message)
    {
        var timestamp = UtcNow().ToString("yyyy-MM-dd HH:mm:ss") + " UTC";
        var line = $"[{timestamp}] VERBOSE  {message}";
        PushBreadcrumb(line);
        WritePayloadToDisk(line, MainLogFilePath);
    }

    // ── Legacy API (kept for backward-compat; delegates to typed methods) ─────

    /// <inheritdoc cref="LogCritical"/>
    /// <remarks>Legacy overload - prefer <see cref="LogCritical"/> or <see cref="LogWarning"/>.</remarks>
    public static void WriteDiagnosticLog(
        string source,
        Exception exception,
        bool isTerminating,
        string severity,
        string? operation = null) =>
        WriteToLog(source, exception, isTerminating, ParseSeverity(severity), operation);

    // ── Internal write helpers ────────────────────────────────────────────────

    private static void WriteToLog(
        string source,
        Exception exception,
        bool isTerminating,
        KodoSeverity severity,
        string? operation)
    {
        var payload = string.Empty;
        try
        {
            payload = BuildDiagnosticPayload(source, exception, isTerminating, severity, operation);
        }
        catch
        {
            payload = $"[Kodo] Log payload build failed. Source={source} " +
                      $"Exception={exception?.GetType()}:{exception?.Message}";
        }

        // Every severity tier goes to kodo.log.
        PushBreadcrumb(payload);
        WritePayloadToDisk(payload, MainLogFilePath);
    }

    // Writes crash.log; called only from LogCritical.
    private static void WriteCrashLog(
        string source,
        Exception exception,
        bool isTerminating,
        string? operation)
    {
        try
        {
            var sb = new StringBuilder();
            sb.AppendLine("════════════════════════════════════════════════════════════");
            sb.Append('[').Append(UtcNow().ToString("yyyy-MM-dd HH:mm:ss")).AppendLine(" UTC] CRASH REPORT");
            sb.AppendLine("════════════════════════════════════════════════════════════");
            sb.AppendLine();

            // ── Recent activity leading up to the crash ───────────────────────
            var crumbs = DrainBreadcrumbs();
            if (crumbs.Count > 0)
            {
                sb.AppendLine("── Recent activity ──────────────────────────────────────────");
                foreach (var crumb in crumbs)
                    sb.AppendLine(crumb);
                sb.AppendLine();
            }

            // ── Crash payload ─────────────────────────────────────────────────
            sb.AppendLine("── Crash ────────────────────────────────────────────────────");
            sb.AppendLine(BuildDiagnosticPayload(source, exception, isTerminating, KodoSeverity.Critical, operation));

            WritePayloadToDisk(sb.ToString(), CrashLogFilePath);
        }
        catch { /* crash log generation must never itself crash */ }
    }

    private static void WritePayloadToDisk(string payload, string primaryPath)
    {
        // Attempt 1: designated log path under %AppData%\Kodo
        try
        {
            Directory.CreateDirectory(LogDirectoryPath);
            File.AppendAllText(primaryPath, payload + Environment.NewLine);
            return;
        }
        catch { /* fall through */ }

        // Attempt 2: temp directory (survives even if AppData is inaccessible)
        try
        {
            var tempLog = Path.Combine(Path.GetTempPath(), Path.GetFileName(primaryPath));
            File.AppendAllText(tempLog, payload + Environment.NewLine);
            return;
        }
        catch { /* fall through */ }

        // Attempt 3: desktop (last resort - at least the user can find it)
        try
        {
            var desktopLog = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Path.GetFileName(primaryPath));
            File.AppendAllText(desktopLog, payload + Environment.NewLine);
        }
        catch { /* all attempts exhausted - swallow */ }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static string SeverityLabel(KodoSeverity severity) => severity switch
    {
        KodoSeverity.Critical => "CRITICAL",
        KodoSeverity.Warning  => "WARNING",
        KodoSeverity.Debug    => "DEBUG",
        _                     => "INFO",
    };

    private static KodoSeverity ParseSeverity(string label) => label.ToUpperInvariant() switch
    {
        "CRASH" or "CRITICAL" or "FATAL" => KodoSeverity.Critical,
        "DEBUG"                           => KodoSeverity.Debug,
        _                                 => KodoSeverity.Warning,
    };
}
