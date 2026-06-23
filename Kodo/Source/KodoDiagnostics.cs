using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace Kodo;

// ── Severity tiers ────────────────────────────────────────────────────────────
//
//  Critical  – unhandled exceptions, startup crashes, data-loss risk.
//              Written to crash.log.  Always shows a crash dialog.
//
//  Warning   – recoverable operation failures (network, file I/O, extension
//              load errors, etc.).  Written to warnings.log.  Shows the
//              warning dialog in the UI.
//
//  Debug     – internal diagnostics that should not surface to the user
//              (cache hits, suppressed duplicates, dev info).
//              Written only to Debug output; never produces a file or dialog
//              unless an exception is attached, in which case it appends to
//              warnings.log silently.
//
public enum KodoSeverity { Critical, Warning, Debug }

internal static class KodoDiagnostics
{
    public static string AppVersion { get; } = ResolveAppVersion();
    private static readonly string CurrentAppVersion = AppVersion;

    private static string ResolveAppVersion()
    {
        var raw = Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion ?? "v0.0.0";

        // .NET SDK projects automatically append +<git-commit-hash> to
        // InformationalVersion (e.g. "v1.0.3+abc1234def"). Strip the suffix
        // so the logged version matches the declared <InformationalVersion>
        // in the .csproj and is not confused with a different build.
        var plusIndex = raw.IndexOf('+');
        return plusIndex >= 0 ? raw[..plusIndex] : raw;
    }

    // ── Log paths ─────────────────────────────────────────────────────────────

    public static string LogDirectoryPath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Kodo");

    // Critical / unhandled-exception log (was the only log file before).
    public static string CrashLogFilePath => Path.Combine(LogDirectoryPath, "crash.log");

    // Recoverable-warning log - operation failures that showed a dialog but
    // did not crash the app.  Kept separate so users and devs can quickly
    // distinguish "the app died" from "an extension failed to load".
    public static string WarningsLogFilePath => Path.Combine(LogDirectoryPath, "warnings.log");

    // Legacy alias kept so existing callers that read LogFilePath still compile
    // without changes.  Points at crash.log (unchanged behaviour).
    public static string LogFilePath => CrashLogFilePath;

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
        sb.Append("Version: ").AppendLine(CurrentAppVersion);
        sb.Append("OS: ").AppendLine(RuntimeInformation.OSDescription);
        sb.Append("Runtime: ").AppendLine(RuntimeInformation.FrameworkDescription);
        sb.Append("Architecture: ").Append(RuntimeInformation.ProcessArchitecture)
          .Append(" / ").AppendLine(Environment.Is64BitProcess ? "64-bit" : "32-bit");
        sb.Append("Log Path: ").AppendLine(severity == KodoSeverity.Critical ? CrashLogFilePath : WarningsLogFilePath);
        sb.AppendLine();
        sb.AppendLine(exception.ToString());
        return sb.ToString();
    }

    // Overload that accepts a raw severity string for legacy call sites that
    // haven't been migrated yet.  Maps "Crash" → Critical, anything else → Warning.
    public static string BuildDiagnosticPayload(
        string source,
        Exception exception,
        bool isTerminating,
        string severityLabel,
        string? operation = null) =>
        BuildDiagnosticPayload(source, exception, isTerminating, ParseSeverity(severityLabel), operation);

    public static string BuildDiagnosticSummary(
        string source,
        bool isTerminating,
        string? operation = null)
    {
        var timestamp = UtcNow().ToString("yyyy-MM-dd HH:mm:ss") + " UTC";
        var summary = new StringBuilder();
        summary.Append("Time: ").Append(timestamp)
               .Append("  |  Source: ").Append(source)
               .Append("  |  Version: ").Append(CurrentAppVersion);

        if (!string.IsNullOrWhiteSpace(operation))
            summary.Append("  |  Operation: ").Append(operation);

        summary.Append("  |  ").Append(isTerminating ? "Terminating" : "Recoverable");
        return summary.ToString();
    }

    // ── Typed logging API ─────────────────────────────────────────────────────

    /// <summary>
    /// Log a Critical-severity event to crash.log.
    /// Use for unhandled exceptions and situations where data loss may have occurred.
    /// </summary>
    public static void LogCritical(
        string source,
        Exception exception,
        bool isTerminating,
        string? operation = null) =>
        WriteToLog(source, exception, isTerminating, KodoSeverity.Critical, operation);

    /// <summary>
    /// Log a Warning-severity event to warnings.log.
    /// Use for recoverable failures that the user has been (or will be) informed about.
    /// </summary>
    public static void LogWarning(
        string source,
        Exception exception,
        string? operation = null) =>
        WriteToLog(source, exception, isTerminating: false, KodoSeverity.Warning, operation);

    /// <summary>
    /// Emit a Debug trace.  Only writes to the Debug output; if an exception is
    /// provided it is also appended silently to warnings.log.
    /// </summary>
    public static void LogDebug(string message, Exception? exception = null)
    {
        try
        {
            Debug.WriteLine(exception is null
                ? $"[Kodo] {message}"
                : $"[Kodo] {message}{Environment.NewLine}{exception}");

            if (exception is not null)
                WriteToLog("KodoDiagnostics.Debug", exception, false, KodoSeverity.Debug, message);
        }
        catch { /* never throw from a debug trace */ }
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

    /// <inheritdoc cref="LogDebug"/>
    /// <remarks>Legacy overload - prefer <see cref="LogDebug"/>.</remarks>
    public static void WriteDebugFallback(string message, Exception? exception = null) =>
        LogDebug(message, exception);

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

        // Debug entries go to warnings.log (silently, no dialog); Critical goes
        // to crash.log; Warning goes to warnings.log.
        var targetPath = severity == KodoSeverity.Critical ? CrashLogFilePath : WarningsLogFilePath;
        WritePayloadToDisk(payload, targetPath);
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