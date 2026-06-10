using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Collections.Concurrent;

namespace Kodo;

internal static class KodoDiagnostics
{
    private static readonly string CurrentAppVersion = ResolveAppVersion();
    private static readonly ConcurrentDictionary<string, byte> LoggedExceptionFingerprints = new(StringComparer.Ordinal);

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

    public static string LogDirectoryPath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Kodo");

    public static string LogFilePath => Path.Combine(LogDirectoryPath, "crash.log");

    public static DateTime UtcNow() => DateTime.UtcNow;

    public static string BuildDiagnosticPayload(string source, Exception exception, bool isTerminating, string severity, string? operation = null)
    {
        var timestamp = UtcNow();
        var sb = new StringBuilder();
        sb.Append('[').Append(timestamp.ToString("yyyy-MM-dd HH:mm:ss")).Append(" UTC]").Append(' ')
            .Append(severity);

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
        sb.Append("Log Path: ").AppendLine(LogFilePath);
        sb.AppendLine();
        sb.AppendLine(exception.ToString());
        return sb.ToString();
    }

    public static string BuildDiagnosticSummary(string source, bool isTerminating, string? operation = null)
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

    public static void WriteDiagnosticLog(string source, Exception exception, bool isTerminating, string severity, string? operation = null)
    {
        var payload = string.Empty;
        try { payload = BuildDiagnosticPayload(source, exception, isTerminating, severity, operation); }
        catch { payload = $"[Kodo] Log payload build failed. Source={source} Exception={exception?.GetType()}:{exception?.Message}"; }

        // Attempt 1: primary log path under %AppData%\Kodo
        try
        {
            Directory.CreateDirectory(LogDirectoryPath);
            File.AppendAllText(LogFilePath, payload + Environment.NewLine);
            return;
        }
        catch { /* fall through to attempt 2 */ }

        // Attempt 2: temp directory (survives even if AppData is inaccessible)
        try
        {
            var tempLog = Path.Combine(Path.GetTempPath(), "kodo-crash.log");
            File.AppendAllText(tempLog, payload + Environment.NewLine);
            return;
        }
        catch { /* fall through to attempt 3 */ }

        // Attempt 3: desktop (last resort - at least the user can find it)
        try
        {
            var desktopLog = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "kodo-crash.log");
            File.AppendAllText(desktopLog, payload + Environment.NewLine);
        }
        catch { /* all attempts exhausted - swallow */ }
    }

    public static void WriteDebugFallback(string message, Exception? exception = null)
    {
        try
        {
            Debug.WriteLine(exception is null ? $"[Kodo] {message}" : $"[Kodo] {message}{Environment.NewLine}{exception}");
            if (exception is not null)
            {
                WriteDiagnosticLog(
                    source: "KodoDiagnostics.DebugFallback",
                    exception: exception,
                    isTerminating: false,
                    severity: "Debug",
                    operation: message);
            }
        }
        catch
        {
        }
    }

    public static bool TryWriteFirstChanceLog(string source, Exception exception)
    {
        if (exception is OperationCanceledException)
            return false;

        var fingerprint = $"{source}|{exception.GetType().FullName}|{exception.Message}|{exception.StackTrace}";
        if (!LoggedExceptionFingerprints.TryAdd(fingerprint, 0))
            return false;

        WriteDiagnosticLog(source, exception, isTerminating: false, severity: "FirstChance");
        return true;
    }
}