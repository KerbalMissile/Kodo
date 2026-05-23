using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace Kodo;

internal static class KodoDiagnostics
{
    private static readonly string CurrentAppVersion =
        Assembly.GetExecutingAssembly()
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
            ?.InformationalVersion ?? "v0.0.0";

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
        try
        {
            Directory.CreateDirectory(LogDirectoryPath);
            var payload = BuildDiagnosticPayload(source, exception, isTerminating, severity, operation);
            File.AppendAllText(LogFilePath, payload + Environment.NewLine);
        }
        catch
        {
        }
    }

    public static void WriteDebugFallback(string message, Exception? exception = null)
    {
        try
        {
            Debug.WriteLine(exception is null ? $"[Kodo] {message}" : $"[Kodo] {message}{Environment.NewLine}{exception}");
        }
        catch
        {
        }
    }
}
