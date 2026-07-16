using Avalonia;
using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace Kodo;

class Program
{
    [DllImport("kernel32.dll")]
    private static extern bool AttachConsole(uint dwProcessId);

    [STAThread]
    public static void Main(string[] args)
    {
        AttachConsole(0xFFFFFFFF);
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        AptabaseClient.Initialize();

        AppDomain.CurrentDomain.UnhandledException += static (_, e) =>
        {
            if (e.ExceptionObject is not Exception ex) return;
            try
            {
                AptabaseClient.TrackEvent("app_crash", ex.Message);
                KodoDiagnostics.LogCritical(
                    "Program.Main.UnhandledException",
                    ex,
                    isTerminating: e.IsTerminating,
                    operation: "Startup");
            }
            catch { }
        };

        var app = BuildAvaloniaApp();
        
        // Flush telemetry on exit
        try
        {
            app.StartWithClassicDesktopLifetime(args);
        }
        finally
        {
            // Give Aptabase time to send final events
            Task.Run(async () => await AptabaseClient.FlushAsync()).Wait(TimeSpan.FromSeconds(2));
        }
    }

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}
