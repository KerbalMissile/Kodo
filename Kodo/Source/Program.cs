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

        // AppDomain.UnhandledException is handled by App.Initialize() (CurrentDomain_OnUnhandledException),
        // which logs, tracks, and shows the crash dialog - registering it here too duplicated crash.log entries.
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