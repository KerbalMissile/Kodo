using Avalonia;
using System;
using System.Runtime.InteropServices;

namespace Kodo;

class Program
{
    // Attach to the console of the parent process (e.g. a terminal you launched
    // from). This is a no-op when the app is double-clicked normally, so no
    // console window flashes open. ATTACH_PARENT_PROCESS = 0xFFFFFFFF.
    [DllImport("kernel32.dll")]
    private static extern bool AttachConsole(uint dwProcessId);

    // Initialization code. Don't use any Avalonia, third-party APIs or any
    // SynchronizationContext-reliant code before AppMain is called: things aren't initialized
    // yet and stuff might break.
    [STAThread]
    public static void Main(string[] args)
    {
        // Re-attach stdout/stderr to the parent terminal so Console.WriteLine
        // output is visible when running from a shell. Has no effect when
        // launched via Explorer or Start Menu.
        AttachConsole(0xFFFFFFFF);

        // Enable legacy code pages like Windows-1252 for file re-save and
        // extension-driven workflows on .NET.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Last-resort handler: catches crashes that happen before or during
        // Avalonia startup (e.g. static constructors, AXAML load failures).
        // App.axaml.cs registers its own AppDomain handler inside Initialize(),
        // but that runs after BuildAvaloniaApp(), so anything that throws before
        // that point would be completely silent without this guard.
        AppDomain.CurrentDomain.UnhandledException += static (_, e) =>
        {
            if (e.ExceptionObject is not Exception ex) return;
            try
            {
                KodoDiagnostics.WriteDiagnosticLog(
                    "Program.Main.UnhandledException",
                    ex,
                    isTerminating: e.IsTerminating,
                    severity: "Crash",
                    operation: "Startup");
            }
            catch { /* cannot crash the crash handler */ }
        };

        BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    // Avalonia configuration, don't remove; also used by visual designer.
    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}