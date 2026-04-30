// Licensed under the Kodo Public License v1.0
// April 19th, 2026 - KerbalMissile - Added proper comments
using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using System;
using System.IO;
using System.Threading.Tasks;

namespace Kodo;

// One of the main entry points for the application, responsible for initializing and starting the app. This class is mostly boilerplate for Avalonia applications.
public partial class App : Application

    // Initializes the application by loading the XAML defined in App.axaml, which sets up resources and styles.
{
    public override void Initialize()
    {
        AppDomain.CurrentDomain.UnhandledException += CurrentDomain_OnUnhandledException;
        TaskScheduler.UnobservedTaskException += TaskScheduler_OnUnobservedTaskException;
        AvaloniaXamlLoader.Load(this);
    }

    // Called when the application has finished initializing. Here we check if we're running in a desktop environment and if so, we create and show the main window of the application.
    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            desktop.MainWindow = new MainWindow();
        }

        base.OnFrameworkInitializationCompleted();
    }

    private static void CurrentDomain_OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception exception)
            WriteCrashLog("AppDomain.UnhandledException", exception);
    }

    private static void TaskScheduler_OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        WriteCrashLog("TaskScheduler.UnobservedTaskException", e.Exception);
        e.SetObserved();
    }

    private static void WriteCrashLog(string source, Exception exception)
    {
        try
        {
            var logDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Kodo");
            Directory.CreateDirectory(logDirectory);
            var logPath = Path.Combine(logDirectory, "crash.log");
            var content =
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {source}{Environment.NewLine}{exception}{Environment.NewLine}{Environment.NewLine}";
            File.AppendAllText(logPath, content);
        }
        catch
        {
            // Last-resort logging should never crash the app.
        }
    }
}
