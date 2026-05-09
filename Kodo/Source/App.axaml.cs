// Licensed under the Kodo Public License v1.0
// May 8th, 2026 - Added error popup dialog that mirrors the crash log entry
// April 19th, 2026 - KerbalMissile - Added proper comments
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Layout;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Avalonia.Threading;
using Microsoft.Win32;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
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
            // args[0] is the file path when launched via "Open with" or double-click
            var startupFilePath = desktop.Args?.Length > 0 ? desktop.Args[0] : null;
            desktop.MainWindow = new MainWindow(startupFilePath);
        }

#if !DEBUG
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            RegisterFileAssociations();
#endif

        base.OnFrameworkInitializationCompleted();
    }

    [SupportedOSPlatform("windows")]
    private static void RegisterFileAssociations()
    {
        try
        {
            var exe = Environment.ProcessPath;
            if (string.IsNullOrWhiteSpace(exe)) return;

            var command = $"\"{exe}\" \"%1\"";

            // Register the application itself
            using (var appKey = Registry.CurrentUser.CreateSubKey(@"Software\Classes\Applications\Kodo.exe"))
            {
                appKey.SetValue("FriendlyAppName", "Kodo");
                using var openKey = appKey.CreateSubKey(@"shell\open\command");
                openKey.SetValue("", command);
            }

            // Extensions to appear in "Open with"
            string[] extensions =
            [
                ".txt", ".md", ".cs", ".fs", ".vb",
                ".js", ".ts", ".jsx", ".tsx",
                ".html", ".htm", ".css", ".scss", ".sass",
                ".json", ".xml", ".yaml", ".yml", ".toml",
                ".py", ".rb", ".go", ".rs", ".cpp", ".c", ".h",
                ".sh", ".bat", ".ps1",
                ".sln", ".csproj", ".fsproj",
                ".gitignore", ".env", ".ini", ".cfg", ".config",
                ".log",
            ];

            foreach (var ext in extensions)
            {
                using var extKey = Registry.CurrentUser.CreateSubKey(
                    $@"Software\Classes\{ext}\OpenWithList\Kodo.exe");
                extKey?.SetValue("", "");
            }
        }
        catch
        {
            // Registration failure should never crash the app.
        }
    }

    private static void CurrentDomain_OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception exception)
        {
            WriteCrashLog("AppDomain.UnhandledException", exception);
            ShowErrorDialog("AppDomain.UnhandledException", exception);
        }
    }

    private static void TaskScheduler_OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        WriteCrashLog("TaskScheduler.UnobservedTaskException", e.Exception);
        ShowErrorDialog("TaskScheduler.UnobservedTaskException", e.Exception);
        e.SetObserved();
    }

    // Dispatches an error popup to the UI thread. The dialog mirrors the crash log entry so
    // the user sees the same source and stack trace that was written to disk.
    private static void ShowErrorDialog(string source, Exception exception)
    {
        try
        {
            // Get the crash log path to show the user where the full log was written.
            var logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Kodo", "crash.log");

            Dispatcher.UIThread.Post(() =>
            {
                try
                {
                    // Find the current main window to use as the dialog owner, if available.
                    Window? owner = null;
                    if (Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
                        owner = desktop.MainWindow;

                    var dialog = BuildErrorDialog(source, exception, logPath, owner);

                    if (owner is not null)
                        dialog.ShowDialog(owner);
                    else
                        dialog.Show();
                }
                catch
                {
                    // Dialog creation should never crash the app.
                }
            }, DispatcherPriority.MaxValue);
        }
        catch
        {
            // Dispatcher post should never crash the app.
        }
    }

    // Builds the error dialog window entirely in code so it works without any AXAML dependency.
    private static Window BuildErrorDialog(string source, Exception exception, string logPath, Window? owner)
    {
        // --- Header ---
        var titleText = new TextBlock
        {
            Text = "An unexpected error occurred",
            FontSize = 16,
            FontWeight = FontWeight.SemiBold,
            Foreground = Brushes.White,
            TextWrapping = TextWrapping.Wrap,
        };

        var subtitleText = new TextBlock
        {
            Text = "Kodo encountered an unhandled error. Your work may not have been affected, but this event has been logged.",
            FontSize = 13,
            Foreground = new SolidColorBrush(Color.Parse("#A0A0A0")),
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 4, 0, 0),
        };

        // --- Source badge ---
        var sourceBadge = new Border
        {
            Background = new SolidColorBrush(Color.Parse("#2B2B2B")),
            BorderBrush = new SolidColorBrush(Color.Parse("#3A3A3A")),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(6),
            Padding = new Thickness(10, 5),
            HorizontalAlignment = HorizontalAlignment.Left,
            Child = new TextBlock
            {
                Text = source,
                FontSize = 12,
                FontFamily = new FontFamily("Cascadia Code,Consolas,Menlo,monospace"),
                Foreground = new SolidColorBrush(Color.Parse("#9CDCFE")),
            },
        };

        // --- Exception details (scrollable) ---
        var exceptionText = new SelectableTextBlock
        {
            Text = exception.ToString(),
            FontSize = 12,
            FontFamily = new FontFamily("Cascadia Code,Consolas,Menlo,monospace"),
            Foreground = new SolidColorBrush(Color.Parse("#CE9178")),
            TextWrapping = TextWrapping.Wrap,
        };

        var exceptionScroll = new ScrollViewer
        {
            Content = exceptionText,
            MaxHeight = 260,
            VerticalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
        };

        var exceptionBorder = new Border
        {
            Background = new SolidColorBrush(Color.Parse("#1A1A1A")),
            BorderBrush = new SolidColorBrush(Color.Parse("#3A3A3A")),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(8),
            Padding = new Thickness(12),
            Child = exceptionScroll,
        };

        // --- Log path note ---
        var logPathText = new TextBlock
        {
            Text = $"Full log written to: {logPath}",
            FontSize = 11,
            Foreground = new SolidColorBrush(Color.Parse("#606060")),
            TextWrapping = TextWrapping.Wrap,
        };

        // --- Dismiss button ---
        var dismissButton = new Button
        {
            Content = "Dismiss",
            HorizontalAlignment = HorizontalAlignment.Right,
            Padding = new Thickness(20, 8),
            Background = new SolidColorBrush(Color.Parse("#8C00FF")),
            Foreground = Brushes.White,
            BorderThickness = new Thickness(0),
            CornerRadius = new CornerRadius(8),
        };

        // --- Layout ---
        var content = new StackPanel
        {
            Spacing = 12,
            Margin = new Thickness(20),
            Children =
            {
                titleText,
                subtitleText,
                sourceBadge,
                exceptionBorder,
                logPathText,
                dismissButton,
            },
        };

        var dialog = new Window
        {
            Title = "Kodo — Unhandled Error",
            Width = 560,
            SizeToContent = SizeToContent.Height,
            MinWidth = 400,
            MinHeight = 200,
            MaxHeight = 700,
            CanResize = true,
            WindowStartupLocation = owner is not null
                ? WindowStartupLocation.CenterOwner
                : WindowStartupLocation.CenterScreen,
            Background = new SolidColorBrush(Color.Parse("#1E1E1E")),
            Content = content,
        };

        // Wire up the dismiss button now that we have a reference to the dialog.
        dismissButton.Click += (_, _) => dialog.Close();

        return dialog;
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