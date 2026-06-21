// Licensed under the Kodo Public License v1.1
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Input.Platform;
using Avalonia.Layout;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Avalonia.Threading;
using Microsoft.Win32;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Threading;
using System.Threading.Tasks;

namespace Kodo;

// One of the main entry points for the application, responsible for initializing and
// starting the app. Handles global unhandled-exception wiring, crash logging, and
// the error/crash dialog UI - all built in code so this file has no AXAML dependency.
public partial class App : Application
{
    // Prevent crash-dialog storms from re-entering the UI thread and starving input.
    private static int _isCrashDialogOpen;
    // ── Shared colours ───────────────────────────────────────────────────────
    // Both the crash dialog (App.axaml.cs) and the warning dialog (MainWindow.axaml.cs)
    // use the same dark-surface palette so they look consistent regardless of which
    // theme the user has selected.  Centralising them here means a single change
    // propagates to both dialogs.

    private static readonly Color KodoDarkSurface     = Color.Parse("#1E1E1E");
    private static readonly Color KodoDarkSurfaceDeep = Color.Parse("#1A1A1A");
    private static readonly Color KodoDarkBorder      = Color.Parse("#3A3A3A");
    private static readonly Color KodoDarkBadgeBg     = Color.Parse("#2B2B2B");
    private static readonly Color KodoTextMuted       = Color.Parse("#A0A0A0");
    private static readonly Color KodoTextDim         = Color.Parse("#606060");
    private static readonly Color KodoTokenBlue       = Color.Parse("#9CDCFE");  // source badge
    private static readonly Color KodoTokenOrange     = Color.Parse("#CE9178");  // stack trace

    // ── Initialization ───────────────────────────────────────────────────────

    // Loads AXAML resources/styles and wires up global exception handlers before
    // the framework has finished starting up.
    public override void Initialize()
    {
        AppDomain.CurrentDomain.UnhandledException += CurrentDomain_OnUnhandledException;
        TaskScheduler.UnobservedTaskException       += TaskScheduler_OnUnobservedTaskException;
        Dispatcher.UIThread.UnhandledException      += DispatcherUiThread_OnUnhandledException;
        AvaloniaXamlLoader.Load(this);
    }

    // Called when the framework has finished initializing. Creates the main window,
    // optionally pre-loading a file when Kodo is launched via "Open with" or by
    // double-clicking a file in Explorer.
    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            // desktop.Args[0] is the file path when launched via "Open with" / double-click.
            var startupFilePath = desktop.Args?.Length > 0 ? desktop.Args[0] : null;
            desktop.MainWindow  = new MainWindow(startupFilePath);
        }

#if !DEBUG
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            RegisterFileAssociations();
            CheckForUpdatesInBackground();
        }
#endif

        base.OnFrameworkInitializationCompleted();
    }

    // ── Auto-update ───────────────────────────────────────────────────────────

    // Fires a one-shot, fire-and-forget update check a few seconds after launch
    // so it never competes with startup for CPU/network. Entirely best-effort:
    // any failure here is swallowed by UpdateService itself and never surfaces
    // as a crash or dialog - the user simply won't see an update prompt.
    private static void CheckForUpdatesInBackground()
    {
        _ = Task.Run(async () =>
        {
            try
            {
                if (!UpdateService.IsAutoUpdateEnabledInSettings())
                    return;

                await Task.Delay(TimeSpan.FromSeconds(4));

                var update = await UpdateService.CheckForUpdateAsync();
                if (update is not null)
                    UpdateDialog.ShowFor(update);
            }
            catch
            {
                // Update checking must never crash the app.
            }
        });
    }

    // ── Windows file-association registration ────────────────────────────────

    [SupportedOSPlatform("windows")]
    private static void RegisterFileAssociations()
    {
        try
        {
            var exe = Environment.ProcessPath;
            if (string.IsNullOrWhiteSpace(exe)) return;

            var command = $"\"{exe}\" \"%1\"";

            // Register the application itself under HKCU so no elevation is needed.
            using (var appKey = Registry.CurrentUser.CreateSubKey(@"Software\Classes\Applications\Kodo.exe"))
            {
                appKey.SetValue("FriendlyAppName", "Kodo");
                using var openKey = appKey.CreateSubKey(@"shell\open\command");
                openKey.SetValue("", command);
            }

            // Extensions to appear in the "Open with" menu.
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
            // Registration failure must never crash the app; "Open with" is a convenience
            // feature and the app is fully functional without it.
        }
    }

    // ── Global exception handlers ────────────────────────────────────────────

    private static void CurrentDomain_OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is not Exception exception) return;

        KodoDiagnostics.WriteDiagnosticLog("AppDomain.UnhandledException", exception, e.IsTerminating, "Crash");
        ShowCrashDialog("AppDomain.UnhandledException", exception, isTerminating: e.IsTerminating);
    }

    private static void TaskScheduler_OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        KodoDiagnostics.WriteDiagnosticLog("TaskScheduler.UnobservedTaskException", e.Exception, false, "Crash");
        // Mark as observed first so the runtime does not re-throw it after we return.
        e.SetObserved();
        // This source is never IsTerminating, so the user can dismiss and keep working.
        ShowCrashDialog("TaskScheduler.UnobservedTaskException", e.Exception, isTerminating: false);
    }

    private static void DispatcherUiThread_OnUnhandledException(object? sender, DispatcherUnhandledExceptionEventArgs e)
    {
        KodoDiagnostics.WriteDiagnosticLog("Dispatcher.UIThread.UnhandledException", e.Exception, false, "Crash");
        ShowCrashDialog("Dispatcher.UIThread.UnhandledException", e.Exception, isTerminating: false);
        e.Handled = true;
    }

    // ── Crash dialog ─────────────────────────────────────────────────────────
    //
    // ShowCrashDialog dispatches a modal error dialog to the UI thread.
    //
    // WHY InvokeAsync + GetAwaiter().GetResult()?
    //   AppDomain.UnhandledException fires on a background (finalizer/thread-pool) thread
    //   just before the CLR may tear down the process (IsTerminating = true).  If we only
    //   Post() the work and return, the process exits before the UI thread renders anything.
    //   Blocking the crash-handler thread with GetAwaiter().GetResult() keeps the process
    //   alive until the user dismisses the dialog.
    //
    // EARLY-RETURN GUARD:
    //   On Linux/macOS the Dispatcher may already be stopped when a terminal exception
    //   fires.  Attempting to invoke on a stopped Dispatcher deadlocks.  We check
    //   CanCurrentThreadAccess first; if the answer is "no Dispatcher at all", we just
    //   return rather than hanging.

    private static void ShowCrashDialog(string source, Exception exception, bool isTerminating)
    {
        if (Interlocked.CompareExchange(ref _isCrashDialogOpen, 1, 0) != 0)
            return;

        // Reset the guard immediately - we use fire-and-forget below, so the
        // guard would never be cleared in the finally block on the background
        // thread, permanently preventing any future dialog.
        Interlocked.Exchange(ref _isCrashDialogOpen, 0);

        try
        {
            var logPath = KodoDiagnostics.LogFilePath;

            if (Dispatcher.UIThread.CheckAccess())
            {
                // Already on the UI thread - fire and forget.
                _ = ShowCrashDialogOnUiThreadAsync(source, exception, logPath, isTerminating);
                return;
            }

            // WHY Post() instead of InvokeAsync().GetAwaiter().GetResult():
            //
            // AppDomain.UnhandledException fires on a thread-pool thread after the
            // UI thread has already faulted. Calling GetResult() here blocks the
            // thread-pool thread waiting for the UI thread to drain its queue - but
            // if the UI thread is dead, that queue is never drained and the entire
            // process freezes at the OS level. The freeze is permanent: no window,
            // no crash log flush, nothing. Task Manager shows 0 % CPU forever.
            //
            // Instead, we Post() the dialog (non-blocking) and, for terminating
            // crashes, pump the dispatcher on this thread long enough for the dialog
            // to appear and for the user to dismiss it. The 30-second ceiling keeps
            // the process from hanging if the UI thread never recovers.
            if (isTerminating)
            {
                // Non-blocking post - returns immediately even if the UI thread is busy.
                Dispatcher.UIThread.Post(
                    () => _ = ShowCrashDialogOnUiThreadAsync(source, exception, logPath, isTerminating),
                    DispatcherPriority.MaxValue);

                // Give the UI thread up to 30 s to show and dismiss the dialog
                // before the CLR tears the process down. We sleep in short bursts
                // so we don't spin-waste CPU, and we exit the loop the moment the
                // dialog reports it has been dismissed.
                for (var i = 0; i < 300 && _isCrashDialogOpen == 0; i++)
                    Thread.Sleep(100);
            }
            else
            {
                // Recoverable crash - just post; don't block the caller's thread at all.
                Dispatcher.UIThread.Post(
                    () => _ = ShowCrashDialogOnUiThreadAsync(source, exception, logPath, isTerminating),
                    DispatcherPriority.MaxValue);
            }
        }
        catch
        {
            // Dispatcher invocation must never crash the app.
        }
    }

    private static async Task ShowCrashDialogOnUiThreadAsync(string source, Exception exception, string logPath, bool isTerminating)
    {
        // Signal "dialog is open" so the spin-wait loop in ShowCrashDialog
        // (terminating path) knows to keep the process alive.
        Interlocked.Exchange(ref _isCrashDialogOpen, 1);
        try
        {
            // Only use the main window as owner when it is still open and visible.
            // A closing or already-closed window causes ShowDialog to throw.
            Window? owner = null;
            if (Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            {
                var main = desktop.MainWindow;
                if (main is { IsVisible: true })
                    owner = main;
            }

            var dialog = BuildCrashDialog(source, exception, logPath, isTerminating, owner);

            // ShowDialog requires a non-null owner; fall back to Show() + TCS when none.
            if (owner is not null)
            {
                await dialog.ShowDialog(owner);
            }
            else
            {
                var tcs = new TaskCompletionSource<bool>();
                dialog.Closed += (_, _) => tcs.TrySetResult(true);
                dialog.Show();
                await tcs.Task;
            }
        }
        catch
        {
            // The crash dialog itself must never crash the app.
        }
        finally
        {
            // Signal "dialog dismissed" so the spin-wait loop exits.
            Interlocked.Exchange(ref _isCrashDialogOpen, 0);
        }
    }

    // Builds the crash dialog entirely in code so it has no AXAML dependency
    // and works even if App.axaml has not loaded yet.
    private static Window BuildCrashDialog(
        string  source,
        Exception exception,
        string  logPath,
        bool    isTerminating,
        Window? owner)
    {
        // --- Header ---
        var titleText = new TextBlock
        {
            Text        = "Kodo crashed",
            FontSize    = 16,
            FontWeight  = FontWeight.SemiBold,
            Foreground  = Brushes.White,
            TextWrapping = TextWrapping.Wrap,
        };

        var subtitleText = new TextBlock
        {
            Text = isTerminating
                ? "An unrecoverable error occurred and Kodo will now close. The crash details have been saved."
                : "An unexpected error occurred, but Kodo may still be running. The crash details have been saved.",
            FontSize    = 13,
            Foreground  = new SolidColorBrush(KodoTextMuted),
            TextWrapping = TextWrapping.Wrap,
            Margin      = new Thickness(0, 4, 0, 0),
        };

        // Terminating warning - only shown when IsTerminating is true so the user
        // knows the app is about to exit regardless of what they click.
        var terminatingBanner = new Border
        {
            IsVisible        = isTerminating,
            Background       = new SolidColorBrush(Color.Parse("#3D1A00")),
            BorderBrush      = new SolidColorBrush(Color.Parse("#7A3A00")),
            BorderThickness  = new Thickness(1),
            CornerRadius     = new CornerRadius(6),
            Padding          = new Thickness(10, 6),
            Child = new TextBlock
            {
                Text        = "⚠ The application will close after you dismiss this dialog.",
                FontSize    = 12,
                Foreground  = new SolidColorBrush(Color.Parse("#FFA040")),
                TextWrapping = TextWrapping.Wrap,
            },
        };

        // --- Source badge ---
        var sourceBadge = new Border
        {
            Background       = new SolidColorBrush(KodoDarkBadgeBg),
            BorderBrush      = new SolidColorBrush(KodoDarkBorder),
            BorderThickness  = new Thickness(1),
            CornerRadius     = new CornerRadius(6),
            Padding          = new Thickness(10, 5),
            HorizontalAlignment = HorizontalAlignment.Left,
            Child = new TextBlock
            {
                Text       = source,
                FontSize   = 12,
                FontFamily = new FontFamily("Cascadia Code,Consolas,Menlo,monospace"),
                Foreground = new SolidColorBrush(KodoTokenBlue),
            },
        };

        var metadataText = new SelectableTextBlock
        {
            Text         = KodoDiagnostics.BuildDiagnosticSummary(source, isTerminating),
            FontSize     = 11,
            FontFamily   = new FontFamily("Cascadia Code,Consolas,Menlo,monospace"),
            Foreground   = new SolidColorBrush(KodoTextMuted),
            TextWrapping = TextWrapping.Wrap,
        };

        // --- Exception details (scrollable, selectable) ---
        var exceptionText = new SelectableTextBlock
        {
            Text       = KodoDiagnostics.BuildDiagnosticPayload(source, exception, isTerminating, "Crash"),
            FontSize   = 12,
            FontFamily = new FontFamily("Cascadia Code,Consolas,Menlo,monospace"),
            Foreground = new SolidColorBrush(KodoTokenOrange),
            TextWrapping = TextWrapping.Wrap,
        };

        var exceptionScroll = new ScrollViewer
        {
            Content  = exceptionText,
            MaxHeight = 260,
            VerticalScrollBarVisibility   = Avalonia.Controls.Primitives.ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = Avalonia.Controls.Primitives.ScrollBarVisibility.Disabled,
        };

        var exceptionBorder = new Border
        {
            Background      = new SolidColorBrush(KodoDarkSurfaceDeep),
            BorderBrush     = new SolidColorBrush(KodoDarkBorder),
            BorderThickness = new Thickness(1),
            CornerRadius    = new CornerRadius(8),
            Padding         = new Thickness(12),
            Child           = exceptionScroll,
        };

        // --- Log path note ---
        var logPathText = new TextBlock
        {
            Text         = $"Full log written to: {logPath}",
            FontSize     = 11,
            Foreground   = new SolidColorBrush(KodoTextDim),
            TextWrapping = TextWrapping.Wrap,
        };

        // --- Action buttons ---
        var copyButton = new Button
        {
            Content             = "Copy to Clipboard",
            HorizontalAlignment = HorizontalAlignment.Left,
            Padding             = new Thickness(16, 8),
            Background          = new SolidColorBrush(KodoDarkBadgeBg),
            Foreground          = new SolidColorBrush(KodoTextMuted),
            BorderBrush         = new SolidColorBrush(KodoDarkBorder),
            BorderThickness     = new Thickness(1),
            CornerRadius        = new CornerRadius(8),
        };

        var (accentColor, accentForeground) = AccentResolver.GetCurrentAccent();

        var dismissButton = new Button
        {
            Content             = isTerminating ? "Close" : "Dismiss",
            HorizontalAlignment = HorizontalAlignment.Right,
            Padding             = new Thickness(20, 8),
            Background          = new SolidColorBrush(accentColor),
            Foreground          = new SolidColorBrush(accentForeground),
            BorderThickness     = new Thickness(0),
            CornerRadius        = new CornerRadius(8),
        };

        var buttonRow = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
        buttonRow.Children.Add(copyButton);
        Grid.SetColumn(dismissButton, 1);
        buttonRow.Children.Add(dismissButton);

        // --- Layout ---
        var content = new StackPanel
        {
            Spacing  = 12,
            Margin   = new Thickness(20),
            Children =
            {
                titleText,
                subtitleText,
                terminatingBanner,
                sourceBadge,
                metadataText,
                exceptionBorder,
                logPathText,
                buttonRow,
            },
        };

        var dialog = new Window
        {
            // "Crash Report" distinguishes this from the recoverable warning dialog.
            Title  = "Kodo - Crash Report",
            Width  = 560,
            SizeToContent = SizeToContent.Height,
            MinWidth  = 400,
            MinHeight = 200,
            MaxHeight = 740,
            CanResize = true,
            WindowStartupLocation = owner is not null
                ? WindowStartupLocation.CenterOwner
                : WindowStartupLocation.CenterScreen,
            Background = new SolidColorBrush(KodoDarkSurface),
            Content    = content,
        };

        // Copy the full exception + source to clipboard so the user can paste it
        // into a GitHub issue or Discord message without manually selecting text.
        copyButton.Click += async (_, _) =>
        {
            try
            {
                var clip = TopLevel.GetTopLevel(dialog)?.Clipboard;
                if (clip is not null)
                {
                    var text = KodoDiagnostics.BuildDiagnosticPayload(source, exception, isTerminating, "Crash");
                    await clip.SetTextAsync(text);
                    copyButton.Content  = "Copied!";
                    copyButton.Foreground = Brushes.White;
                }
            }
            catch
            {
                // Clipboard failures must not crash the crash dialog.
            }
        };

        dismissButton.Click += (_, _) => dialog.Close();

        return dialog;
    }

}