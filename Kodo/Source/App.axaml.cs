// Licensed under the Kodo Public License v1.1
using System.Diagnostics;
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

// Handles global unhandled-exception wiring, crash logging, and the crash dialog UI.
public partial class App : Application
{
    // Prevents crash-dialog storms: only one crash dialog may be in flight at a time.
    // Set by the CompareExchange guard in ShowCrashDialog, held for its full lifetime.
    // The terminating-exception spin-wait reads this flag too, to know when it's safe to proceed.
    private static int _isCrashDialogOpen;
    // Shared colours from DialogPalette, so every code-built dialog matches.

    private static readonly Color KodoDarkSurface     = DialogPalette.Surface;
    private static readonly Color KodoDarkSurfaceDeep = DialogPalette.SurfaceDeep;
    private static readonly Color KodoDarkBorder      = DialogPalette.Border;
    private static readonly Color KodoDarkBadgeBg     = DialogPalette.BadgeBg;
    private static readonly Color KodoTextMuted       = DialogPalette.TextMuted;
    private static readonly Color KodoTextDim         = DialogPalette.TextDim;
    private static readonly Color KodoTokenBlue       = DialogPalette.TokenBlue;  // source badge
    private static readonly Color KodoTokenOrange     = DialogPalette.TokenOrange;  // stack trace

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

    // Called once the framework finishes initializing; creates the main window.
    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            // desktop.Args[0] is the file path when launched via "Open with" / double-click.
            var startupFilePath = desktop.Args?.Length > 0 ? desktop.Args[0] : null;
            desktop.MainWindow  = new MainWindow(startupFilePath);

            AptabaseClient.TrackEvent("app_launched");
            desktop.Exit += async (_, _) => await AptabaseClient.FlushAsync();
        }

#if !DEBUG
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            RegisterFileAssociations();
        }
#endif
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            // No longer behind #if !DEBUG - that guard blocked the updater in Debug builds.
            if (!CheckPendingUpdateSentinel())
                CheckForUpdatesInBackground();
        }

        base.OnFrameworkInitializationCompleted();
    }

    // ── Auto-update ───────────────────────────────────────────────────────────

    // Checks for a pending-update sentinel from KodoUpdater; shows UpdateDialog if newer.
    // Runs before the live GitHub check so a pending sentinel always wins over a redundant download.
    private static bool CheckPendingUpdateSentinel()
    {
        try
        {
            var pending = PendingUpdateService.TryGetPendingUpdate();
            if (pending is null) return false;

            var (version, installerPath) = pending.Value;
            var update = new UpdateInfo(
                Version: version,
                ReleaseNotesUrl: $"https://github.com/SS-YYC/Kodo/releases",
                AssetDownloadUrl: string.Empty, // unused: installer is already on disk
                AssetName: Path.GetFileName(installerPath),
                AssetSizeBytes: 0);

            UpdateDialog.ShowFor(update, installerPath);
            return true;
        }
        catch
        {
            // Sentinel handling is best-effort; fall through to the normal
            // live update check below if anything here goes wrong.
            return false;
        }
    }

    // Fires a best-effort update check a few seconds after launch; failures are swallowed.
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
                if (update is null)
                    return;

                // "Update automatically without asking": skip the dialog
                // entirely and install silently in the background.
                if (UpdateService.IsAutoUpdateInBackgroundEnabledInSettings())
                    await UpdateService.SilentlyInstallAsync(update);
                else
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

        // Critical: unhandled AppDomain exception - may terminate the process.
        KodoDiagnostics.LogCritical("AppDomain.UnhandledException", exception, e.IsTerminating);
        ShowCrashDialog("AppDomain.UnhandledException", exception, isTerminating: e.IsTerminating);
    }

    private static void TaskScheduler_OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        // Unobserved Task exception - recoverable; logged as a Warning, no crash.log.
        KodoDiagnostics.LogWarning("TaskScheduler.UnobservedTaskException", e.Exception, operation: "Background task");
        // Mark as observed first so the runtime does not re-throw it after we return.
        e.SetObserved();
        // Show the crash-style dialog but marked as non-terminating so the user
        // knows the app is still alive and they can keep working.
        ShowCrashDialog("TaskScheduler.UnobservedTaskException", e.Exception, isTerminating: false);
    }

    private static void DispatcherUiThread_OnUnhandledException(object? sender, DispatcherUnhandledExceptionEventArgs e)
    {
        // Critical: exception on the UI thread - may leave the UI in a broken state.
        KodoDiagnostics.LogCritical("Dispatcher.UIThread.UnhandledException", e.Exception, isTerminating: false);
        ShowCrashDialog("Dispatcher.UIThread.UnhandledException", e.Exception, isTerminating: false);
        e.Handled = true;
    }

    // ShowCrashDialog dispatches a modal error dialog to the UI thread.

    private static void ShowCrashDialog(string source, Exception exception, bool isTerminating)
    {
        if (Interlocked.CompareExchange(ref _isCrashDialogOpen, 1, 0) != 0)
            return;

        try
        {
            var logPath = KodoDiagnostics.MainLogFilePath;

            if (Dispatcher.UIThread.CheckAccess())
            {
                // Already on the UI thread - fire and forget.
                _ = ShowCrashDialogOnUiThreadAsync(source, exception, logPath, isTerminating);
                return;
            }

            // Posts rather than blocking, since the UI thread may already be dead.
            // For terminating crashes, pumps the dispatcher here so the dialog can show, capped at 30s.
            if (isTerminating)
            {
                // Non-blocking post - returns immediately even if the UI thread is busy.
                Dispatcher.UIThread.Post(
                    () => _ = ShowCrashDialogOnUiThreadAsync(source, exception, logPath, isTerminating),
                    DispatcherPriority.MaxValue);

                // Gives the UI thread up to 30s to show/dismiss the dialog before teardown.
                for (var i = 0; i < 300 && _isCrashDialogOpen == 1; i++)
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
            Text       = KodoDiagnostics.BuildDiagnosticPayload(source, exception, isTerminating, KodoSeverity.Critical),
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
            Text         = $"Full details in: {logPath}",
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

        var reportButton = new Button
        {
            Content             = "Report on GitHub",
            HorizontalAlignment = HorizontalAlignment.Left,
            Padding             = new Thickness(16, 8),
            Background          = new SolidColorBrush(KodoDarkBadgeBg),
            Foreground          = new SolidColorBrush(KodoTextMuted),
            BorderBrush         = new SolidColorBrush(KodoDarkBorder),
            BorderThickness     = new Thickness(1),
            CornerRadius        = new CornerRadius(8),
            Margin              = new Thickness(8, 0, 0, 0),
        };

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

        // Left side: Copy + Report. Right side: Dismiss.
        var leftButtons = new StackPanel
        {
            Orientation = Avalonia.Layout.Orientation.Horizontal,
            Children    = { copyButton, reportButton },
        };

        var buttonRow = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
        buttonRow.Children.Add(leftButtons);
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
                    var text = KodoDiagnostics.BuildDiagnosticPayload(source, exception, isTerminating, KodoSeverity.Critical);
                    await clip.SetTextAsync(text);
                    copyButton.Content   = "Copied!";
                    copyButton.Foreground = Brushes.White;
                }
            }
            catch
            {
                // Clipboard failures must not crash the crash dialog.
            }
        };

        reportButton.Click += (_, _) =>
        {
            try
            {
                // Pre-fill a GitHub issue with the exception type as the title.
                // The user can paste the clipboard payload into the body.
                var title = Uri.EscapeDataString($"[Crash] {exception.GetType().Name}: {exception.Message}"
                    .Replace("\r", "").Replace("\n", " ").Trim());
                var url = $"https://github.com/KerbalMissile/Kodo/issues/new?title={title}" +
                          $"&labels=bug&template=bug_report.md";
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            }
            catch
            {
                // Opening the browser must not crash the crash dialog.
            }
        };

        dismissButton.Click += (_, _) => dialog.Close();

        return dialog;
    }

}
