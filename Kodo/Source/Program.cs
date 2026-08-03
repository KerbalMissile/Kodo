// Licensed under the GPL-v3.0
using Avalonia;
using System;
using System.IO;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
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

        // If Kodo is already running, hand this launch's file (e.g. from "Open with"
        // / double-click in Explorer) off to that instance as a new tab and exit -
        // don't spin up a second window.
        if (!SingleInstance.TryAcquire())
        {
            var handoffPath = args.Length > 0 ? args[0] : null;
            SingleInstance.SendActivationRequest(handoffPath);
            return;
        }

        AptabaseClient.Initialize();

        // Unhandled exceptions are handled in App.Initialize()
        var app = BuildAvaloniaApp();

        try
        {
            app.StartWithClassicDesktopLifetime(args);
        }
        finally
        {
            Task.Run(async () => await AptabaseClient.FlushAsync()).Wait(TimeSpan.FromSeconds(2));
            SingleInstance.Release();
        }
    }

    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}

// Ensures only one Kodo window exists per user session. The first process to
// start owns the mutex and becomes "primary" - it opens a named pipe and
// listens for file paths handed off by later launches (e.g. double-clicking
// another file in Explorer while Kodo is already open). Any later launch
// fails to acquire the mutex, forwards its startup file path to the primary
// instance over the pipe, and exits without ever creating a window.
internal static class SingleInstance
{
    private const string MutexName = @"Local\Kodo_SingleInstance_Mutex_9F3E2C1A";
    private const string PipeName  = "Kodo_SingleInstance_Pipe_9F3E2C1A";

    private static Mutex? _mutex;

    // True if this process is the primary (first) instance and owns the mutex.
    public static bool TryAcquire()
    {
        _mutex = new Mutex(initiallyOwned: true, MutexName, out var createdNew);
        if (createdNew) return true;

        // Didn't create it - someone else owns it. Don't hold a handle around.
        _mutex.Dispose();
        _mutex = null;
        return false;
    }

    public static void Release()
    {
        try
        {
            _mutex?.ReleaseMutex();
            _mutex?.Dispose();
        }
        catch { /* best effort on shutdown */ }
    }

    // Called by the primary instance once its MainWindow exists. Listens forever
    // (until the process exits) for activation requests from later launches.
    public static void StartListening(Action<string?> onActivationRequested)
    {
        _ = Task.Run(async () =>
        {
            while (true)
            {
                try
                {
                    using var server = new NamedPipeServerStream(
                        PipeName,
                        PipeDirection.In,
                        maxNumberOfServerInstances: 1,
                        PipeTransmissionMode.Byte,
                        PipeOptions.Asynchronous);

                    await server.WaitForConnectionAsync();

                    using var reader = new StreamReader(server, Encoding.UTF8, leaveOpen: true);
                    var payload = await reader.ReadToEndAsync();

                    // Empty payload just means "bring the window to front", no file to open.
                    onActivationRequested(string.IsNullOrWhiteSpace(payload) ? null : payload);
                }
                catch
                {
                    // Pipe hiccup - brief backoff so a persistent failure doesn't spin the loop.
                    await Task.Delay(500);
                }
            }
            // ReSharper disable once FunctionNeverReturns
        });
    }

    // Called by a secondary launch. Hands the startup file path (if any) to the
    // already-running primary instance so it can open it as a tab and come to front.
    public static void SendActivationRequest(string? filePath)
    {
        try
        {
            using var client = new NamedPipeClientStream(".", PipeName, PipeDirection.Out);
            client.Connect(timeout: 2000);

            using var writer = new StreamWriter(client, Encoding.UTF8);
            writer.Write(filePath ?? string.Empty);
            writer.Flush();
        }
        catch
        {
            // If the primary instance is unreachable (e.g. it's shutting down right
            // now), there's nothing useful to do - the user can just relaunch Kodo.
        }
    }
}