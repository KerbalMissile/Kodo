using System;
using System.Runtime.InteropServices;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Platform;
using Avalonia.Platform;

namespace Kodo;

public sealed class EmbeddedTerminalHost : NativeControlHost
{
    private IntPtr _hostHandle;

    public event EventHandler? HostHandleReady;
    public event EventHandler? HostBoundsChanged;

    public IntPtr HostHandle => _hostHandle;

    public EmbeddedTerminalHost()
    {
        AttachedToVisualTree += (_, _) => HostBoundsChanged?.Invoke(this, EventArgs.Empty);
    }

    protected override Size ArrangeOverride(Size finalSize)
    {
        var result = base.ArrangeOverride(finalSize);
        HostBoundsChanged?.Invoke(this, EventArgs.Empty);
        return result;
    }

    protected override IPlatformHandle CreateNativeControlCore(IPlatformHandle parent)
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return base.CreateNativeControlCore(parent);

        var hwnd = NativeMethods.CreateWindowExW(
            0,
            "static",
            string.Empty,
            NativeMethods.WS_CHILD | NativeMethods.WS_VISIBLE | NativeMethods.WS_CLIPCHILDREN | NativeMethods.WS_CLIPSIBLINGS,
            0, 0, 0, 0,
            parent.Handle,
            IntPtr.Zero,
            IntPtr.Zero,
            IntPtr.Zero);

        _hostHandle = hwnd;
        HostHandleReady?.Invoke(this, EventArgs.Empty);
        return new PlatformHandle(hwnd, "HWND");
    }

    protected override void DestroyNativeControlCore(IPlatformHandle control)
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) && control.Handle != IntPtr.Zero)
            NativeMethods.DestroyWindow(control.Handle);

        _hostHandle = IntPtr.Zero;
        base.DestroyNativeControlCore(control);
    }

    public PixelSize GetHostClientSize()
    {
        if (_hostHandle == IntPtr.Zero || !RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return default;

        NativeMethods.GetClientRect(_hostHandle, out var rect);
        return new PixelSize(Math.Max(0, rect.Right - rect.Left), Math.Max(0, rect.Bottom - rect.Top));
    }

    private static class NativeMethods
    {
        public const int WS_CHILD = 0x40000000;
        public const int WS_VISIBLE = 0x10000000;
        public const int WS_CLIPSIBLINGS = 0x04000000;
        public const int WS_CLIPCHILDREN = 0x02000000;

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll", EntryPoint = "CreateWindowExW", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr CreateWindowExW(
            int exStyle,
            string className,
            string windowName,
            int style,
            int x,
            int y,
            int width,
            int height,
            IntPtr parentHandle,
            IntPtr menuHandle,
            IntPtr instanceHandle,
            IntPtr param);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DestroyWindow(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetClientRect(IntPtr hwnd, out RECT rect);
    }
}
