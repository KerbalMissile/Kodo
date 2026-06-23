using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Threading;

namespace Kodo;

// ── Public surface ────────────────────────────────────────────────────────────

/// <summary>
/// A self-contained Avalonia control that hosts a shell via Windows
/// Pseudo Console (ConPTY) and renders its VT/ANSI output directly.
/// Drop it into AXAML exactly where EmbeddedTerminalHost used to live.
/// </summary>
public sealed class ConsoleTerminal : Control
{
    // ── Standard 16-colour ANSI palette ──────────────────────────────────────
    private static readonly Color[] AnsiPalette =
    [
        Color.FromRgb(12,  12,  12),   // 0  Black
        Color.FromRgb(197, 15,  31),   // 1  Red
        Color.FromRgb(19,  161, 14),   // 2  Green
        Color.FromRgb(193, 156, 0),    // 3  Yellow
        Color.FromRgb(0,   55,  218),  // 4  Blue
        Color.FromRgb(136, 23,  152),  // 5  Magenta
        Color.FromRgb(58,  150, 221),  // 6  Cyan
        Color.FromRgb(204, 204, 204),  // 7  White
        Color.FromRgb(118, 118, 118),  // 8  Bright Black
        Color.FromRgb(231, 72,  86),   // 9  Bright Red
        Color.FromRgb(22,  198, 12),   // 10 Bright Green
        Color.FromRgb(249, 241, 165),  // 11 Bright Yellow
        Color.FromRgb(59,  120, 255),  // 12 Bright Blue
        Color.FromRgb(180, 0,   158),  // 13 Bright Magenta
        Color.FromRgb(97,  214, 214),  // 14 Bright Cyan
        Color.FromRgb(242, 242, 242),  // 15 Bright White
    ];

    private static readonly Color DefaultFg = Color.FromRgb(204, 204, 204);
    private static readonly Color DefaultBg = Color.FromRgb(12,  12,  12);

    // ── Layout ────────────────────────────────────────────────────────────────
    private const double CellW = 8.4;
    private const double CellH = 17.0;
    private const string FontFamily = "Cascadia Mono,Consolas,Courier New,monospace";
    private const double FontSize = 13.0;

    // ── State ─────────────────────────────────────────────────────────────────
    private readonly object _lock = new();
    private TermCell[,] _cells = new TermCell[24, 80];
    private int _rows = 24, _cols = 80;
    private int _cursorRow, _cursorCol;
    private bool _cursorVisible = true;

    private Color _fg = DefaultFg, _bg = DefaultBg;
    private bool _bold, _underline, _reverse;

    // VT parser state
    public enum ParseState { Ground, Escape, CsiEntry, CsiParam, CsiIgnore, OscString }
    private ParseState _parseState = ParseState.Ground;
    private readonly StringBuilder _csiParam = new();
    private readonly StringBuilder _oscBuf   = new();

    // ConPTY
    private IntPtr  _hPcon   = IntPtr.Zero;
    private IntPtr  _hProcess = IntPtr.Zero;
    private IntPtr  _hThread  = IntPtr.Zero;
    private Stream? _writeStream;
    private Stream? _readStream;
    private CancellationTokenSource? _cts;

    // When non-zero, the read loop discards all shell output until this
    // timestamp (from Environment.TickCount64) has passed. Set by Start()
    // when a snapshot restore is pending; keeps discarding through the shell's
    // entire startup sequence (including the CSI 2 J full-screen clear that
    // PowerShell/cmd emit as part of drawing their first prompt), so the
    // snapshot remains visible until the shell settles and redraws normally.
    private long _suppressOutputUntilTick;

    // Blink timer for cursor
    private readonly DispatcherTimer _blinkTimer;
    private bool _cursorBlinkOn = true;

    // ── Constructor ───────────────────────────────────────────────────────────
    public ConsoleTerminal()
    {
        Focusable = true;
        ClipToBounds = true;

        _blinkTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(530) };
        _blinkTimer.Tick += (_, _) => { _cursorBlinkOn = !_cursorBlinkOn; InvalidateVisual(); };
        _blinkTimer.Start();

        AttachedToVisualTree   += (_, _) => Focus();
        // NOTE: Do NOT call Stop() on DetachedFromVisualTree. In Avalonia,
        // toggling IsVisible on a parent Grid causes children to detach and
        // re-attach from the visual tree, which would kill the ConPTY process
        // every time the terminal panel is shown or hidden. The process is
        // stopped explicitly by the MainWindow (via Stop() / ActiveTerminalSession
        // setter) when a session is actually closed or the window is shutting down.
    }

    // ── Public API ────────────────────────────────────────────────────────────

    /// <summary>
    /// Fired when the shell process exits. The event argument is the process
    /// handle (<c>_hProcess</c>) that was running when <see cref="Start"/> was
    /// called. Subscribers use it to verify the exit belongs to the process they
    /// started, not a stale wake-up from a previous <see cref="Stop"/> call.
    /// </summary>
    public event EventHandler<IntPtr>? SessionExited;

    /// <summary>Start a shell. Stops any previously running session first.</summary>
    /// <param name="suppressOutputUntilRestored">
    /// Pass <c>true</c> when the caller is about to call <see cref="RestoreSnapshot"/>
    /// immediately after this returns. The read loop will discard shell output until
    /// <see cref="RestoreSnapshot"/> clears the flag, preventing the shell's early
    /// prompt from racing against and overwriting the restored buffer.
    /// Leave <c>false</c> (default) for brand-new sessions with no saved buffer.
    /// </param>
    public void Start(string shellPath, string arguments, string workingDirectory,
                      bool suppressOutputUntilRestored = false)
    {
        Stop();
        _cts = new CancellationTokenSource();

        try
        {
            var (cols, rows) = CalcSize();
            _cols = cols; _rows = rows;
            // Only allocate a fresh cell grid when we are NOT about to restore a
            // snapshot. If a restore is pending, RestoreSnapshot() will write the
            // grid contents and the dimensions are already correct from the last
            // ArrangeOverride. Calling ResizeCells here with a potentially stale
            // Bounds.Size would wipe the grid to the wrong size, causing the
            // subsequent ArrangeOverride (which uses the real layout size) to see
            // a dimension mismatch and wipe it again.
            if (!suppressOutputUntilRestored)
                ResizeCells(rows, cols);

            // Create pipe pairs for ConPTY I/O
            NativeConPty.CreatePipe(out var hReadPtyOutput, out var hWritePtyOutput, IntPtr.Zero, 0);
            NativeConPty.CreatePipe(out var hReadPtyInput,  out var hWritePtyInput,  IntPtr.Zero, 0);

            // Create the Pseudo Console
            var result = NativeConPty.CreatePseudoConsole(
                new NativeConPty.COORD { X = (short)cols, Y = (short)rows },
                hReadPtyInput, hWritePtyOutput,
                0, out _hPcon);

            if (result != 0)
                throw new InvalidOperationException($"CreatePseudoConsole failed: 0x{result:X}");

            // Close the handles we handed to ConPTY
            NativeConPty.CloseHandle(hReadPtyInput);
            NativeConPty.CloseHandle(hWritePtyOutput);

            // Wrap native handles in .NET streams
            _writeStream = new FileStream(
                new Microsoft.Win32.SafeHandles.SafeFileHandle(hWritePtyInput,  true), FileAccess.Write);
            _readStream  = new FileStream(
                new Microsoft.Win32.SafeHandles.SafeFileHandle(hReadPtyOutput, true), FileAccess.Read);

            // Launch the shell inside the ConPTY
            var cmdLine = new StringBuilder($"\"{shellPath}\" {arguments}");
            NativeConPty.LaunchProcess(cmdLine, workingDirectory, _hPcon,
                out _hProcess, out _hThread);

            // If the caller is about to restore a snapshot, suppress output for
            // 500 ms - long enough for the shell's full startup sequence (including
            // the CSI 2 J screen-clear that PowerShell/cmd/bash all send while
            // drawing their first prompt) to pass without wiping the snapshot.
            _suppressOutputUntilTick = suppressOutputUntilRestored
                ? Environment.TickCount64 + 500
                : 0;

            // Start reading output
            _ = Task.Run(() => ReadOutputLoop(_cts.Token), _cts.Token);

            // Watch for process exit. Capture the handle into a local so the
            // background task holds the exact value that was alive at Start() time.
            // The UI-thread post passes it as the event argument so the subscriber
            // can reject stale wake-ups that belong to a since-stopped process.
            var watchedHandle = _hProcess;
            _ = Task.Run(() =>
            {
                NativeConPty.WaitForSingleObject(watchedHandle, 0xFFFFFFFF);
                Dispatcher.UIThread.Post(() =>
                {
                    SessionExited?.Invoke(this, watchedHandle);
                });
            }, _cts.Token);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[ConPTY] Start failed: {ex}");
        }
    }

    public void Stop()
    {
        _suppressOutputUntilTick = 0;
        _cts?.Cancel();
        _cts = null;

        try { _writeStream?.Dispose(); } catch { }
        try { _readStream?.Dispose();  } catch { }
        _writeStream = null;
        _readStream  = null;

        if (_hPcon != IntPtr.Zero)
        {
            NativeConPty.ClosePseudoConsole(_hPcon);
            _hPcon = IntPtr.Zero;
        }
        if (_hProcess != IntPtr.Zero) { NativeConPty.CloseHandle(_hProcess); _hProcess = IntPtr.Zero; }
        if (_hThread  != IntPtr.Zero) { NativeConPty.CloseHandle(_hThread);  _hThread  = IntPtr.Zero; }
    }

    public void Resize(int cols, int rows)
    {
        if (cols <= 0 || rows <= 0 || (cols == _cols && rows == _rows))
            return;

        _cols = cols; _rows = rows;
        ResizeCells(rows, cols);

        if (_hPcon != IntPtr.Zero)
            NativeConPty.ResizePseudoConsole(_hPcon,
                new NativeConPty.COORD { X = (short)cols, Y = (short)rows });
    }

    public void SendInput(string text)
    {
        if (_writeStream is null || string.IsNullOrEmpty(text)) return;
        var bytes = Encoding.UTF8.GetBytes(text);
        try { _writeStream.Write(bytes, 0, bytes.Length); _writeStream.Flush(); }
        catch (Exception ex) { Console.WriteLine($"[ConPTY] SendInput failed: {ex.Message}"); }
    }

    public void SendKey(Key key, KeyModifiers mods)
    {
        var seq = KeyToVt(key, mods);
        if (seq is not null) SendInput(seq);
    }

    /// <summary>True while a ConPTY process is running in this control.</summary>
    public bool HasLiveProcess => _hPcon != IntPtr.Zero;

    /// <summary>
    /// The process handle that was returned by the most recent successful
    /// <see cref="Start"/> call. Used by the exit handler in MainWindow to
    /// verify that a <see cref="SessionExited"/> post belongs to the process
    /// it subscribed for, not a stale wake-up from a previous Stop().
    /// Zero when no process is running.
    /// </summary>
    public IntPtr CurrentProcessHandle => _hProcess;

    /// <summary>
    /// Captures the current screen buffer and cursor state so it can be
    /// restored later when switching back to this session.
    /// </summary>
    public TerminalSnapshot SaveSnapshot()
    {
        lock (_lock)
        {
            var cells = (TermCell[,])_cells.Clone();
            return new TerminalSnapshot(
                cells, _rows, _cols,
                _cursorRow, _cursorCol, _cursorVisible,
                _fg, _bg, _bold, _underline, _reverse,
                _parseState, _csiParam.ToString());
        }
    }

    /// <summary>
    /// Replaces the current screen buffer with a previously saved snapshot,
    /// preserving the running ConPTY process. Called after switching sessions
    /// so the incoming session's output is immediately visible.
    /// </summary>
    public void RestoreSnapshot(TerminalSnapshot snap)
    {
        lock (_lock)
        {
            // Restore cell content into the *current* grid dimensions rather than
            // the snapshot's saved dimensions. This is critical: ArrangeOverride
            // fires a layout pass shortly after this call and invokes Resize() with
            // the current control size. If _rows/_cols already match the control
            // size, Resize() is a no-op and the restored content survives. If we
            // instead restored the snapshot's (possibly different) dimensions,
            // the next Resize() would allocate a new grid and zero out everything.
            var copyR = Math.Min(_rows, snap.Rows);
            var copyC = Math.Min(_cols, snap.Cols);
            var newCells = new TermCell[_rows, _cols];
            for (var r = 0; r < copyR; r++)
            for (var c = 0; c < copyC; c++)
                newCells[r, c] = snap.Cells[r, c];
            _cells = newCells;

            _cursorRow     = Math.Min(snap.CursorRow, _rows - 1);
            _cursorCol     = Math.Min(snap.CursorCol, _cols - 1);
            _cursorVisible = snap.CursorVisible;
            _fg            = snap.Fg;
            _bg            = snap.Bg;
            _bold          = snap.Bold;
            _underline     = snap.Underline;
            _reverse       = snap.Reverse;
            _parseState    = snap.ParseState;
            _csiParam.Clear();
            _csiParam.Append(snap.CsiParam);

            // The suppression window set in Start() keeps the shell's startup
            // sequence (including its CSI 2 J screen-clear) from wiping the
            // snapshot. No need to clear it here - it will expire on its own
            // after 500 ms, at which point the shell's redrawn prompt takes over.
        }
        // Redraw immediately with the restored content.
        InvalidateVisual();
    }

    // ── Avalonia overrides ────────────────────────────────────────────────────

    protected override void OnKeyDown(KeyEventArgs e)
    {
        // Handle everything ourselves before calling base so that window-level
        // tunnel handlers (editor shortcuts, etc.) cannot swallow terminal keys.

        // 1. Special keys (arrows, F-keys, Home/End, …) - these never produce a
        //    TextInput event, so we must intercept them here.
        var seq = KeyToVt(e.Key, e.KeyModifiers);
        if (seq is not null)
        {
            SendInput(seq);
            e.Handled = true;
            base.OnKeyDown(e);
            return;
        }

        // 2. Ctrl+letter → control character (Ctrl+C = 3, Ctrl+D = 4, …).
        //    Mark handled so the window shortcuts (Ctrl+V, Ctrl+C, …) don't fire.
        if (e.KeyModifiers.HasFlag(KeyModifiers.Control) &&
            !e.KeyModifiers.HasFlag(KeyModifiers.Alt))
        {
            var ch = ControlChar(e.Key);
            if (ch.HasValue)
            {
                SendInput(((char)ch.Value).ToString());
                e.Handled = true;
                base.OnKeyDown(e);
                return;
            }
        }

        // 3. Alt+key → ESC-prefix the character so readline / vim get Alt sequences.
        //    Only do this when it won't produce a TextInput event with the right text
        //    (i.e. non-printable Alt combos like Alt+F, Alt+B used by readline).
        if (e.KeyModifiers == KeyModifiers.Alt)
        {
            var altChar = AltKeyChar(e.Key);
            if (altChar is not null)
            {
                SendInput("\x1b" + altChar);
                e.Handled = true;
                base.OnKeyDown(e);
                return;
            }
        }

        // 4. Everything else (plain printable chars, Shift+letter, AltGr combos)
        //    arrives via OnTextInput - nothing to do here except let base run.
        base.OnKeyDown(e);
    }

    protected override void OnTextInput(TextInputEventArgs e)
    {
        base.OnTextInput(e);
        if (!string.IsNullOrEmpty(e.Text)) { SendInput(e.Text); e.Handled = true; }
    }

    protected override Size MeasureOverride(Size available) => available;

    protected override Size ArrangeOverride(Size final)
    {
        // Skip resizing the cell grid while a snapshot restore is pending.
        // The suppression window covers the shell's full startup sequence so
        // a layout pass can't wipe the restored content before it's visible.
        if (Environment.TickCount64 >= _suppressOutputUntilTick)
        {
            var (cols, rows) = CalcSize(final);
            Resize(cols, rows);
        }
        return final;
    }

    protected override void OnPointerPressed(PointerPressedEventArgs e)
    {
        base.OnPointerPressed(e);
        Focus();
    }

    // Cached typefaces - reusing the same object avoids per-call font lookup overhead
    // and ensures consistent glyph metrics across all cells in a frame.
    private static readonly Typeface TypefaceNormal = new(FontFamily, FontStyle.Normal, FontWeight.Regular);
    private static readonly Typeface TypefaceBold   = new(FontFamily, FontStyle.Normal, FontWeight.Bold);

    public override void Render(DrawingContext ctx)
    {
        lock (_lock)
        {
            // Fill the entire control with the default background first so that
            // any cells we don't explicitly paint are correctly erased.
            ctx.FillRectangle(new SolidColorBrush(DefaultBg), new Rect(Bounds.Size));

            var isCursorVisible = _cursorVisible && _cursorBlinkOn;

            for (var r = 0; r < _rows; r++)
            for (var c = 0; c < _cols; c++)
            {
                var cell = _cells[r, c];

                // Use integer pixel positions so sub-pixel accumulation from
                // fractional CellW (8.4 px) never causes glyph drift or overlap.
                var x = (int)Math.Round(c * CellW);
                var x1 = (int)Math.Round((c + 1) * CellW);
                var y = r * CellH;          // CellH is already integral (17.0)
                var w = x1 - x;            // actual pixel width for this column
                var rect = new Rect(x, y, w, CellH);

                var atCursor = isCursorVisible && r == _cursorRow && c == _cursorCol;

                // Always paint the cell background so that when the cursor moves
                // away, the previous cursor cell is fully restored to its real bg.
                var bg = atCursor ? DefaultFg : (cell.Bg ?? DefaultBg);
                if (bg != DefaultBg)
                    ctx.FillRectangle(new SolidColorBrush(bg), rect);

                // Glyph
                if (cell.Char != '\0' && cell.Char != ' ')
                {
                    var fg = atCursor ? DefaultBg : (cell.Fg ?? DefaultFg);
                    var typeface = cell.Bold ? TypefaceBold : TypefaceNormal;

                    var ft = new FormattedText(
                        cell.Char.ToString(),
                        System.Globalization.CultureInfo.InvariantCulture,
                        FlowDirection.LeftToRight,
                        typeface,
                        FontSize,
                        new SolidColorBrush(fg));

                    ctx.DrawText(ft, new Point(x, y));
                }

                // Underline
                if (cell.Underline)
                {
                    var fg = cell.Fg ?? DefaultFg;
                    ctx.DrawLine(new Pen(new SolidColorBrush(fg)),
                        new Point(x, y + CellH - 2), new Point(x + w, y + CellH - 2));
                }
            }
        }
    }

    // ── Private helpers ───────────────────────────────────────────────────────

    private (int cols, int rows) CalcSize(Size? size = null)
    {
        var s = size ?? Bounds.Size;
        // Count how many whole columns fit using the same rounding used in Render
        // (x = round(c * CellW)) so the column count matches the painted grid exactly.
        var cols = Math.Max(10, (int)(s.Width / CellW));
        var rows = Math.Max(3,  (int)(s.Height / CellH));
        return (cols, rows);
    }

    private void ResizeCells(int rows, int cols)
    {
        lock (_lock)
        {
            var next = new TermCell[rows, cols];
            var copyR = Math.Min(rows, _cells.GetLength(0));
            var copyC = Math.Min(cols, _cells.GetLength(1));
            for (var r = 0; r < copyR; r++)
            for (var c = 0; c < copyC; c++)
                next[r, c] = _cells[r, c];
            _cells = next;
            _cursorRow = Math.Min(_cursorRow, rows - 1);
            _cursorCol = Math.Min(_cursorCol, cols - 1);
        }
    }

    // ── Output reader ─────────────────────────────────────────────────────────
    private async Task ReadOutputLoop(CancellationToken ct)
    {
        var buf = new byte[4096];
        try
        {
            while (!ct.IsCancellationRequested)
            {
                var n = await _readStream!.ReadAsync(buf, 0, buf.Length, ct);
                if (n <= 0) break;
                var text = Encoding.UTF8.GetString(buf, 0, n);
                lock (_lock)
                {
                    // Drain the pipe even while suppressed so the shell doesn't
                    // block, but discard bytes until the suppression window expires.
                    // This covers the shell's full startup sequence including the
                    // CSI 2 J screen-clear that most shells emit while drawing their
                    // first prompt - which would otherwise wipe the restored snapshot.
                    if (Environment.TickCount64 >= _suppressOutputUntilTick)
                        foreach (var ch in text) ProcessChar(ch);
                }
                Dispatcher.UIThread.Post(InvalidateVisual, DispatcherPriority.Render);
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex) { Console.WriteLine($"[ConPTY] ReadOutputLoop: {ex.Message}"); }
    }

    // ── VT / ANSI parser ──────────────────────────────────────────────────────
    private void ProcessChar(char ch)
    {
        switch (_parseState)
        {
            case ParseState.Ground:
                switch (ch)
                {
                    case '\x1B': _parseState = ParseState.Escape; break;
                    case '\r':   _cursorCol = 0; break;
                    case '\n':   LineFeed(); break;
                    case '\b':   if (_cursorCol > 0) _cursorCol--; break;
                    case '\t':   _cursorCol = Math.Min(_cols - 1, (_cursorCol / 8 + 1) * 8); break;
                    case '\a':   break; // bell - ignore
                    case '\x0E': break; // SO (shift-out) - ignore, no alternate charset
                    case '\x0F': break; // SI (shift-in)  - ignore
                    default:
                        if (ch >= ' ')
                        {
                            SetCell(_cursorRow, _cursorCol, ch);
                            _cursorCol++;
                            if (_cursorCol >= _cols) { _cursorCol = 0; LineFeed(); }
                        }
                        break;
                }
                break;

            case ParseState.Escape:
                switch (ch)
                {
                    case '[': _csiParam.Clear(); _parseState = ParseState.CsiEntry; break;
                    case ']': _oscBuf.Clear();   _parseState = ParseState.OscString; break;
                    case 'M': ReverseLineFeed(); _parseState = ParseState.Ground; break;
                    case 'c': ResetTerminal();   _parseState = ParseState.Ground; break;
                    default:                     _parseState = ParseState.Ground; break;
                }
                break;

            case ParseState.CsiEntry:
            case ParseState.CsiParam:
                if (ch == '?' || ch == '>' || ch == '!')
                {
                    _csiParam.Append(ch);
                    _parseState = ParseState.CsiParam;
                }
                else if (ch >= '0' && ch <= '9' || ch == ';')
                {
                    _csiParam.Append(ch);
                    _parseState = ParseState.CsiParam;
                }
                else if (ch >= 0x40 && ch <= 0x7E)
                {
                    DispatchCsi(ch, _csiParam.ToString());
                    _parseState = ParseState.Ground;
                }
                else
                {
                    _parseState = ParseState.CsiIgnore;
                }
                break;

            case ParseState.CsiIgnore:
                if (ch >= 0x40 && ch <= 0x7E) _parseState = ParseState.Ground;
                break;

            case ParseState.OscString:
                if (ch == '\x07' || ch == '\x9C') _parseState = ParseState.Ground;
                else if (ch == '\x1B') _parseState = ParseState.Escape; // ESC \ = ST
                break;
        }
    }

    private void DispatchCsi(char cmd, string param)
    {
        var priv = param.StartsWith('?');
        var raw  = priv ? param[1..] : param;
        var nums = ParseNums(raw);
        int P(int i, int def = 1) => (i < nums.Count && nums[i] > 0) ? nums[i] : def;
        int P0(int i)             => (i < nums.Count) ? nums[i] : 0;

        switch (cmd)
        {
            // Cursor movement
            case 'A': _cursorRow = Math.Max(0, _cursorRow - P(0)); break;
            case 'B': _cursorRow = Math.Min(_rows - 1, _cursorRow + P(0)); break;
            case 'C': _cursorCol = Math.Min(_cols - 1, _cursorCol + P(0)); break;
            case 'D': _cursorCol = Math.Max(0, _cursorCol - P(0)); break;
            case 'E': _cursorRow = Math.Min(_rows - 1, _cursorRow + P(0)); _cursorCol = 0; break;
            case 'F': _cursorRow = Math.Max(0, _cursorRow - P(0));          _cursorCol = 0; break;
            case 'G': _cursorCol = Math.Min(_cols - 1, Math.Max(0, P(0) - 1)); break;
            case 'H': case 'f':
                _cursorRow = Math.Min(_rows - 1, Math.Max(0, P(0) - 1));
                _cursorCol = Math.Min(_cols - 1, Math.Max(0, P(1) - 1));
                break;
            case 'd': _cursorRow = Math.Min(_rows - 1, Math.Max(0, P(0) - 1)); break;

            // Erase
            case 'J': EraseDisplay(P0(0)); break;
            case 'K': EraseLine(P0(0)); break;

            // SGR - colours and attributes
            case 'm': ApplySgr(nums); break;

            // Cursor visibility
            case 'h': if (priv && P0(0) == 25) _cursorVisible = true;  break;
            case 'l': if (priv && P0(0) == 25) _cursorVisible = false; break;

            // Insert / delete lines
            case 'L': InsertLines(P(0)); break;
            case 'M': DeleteLines(P(0)); break;

            // Insert / delete / erase chars
            case '@': InsertChars(P(0)); break;
            case 'P': DeleteChars(P(0)); break;
            // CSI <n> X - Erase Character: blank n cells at cursor without moving it.
            // PSReadLine uses this to clear the tail of a longer previous command when
            // a shorter history entry replaces it (e.g. after pressing Up arrow).
            case 'X': EraseChars(P(0)); break;

            // Scroll
            case 'S': ScrollUp(P(0));   break;
            case 'T': ScrollDown(P(0)); break;
        }
    }

    private void ApplySgr(List<int> nums)
    {
        if (nums.Count == 0) { ResetAttrs(); return; }
        for (var i = 0; i < nums.Count; i++)
        {
            switch (nums[i])
            {
                case 0:  ResetAttrs(); break;
                case 1:  _bold = true; break;
                case 4:  _underline = true; break;
                case 7:  _reverse = true; break;
                case 22: _bold = false; break;
                case 24: _underline = false; break;
                case 27: _reverse = false; break;
                case >= 30 and <= 37: _fg = AnsiPalette[nums[i] - 30]; break;
                case 38:
                    if (i + 2 < nums.Count && nums[i+1] == 5) { _fg = Palette256(nums[i+2]); i += 2; }
                    else if (i + 4 < nums.Count && nums[i+1] == 2) { _fg = Color.FromRgb((byte)nums[i+2], (byte)nums[i+3], (byte)nums[i+4]); i += 4; }
                    break;
                case 39: _fg = DefaultFg; break;
                case >= 40 and <= 47: _bg = AnsiPalette[nums[i] - 40]; break;
                case 48:
                    if (i + 2 < nums.Count && nums[i+1] == 5) { _bg = Palette256(nums[i+2]); i += 2; }
                    else if (i + 4 < nums.Count && nums[i+1] == 2) { _bg = Color.FromRgb((byte)nums[i+2], (byte)nums[i+3], (byte)nums[i+4]); i += 4; }
                    break;
                case 49: _bg = DefaultBg; break;
                case >= 90 and <= 97:  _fg = AnsiPalette[nums[i] - 90 + 8]; break;
                case >= 100 and <= 107: _bg = AnsiPalette[nums[i] - 100 + 8]; break;
            }
        }
    }

    private void ResetAttrs() { _fg = DefaultFg; _bg = DefaultBg; _bold = false; _underline = false; _reverse = false; }
    private void ResetTerminal() { ResetAttrs(); _cursorRow = 0; _cursorCol = 0; ClearAllCells(); }

    private void SetCell(int r, int c, char ch)
    {
        if (r < 0 || r >= _rows || c < 0 || c >= _cols) return;
        var fg = _reverse ? _bg : _fg;
        var bg = _reverse ? _fg : _bg;
        _cells[r, c] = new TermCell(ch, fg, bg, _bold, _underline);
    }

    private void LineFeed()
    {
        _cursorRow++;
        if (_cursorRow >= _rows) { ScrollUp(1); _cursorRow = _rows - 1; }
    }

    private void ReverseLineFeed()
    {
        _cursorRow--;
        if (_cursorRow < 0) { ScrollDown(1); _cursorRow = 0; }
    }

    private void ScrollUp(int n)
    {
        n = Math.Clamp(n, 0, _rows);
        for (var r = 0; r < _rows - n; r++)
        for (var c = 0; c < _cols; c++)
            _cells[r, c] = _cells[r + n, c];
        for (var r = Math.Max(0, _rows - n); r < _rows; r++)
        for (var c = 0; c < _cols; c++)
            _cells[r, c] = default;
    }

    private void ScrollDown(int n)
    {
        n = Math.Clamp(n, 0, _rows);
        for (var r = _rows - 1; r >= n; r--)
        for (var c = 0; c < _cols; c++)
            _cells[r, c] = _cells[r - n, c];
        for (var r = 0; r < n; r++)
        for (var c = 0; c < _cols; c++)
            _cells[r, c] = default;
    }

    private void EraseDisplay(int mode)
    {
        switch (mode)
        {
            case 0: // cursor to end
                EraseLine(0);
                for (var r = _cursorRow + 1; r < _rows; r++) ClearRow(r);
                break;
            case 1: // start to cursor
                for (var r = 0; r < _cursorRow; r++) ClearRow(r);
                EraseLine(1);
                break;
            case 2: case 3: ClearAllCells(); break;
        }
    }

    private void EraseLine(int mode)
    {
        switch (mode)
        {
            case 0: for (var c = _cursorCol; c < _cols; c++) _cells[_cursorRow, c] = default; break;
            case 1: for (var c = 0; c <= _cursorCol; c++)   _cells[_cursorRow, c] = default; break;
            case 2: ClearRow(_cursorRow); break;
        }
    }

    // CSI <n> X - blank n cells starting at the cursor without moving the cursor.
    // Cursor stays put; cells beyond the line end are not affected.
    private void EraseChars(int n)
    {
        var end = Math.Min(_cursorCol + n, _cols);
        for (var c = _cursorCol; c < end; c++)
            _cells[_cursorRow, c] = default;
    }

    private void ClearRow(int r) { for (var c = 0; c < _cols; c++) _cells[r, c] = default; }

    private void ClearAllCells()
    {
        for (var r = 0; r < _rows; r++)
        for (var c = 0; c < _cols; c++)
            _cells[r, c] = default;
    }

    private void InsertLines(int n)
    {
        for (var r = _rows - 1; r >= _cursorRow + n; r--)
        for (var c = 0; c < _cols; c++)
            _cells[r, c] = _cells[r - n, c];
        for (var r = _cursorRow; r < _cursorRow + n && r < _rows; r++) ClearRow(r);
    }

    private void DeleteLines(int n)
    {
        for (var r = _cursorRow; r < _rows - n; r++)
        for (var c = 0; c < _cols; c++)
            _cells[r, c] = _cells[r + n, c];
        for (var r = Math.Max(0, _rows - n); r < _rows; r++) ClearRow(r);
    }

    private void InsertChars(int n)
    {
        for (var c = _cols - 1; c >= _cursorCol + n; c--)
            _cells[_cursorRow, c] = _cells[_cursorRow, c - n];
        for (var c = _cursorCol; c < _cursorCol + n && c < _cols; c++)
            _cells[_cursorRow, c] = default;
    }

    private void DeleteChars(int n)
    {
        for (var c = _cursorCol; c < _cols - n; c++)
            _cells[_cursorRow, c] = _cells[_cursorRow, c + n];
        for (var c = Math.Max(0, _cols - n); c < _cols; c++)
            _cells[_cursorRow, c] = default;
    }

    // ── Key mapping ───────────────────────────────────────────────────────────
    private static string? KeyToVt(Key key, KeyModifiers mods)
    {
        var ctrl  = mods.HasFlag(KeyModifiers.Control);
        var shift = mods.HasFlag(KeyModifiers.Shift);
        var alt   = mods.HasFlag(KeyModifiers.Alt);
        var mod   = (shift ? 1 : 0) | (alt ? 2 : 0) | (ctrl ? 4 : 0);
        var modSuffix = mod > 0 ? $";{mod + 1}" : "";

        return key switch
        {
            // Arrow keys: plain → CSI A-D; with any modifier → CSI 1;<mod+1> A-D
            Key.Up    => mod > 0 ? $"\x1b[1{modSuffix}A" : "\x1b[A",
            Key.Down  => mod > 0 ? $"\x1b[1{modSuffix}B" : "\x1b[B",
            Key.Right => mod > 0 ? $"\x1b[1{modSuffix}C" : "\x1b[C",
            Key.Left  => mod > 0 ? $"\x1b[1{modSuffix}D" : "\x1b[D",

            // Home / End: unmodified → SS3 form (readline / bash expect \x1bOH/OF);
            //             modified   → CSI 1;<mod+1> H/F
            Key.Home   => mod > 0 ? $"\x1b[1{modSuffix}H" : "\x1bOH",
            Key.End    => mod > 0 ? $"\x1b[1{modSuffix}F" : "\x1bOF",

            Key.Insert   => $"\x1b[2{modSuffix}~",
            Key.Delete   => $"\x1b[3{modSuffix}~",
            Key.PageUp   => $"\x1b[5{modSuffix}~",
            Key.PageDown => $"\x1b[6{modSuffix}~",
            Key.F1  => "\x1bOP",
            Key.F2  => "\x1bOQ",
            Key.F3  => "\x1bOR",
            Key.F4  => "\x1bOS",
            Key.F5  => "\x1b[15~",
            Key.F6  => "\x1b[17~",
            Key.F7  => "\x1b[18~",
            Key.F8  => "\x1b[19~",
            Key.F9  => "\x1b[20~",
            Key.F10 => "\x1b[21~",
            Key.F11 => "\x1b[23~",
            Key.F12 => "\x1b[24~",
            Key.Tab    => shift ? "\x1b[Z" : "\t",
            Key.Return => "\r",
            Key.Escape => "\x1b",
            // Backspace: Ctrl+Backspace → word-erase (^H = 0x08); plain → DEL (0x7f)
            Key.Back => ctrl ? "\x08" : "\x7f",
            _ => null
        };
    }

    private static int? ControlChar(Key key) => key switch
    {
        Key.A => 1,  Key.B => 2,  Key.C => 3,  Key.D => 4,  Key.E => 5,
        Key.F => 6,  Key.G => 7,  Key.H => 8,  Key.I => 9,  Key.J => 10,
        Key.K => 11, Key.L => 12, Key.M => 13, Key.N => 14, Key.O => 15,
        Key.P => 16, Key.Q => 17, Key.R => 18, Key.S => 19, Key.T => 20,
        Key.U => 21, Key.V => 22, Key.W => 23, Key.X => 24, Key.Y => 25,
        Key.Z => 26, _ => null
    };

    // Returns the lowercase ASCII char for readline Alt+key sequences (ESC-prefixed).
    // Only covers the common readline / vim bindings to avoid swallowing AltGr combos
    // that should come through as TextInput (e.g. Alt+E → é on some keyboard layouts).
    private static string? AltKeyChar(Key key) => key switch
    {
        Key.B => "b",   // Alt+B - move word back
        Key.F => "f",   // Alt+F - move word forward
        Key.D => "d",   // Alt+D - delete word forward
        Key.Back => "\x7f", // Alt+Backspace - delete word back (ESC DEL)
        Key.OemPeriod => ".", // Alt+. - insert last argument
        _ => null
    };

    // ── Helpers ───────────────────────────────────────────────────────────────
    private static List<int> ParseNums(string s)
    {
        var result = new List<int>();
        foreach (var part in s.Split(';', StringSplitOptions.RemoveEmptyEntries))
            if (int.TryParse(part, out var n)) result.Add(n);
            else result.Add(0);
        return result;
    }

    private static Color Palette256(int n)
    {
        if (n < 16)  return AnsiPalette[n];
        if (n < 232)
        {
            n -= 16;
            var b = n % 6; var g = (n / 6) % 6; var r = n / 36;
            return Color.FromRgb((byte)(r > 0 ? 55 + r * 40 : 0),
                                  (byte)(g > 0 ? 55 + g * 40 : 0),
                                  (byte)(b > 0 ? 55 + b * 40 : 0));
        }
        var v = (byte)(8 + (n - 232) * 10);
        return Color.FromRgb(v, v, v);
    }
}

// ── Cell struct ───────────────────────────────────────────────────────────────
public readonly record struct TermCell(
    char  Char,
    Color? Fg,
    Color? Bg,
    bool  Bold,
    bool  Underline);

// ── Screen snapshot ───────────────────────────────────────────────────────────
/// <summary>
/// Immutable capture of a <see cref="ConsoleTerminal"/> screen buffer.
/// Stored on <see cref="Kodo.Models.TerminalSession"/> so that switching
/// between sessions restores each session's last-visible output.
/// </summary>
public sealed class TerminalSnapshot(
    TermCell[,] cells,
    int rows, int cols,
    int cursorRow, int cursorCol, bool cursorVisible,
    Color fg, Color bg, bool bold, bool underline, bool reverse,
    ConsoleTerminal.ParseState parseState, string csiParam)
{
    internal TermCell[,] Cells         { get; } = cells;
    public   int         Rows          { get; } = rows;
    public   int         Cols          { get; } = cols;
    internal int         CursorRow     { get; } = cursorRow;
    internal int         CursorCol     { get; } = cursorCol;
    internal bool        CursorVisible { get; } = cursorVisible;
    internal Color       Fg            { get; } = fg;
    internal Color       Bg            { get; } = bg;
    internal bool        Bold          { get; } = bold;
    internal bool        Underline     { get; } = underline;
    internal bool        Reverse       { get; } = reverse;
    internal ConsoleTerminal.ParseState ParseState { get; } = parseState;
    internal string      CsiParam      { get; } = csiParam;
}

// ── ConPTY P/Invoke ───────────────────────────────────────────────────────────
internal static class NativeConPty
{
    [StructLayout(LayoutKind.Sequential)]
    public struct COORD { public short X; public short Y; }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct STARTUPINFOEXW
    {
        public STARTUPINFOW StartupInfo;
        public IntPtr       lpAttributeList;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct STARTUPINFOW
    {
        public int    cb;
        public string? lpReserved, lpDesktop, lpTitle;
        public int    dwX, dwY, dwXSize, dwYSize;
        public int    dwXCountChars, dwYCountChars, dwFillAttribute, dwFlags;
        public short  wShowWindow, cbReserved2;
        public IntPtr lpReserved2, hStdInput, hStdOutput, hStdError;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct PROCESS_INFORMATION
    {
        public IntPtr hProcess, hThread;
        public int    dwProcessId, dwThreadId;
    }

    private const uint EXTENDED_STARTUPINFO_PRESENT = 0x00080000;
    private const uint PROC_THREAD_ATTRIBUTE_PSEUDOCONSOLE = 0x00020016;
    private const int  STARTF_USESTDHANDLES = 0x100;

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool CreatePipe(out IntPtr hRead, out IntPtr hWrite, IntPtr lpAttr, int nSize);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int CreatePseudoConsole(COORD size, IntPtr hInput, IntPtr hOutput, uint flags, out IntPtr hPcon);

    [DllImport("kernel32.dll")]
    public static extern void ClosePseudoConsole(IntPtr hPcon);

    [DllImport("kernel32.dll")]
    public static extern int ResizePseudoConsole(IntPtr hPcon, COORD size);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool CloseHandle(IntPtr h);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern uint WaitForSingleObject(IntPtr hHandle, uint ms);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool InitializeProcThreadAttributeList(IntPtr list, int count, int flags, ref IntPtr size);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool UpdateProcThreadAttribute(IntPtr list, uint flags, IntPtr attr, IntPtr val, IntPtr size, IntPtr prev, IntPtr retSize);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool DeleteProcThreadAttributeList(IntPtr list);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CreateProcessW(
        string? app, StringBuilder cmd,
        IntPtr pAttr, IntPtr tAttr,
        bool inherit, uint flags,
        IntPtr env, string? dir,
        ref STARTUPINFOEXW si, out PROCESS_INFORMATION pi);

    public static void LaunchProcess(StringBuilder cmdLine, string workingDir, IntPtr hPcon,
        out IntPtr hProcess, out IntPtr hThread)
    {
        // Build a PROC_THREAD_ATTRIBUTE_LIST containing the ConPTY handle
        var attrListSize = IntPtr.Zero;
        InitializeProcThreadAttributeList(IntPtr.Zero, 1, 0, ref attrListSize);
        var attrList = Marshal.AllocHGlobal(attrListSize);
        try
        {
            if (!InitializeProcThreadAttributeList(attrList, 1, 0, ref attrListSize))
                throw new InvalidOperationException($"InitializeProcThreadAttributeList: {Marshal.GetLastWin32Error()}");

            var hPconLocal = hPcon;
            var hPconPtr = Marshal.AllocHGlobal(IntPtr.Size);
            Marshal.WriteIntPtr(hPconPtr, hPconLocal);
            try
            {
                if (!UpdateProcThreadAttribute(attrList, 0,
                        (IntPtr)PROC_THREAD_ATTRIBUTE_PSEUDOCONSOLE,
                        hPconLocal, (IntPtr)IntPtr.Size, IntPtr.Zero, IntPtr.Zero))
                    throw new InvalidOperationException($"UpdateProcThreadAttribute: {Marshal.GetLastWin32Error()}");

                var si = new STARTUPINFOEXW
                {
                    StartupInfo = new STARTUPINFOW
                    {
                        cb     = Marshal.SizeOf<STARTUPINFOEXW>(),
                        dwFlags = 0 // ConPTY owns stdio via attribute list; STARTF_USESTDHANDLES with null handles breaks stdin
                    },
                    lpAttributeList = attrList
                };

                if (!CreateProcessW(null, cmdLine, IntPtr.Zero, IntPtr.Zero, false,
                        EXTENDED_STARTUPINFO_PRESENT, IntPtr.Zero, workingDir, ref si, out var pi))
                    throw new InvalidOperationException($"CreateProcessW: {Marshal.GetLastWin32Error()}");

                hProcess = pi.hProcess;
                hThread  = pi.hThread;
            }
            finally { Marshal.FreeHGlobal(hPconPtr); }
        }
        finally
        {
            DeleteProcThreadAttributeList(attrList);
            Marshal.FreeHGlobal(attrList);
        }
    }
}