using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using AvaloniaEdit.CodeCompletion;
using AvaloniaEdit.Document;
using AvaloniaEdit.Editing;

namespace Kodo;

// What an IntelliSenseSuggestion was sourced from - drives sort order and the
// label shown in the completion list.
public enum IntelliSenseKind { Variable, Function, Property, Type, Namespace, Keyword }

// One row in the predictive completion popup. Implements AvaloniaEdit's
// ICompletionData directly so CompletionList can bind to it with no adapter.
public sealed class IntelliSenseSuggestion : ICompletionData
{
    public IntelliSenseKind Kind { get; }
    public string Text { get; }
    public IImage? Image => null;
    // Returning a fully-built control (rather than a plain string) means the completion
    // list renders our own icon-chip + name + kind-label row instead of AvaloniaEdit's
    // default single-line text - this is what gives the popup its VSCode-like look.
    // Built fresh on every access rather than cached: a Control can only have one visual
    // parent at a time, and re-fetching Content is cheap for a ~25 row list.
    public object Content => BuildContentVisual();
    // Deliberately null: AvaloniaEdit shows this in its own separate description
    // popup alongside the completion list. Since the kind label is already drawn
    // in-row (kindBlock, below), a non-null Description here spawned a second
    // overlapping popup that duplicated that same text - the "duplicate window" /
    // "duplicate tooltip" glitch.
    public object? Description => null;
    public double Priority => Kind switch
    {
        IntelliSenseKind.Variable  => 5,
        IntelliSenseKind.Function  => 4,
        IntelliSenseKind.Property  => 3,
        IntelliSenseKind.Type      => 2,
        IntelliSenseKind.Namespace => 1,
        IntelliSenseKind.Keyword   => 0,
        _ => 0,
    };

    public IntelliSenseSuggestion(string text, IntelliSenseKind kind)
    {
        Text = text;
        Kind = kind;
    }

    public void Complete(TextArea textArea, ISegment completionSegment, EventArgs insertionRequestEventArgs) =>
        textArea.Document.Replace(completionSegment, Text);

    private static string KindLabel(IntelliSenseKind kind) => kind switch
    {
        IntelliSenseKind.Variable  => "Variable (this file)",
        IntelliSenseKind.Function  => "Function",
        IntelliSenseKind.Property  => "Property",
        IntelliSenseKind.Type      => "Type",
        IntelliSenseKind.Namespace => "Namespace",
        IntelliSenseKind.Keyword   => "Keyword",
        _ => string.Empty,
    };

    // One-letter icon glyph + accent color per kind, shown as a small rounded chip -
    // a lightweight stand-in for VS Code's Codicon glyphs since no icon font is bundled.
    private static (string Glyph, string Color) GlyphAndColorFor(IntelliSenseKind kind) => kind switch
    {
        IntelliSenseKind.Variable  => ("V", "#3B82F6"),
        IntelliSenseKind.Function  => ("F", "#A855F7"),
        IntelliSenseKind.Property  => ("P", "#14B8A6"),
        IntelliSenseKind.Type      => ("T", "#F97316"),
        IntelliSenseKind.Namespace => ("N", "#10B981"),
        IntelliSenseKind.Keyword   => ("K", "#6366F1"),
        _ => ("•", "#6B7280"),
    };

    // Builds one completion-list row: [icon chip] [name] [kind label, right-aligned] -
    // modeled on VS Code's suggestion widget.
    private Control BuildContentVisual()
    {
        var (glyph, colorHex) = GlyphAndColorFor(Kind);

        var iconChip = new Border
        {
            Width = 20,
            Height = 20,
            CornerRadius = new CornerRadius(4),
            Background = new SolidColorBrush(Color.Parse(colorHex)),
            VerticalAlignment = VerticalAlignment.Center,
            Child = new TextBlock
            {
                Text = glyph,
                FontSize = 11,
                FontWeight = FontWeight.Bold,
                Foreground = Brushes.White,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
            },
        };

        var nameBlock = new TextBlock
        {
            Text = Text,
            Foreground = Brushes.WhiteSmoke,
            FontFamily = new FontFamily("Cascadia Code,Consolas,Menlo,Monospace"),
            FontSize = 13,
            Margin = new Thickness(8, 0, 16, 0),
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            TextWrapping = TextWrapping.NoWrap,
        };

        var kindBlock = new TextBlock
        {
            Text = KindLabel(Kind),
            Foreground = new SolidColorBrush(Color.Parse("#8A8A8A")),
            FontSize = 11,
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Right,
            TextTrimming = TextTrimming.CharacterEllipsis,
            TextWrapping = TextWrapping.NoWrap,
            // Sized to fit the longest label ("Variable (this file)") in full at
            // this font/size - previously capped at 116, which was too narrow and
            // ellipsis-truncated it to "Variable (this f...".
            MaxWidth = 160,
        };

        var row = new Grid
        {
            VerticalAlignment = VerticalAlignment.Center,
            // Without this, text that's still wider than its column (e.g. before
            // trimming kicks in during layout) paints past the column boundary
            // instead of being cut off, which is what produced the overlapping/
            // "ghosted double popup" look.
            ClipToBounds = true,
        };
        row.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));
        row.ColumnDefinitions.Add(new ColumnDefinition(new GridLength(1, GridUnitType.Star)) { MinWidth = 40 });
        row.ColumnDefinitions.Add(new ColumnDefinition(GridLength.Auto));

        Grid.SetColumn(iconChip, 0);
        Grid.SetColumn(nameBlock, 1);
        Grid.SetColumn(kindBlock, 2);
        row.Children.Add(iconChip);
        row.Children.Add(nameBlock);
        row.Children.Add(kindBlock);

        return row;
    }
}

/// <summary>
/// Predictive IntelliSense engine. Pulls language candidates (keywords, types, functions,
/// properties, namespaces) from the active .kox language profile via <see cref="LoadedExtension"/>,
/// and separately tracks variables the user has declared in each open file so they show up as
/// completions even though no extension knows about them.
/// </summary>
public sealed class IntelliSenseEngine
{
    // Per-file variable cache. Keyed by whatever the caller uses to identify a document
    // (file path for saved files, a synthetic key for untitled tabs) so variables survive
    // tab switches but don't leak between unrelated files.
    private readonly Dictionary<string, HashSet<string>> _variablesByFile = new(StringComparer.OrdinalIgnoreCase);

    // ── Variable-declaration heuristics ───────────────────────────────────────
    // Deliberately loose, multi-language regexes rather than a real parser - good enough to
    // surface "things the user just named" across the languages Kodo ships profiles for.

    private static readonly Regex TypedOrKeywordDeclaration = new(
        @"\b(?:var|let|val|const|int|long|short|byte|sbyte|uint|ulong|ushort|float|double|decimal|bool|char|string|object|dynamic|auto)\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:[:,][^=]*)?=(?!=)",
        RegexOptions.Compiled);

    // Rust/Kotlin/Swift style, including the optional `mut` modifier that TypedOrKeywordDeclaration
    // above would otherwise mistake for the variable name (e.g. "let mut x = 5").
    private static readonly Regex ModifiedBinding = new(
        @"\b(?:let|var|val)\s+(?:mut\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*(?::[^=]*)?=(?!=)",
        RegexOptions.Compiled);

    // Bare "name = value" (Python, shell, plain assignment) - anchored to line start so it
    // doesn't fire mid-expression, and the negative lookahead keeps ==, <=, >=, != out.
    private static readonly Regex BareAssignment = new(
        @"^\s*([A-Za-z_][A-Za-z0-9_]*)\s*(?::\s*[A-Za-z_][A-Za-z0-9_.<>\[\],\s]*)?=(?!=)",
        RegexOptions.Compiled);

    // Loop/catch/using binders: foreach (var x in ...), for (int i = ...), catch (Exception ex), using (var f = ...)
    private static readonly Regex LoopOrHandlerBinding = new(
        @"\b(?:foreach|for|catch|using)\s*\(\s*(?:var|[A-Za-z_][A-Za-z0-9_<>\[\],.\s]*)\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:in\b|=|\))",
        RegexOptions.Compiled);

    private static readonly Regex[] DeclarationPatterns =
    {
        TypedOrKeywordDeclaration,
        ModifiedBinding,
        BareAssignment,
        LoopOrHandlerBinding,
    };

    // Words a declaration regex could capture as a "name" but that are really language
    // keywords (control flow, modifiers, the `mut` binding annotation, etc.).
    private static readonly HashSet<string> ReservedWords = new(StringComparer.Ordinal)
    {
        "if", "else", "for", "foreach", "while", "switch", "return", "try", "catch", "finally",
        "new", "class", "struct", "interface", "enum", "namespace", "using", "public", "private",
        "protected", "internal", "static", "readonly", "async", "await", "void", "null", "true",
        "false", "this", "base", "import", "from", "def", "function", "fn", "match", "case", "in",
        "is", "as", "out", "ref", "params", "yield", "break", "continue", "do", "default", "throw",
        "mut", "let", "var", "val", "const",
    };

    /// <summary>
    /// Scans a single line of source and returns the names of any variables it initializes.
    /// This is the one place "what counts as a variable declaration" is decided, so every
    /// caller (live scanning, tests, future callers) sees the same answer.
    /// </summary>
    public static IEnumerable<string> IdentifyVariableInitializations(string lineText)
    {
        if (string.IsNullOrWhiteSpace(lineText))
            return [];

        var trimmed = lineText.TrimStart();
        if (trimmed.StartsWith("//") || trimmed.StartsWith('#') ||
            trimmed.StartsWith('*') || trimmed.StartsWith("'''") || trimmed.StartsWith("\"\"\""))
            return [];

        List<string>? found = null;
        foreach (var pattern in DeclarationPatterns)
        {
            foreach (Match match in pattern.Matches(lineText))
            {
                var name = match.Groups[1].Value;
                if (string.IsNullOrEmpty(name) || ReservedWords.Contains(name))
                    continue;
                (found ??= []).Add(name);
            }
        }

        return found is null ? [] : found.Distinct(StringComparer.Ordinal);
    }

    /// <summary>
    /// Rebuilds the variable cache for one file from its current text. Replaces (rather than
    /// merges into) the previous set so a deleted declaration stops being suggested.
    /// </summary>
    public void ScanDocument(string fileKey, string documentText)
    {
        if (string.IsNullOrEmpty(fileKey))
            return;

        var variables = new HashSet<string>(StringComparer.Ordinal);
        if (!string.IsNullOrEmpty(documentText))
        {
            foreach (var line in documentText.Split('\n'))
            {
                foreach (var name in IdentifyVariableInitializations(line))
                    variables.Add(name);
            }
        }

        _variablesByFile[fileKey] = variables;
    }

    /// <summary>Drops the cached variables for a file, e.g. when its tab is closed.</summary>
    public void ForgetFile(string fileKey) => _variablesByFile.Remove(fileKey);

    public IReadOnlyCollection<string> GetVariables(string fileKey) =>
        _variablesByFile.TryGetValue(fileKey, out var vars) ? vars : Array.Empty<string>();

    public static bool IsWordChar(char c) => char.IsLetterOrDigit(c) || c == '_';

    /// <summary>Walks backward from caretOffset to find the start of the identifier being typed.</summary>
    public static int FindWordStart(string documentText, int caretOffset)
    {
        var start = caretOffset;
        while (start > 0 && IsWordChar(documentText[start - 1]))
            start--;
        return start;
    }

    /// <summary>
    /// Builds the ranked, prefix-filtered candidate list for the completion popup: this file's
    /// own variables first, then whatever the active .kox language profile contributes.
    /// </summary>
    public List<IntelliSenseSuggestion> GetSuggestions(
        string prefix,
        string fileKey,
        LoadedExtension? languageExtension,
        int maxResults = 25)
    {
        var results = new List<IntelliSenseSuggestion>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        void AddCandidates(IEnumerable<string> names, IntelliSenseKind kind)
        {
            foreach (var name in names)
            {
                if (string.IsNullOrEmpty(name)) continue;
                if (!string.IsNullOrEmpty(prefix) &&
                    !name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (!seen.Add(name)) continue;
                results.Add(new IntelliSenseSuggestion(name, kind));
            }
        }

        AddCandidates(GetVariables(fileKey), IntelliSenseKind.Variable);

        if (languageExtension is not null)
        {
            AddCandidates(languageExtension.Functions, IntelliSenseKind.Function);
            AddCandidates(languageExtension.Properties, IntelliSenseKind.Property);
            AddCandidates(languageExtension.Types, IntelliSenseKind.Type);
            AddCandidates(languageExtension.Namespaces, IntelliSenseKind.Namespace);
            AddCandidates(languageExtension.Keywords, IntelliSenseKind.Keyword);
        }

        return results
            .OrderByDescending(s => s.Priority)
            .ThenBy(s => s.Text, StringComparer.OrdinalIgnoreCase)
            .Take(maxResults)
            .ToList();
    }
}