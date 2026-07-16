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

// Source of a CodePredictSuggestion - drives sort order and label.
public enum CodePredictKind { Variable, Function, Property, Type, Namespace, Keyword }

// One row in the predictive completion popup.
public sealed class CodePredictSuggestion : ICompletionData
{
    // Set by MainWindow before building suggestions, from its live theme brushes.
    // Keeps the popup's text readable across light/dark theme and custom accents.
    public static IBrush PanelForeground { get; set; } = Brushes.WhiteSmoke;
    public static IBrush MutedForeground { get; set; } = new SolidColorBrush(Color.Parse("#8A8A8A"));

    public CodePredictKind Kind { get; }
    public string Text { get; }
    public IImage? Image => null;
    // Built lazily, once, on first access - avoids rebuilding the icon+name+kind
    // row's control tree on every measure/arrange/render pass.
    private Control? _content;
    public object Content => _content ??= BuildContentVisual();
    // Null: AvaloniaEdit shows this in its own popup, which duplicated the kind
    // label already drawn in-row (kindBlock below).
    public object? Description => null;
    public double Priority => Kind switch
    {
        CodePredictKind.Variable  => 5,
        CodePredictKind.Function  => 4,
        CodePredictKind.Property  => 3,
        CodePredictKind.Type      => 2,
        CodePredictKind.Namespace => 1,
        CodePredictKind.Keyword   => 0,
        _ => 0,
    };

    public CodePredictSuggestion(string text, CodePredictKind kind)
    {
        Text = text;
        Kind = kind;
    }

    public void Complete(TextArea textArea, ISegment completionSegment, EventArgs insertionRequestEventArgs) =>
        textArea.Document.Replace(completionSegment, Text);

    private static string KindLabel(CodePredictKind kind) => kind switch
    {
        CodePredictKind.Variable  => "Variable (this file)",
        CodePredictKind.Function  => "Function",
        CodePredictKind.Property  => "Property",
        CodePredictKind.Type      => "Type",
        CodePredictKind.Namespace => "Namespace",
        CodePredictKind.Keyword   => "Keyword",
        _ => string.Empty,
    };

    // Icon glyph + accent color per kind - stand-in for VS Code's Codicons.
    private static (string Glyph, string Color) GlyphAndColorFor(CodePredictKind kind) => kind switch
    {
        CodePredictKind.Variable  => ("V", "#3B82F6"),
        CodePredictKind.Function  => ("F", "#A855F7"),
        CodePredictKind.Property  => ("P", "#14B8A6"),
        CodePredictKind.Type      => ("T", "#F97316"),
        CodePredictKind.Namespace => ("N", "#10B981"),
        CodePredictKind.Keyword   => ("K", "#6366F1"),
        _ => ("•", "#6B7280"),
    };

    // Glyph background brushes, built once and reused across every row/rebuild.
    private static readonly Dictionary<CodePredictKind, IBrush> GlyphBrushes =
        Enum.GetValues<CodePredictKind>().ToDictionary(
            k => k,
            k => (IBrush)new SolidColorBrush(Color.Parse(GlyphAndColorFor(k).Color)));

    private static readonly FontFamily MonoFontFamily = new("Cascadia Code,Consolas,Menlo,Monospace");

    // [icon chip] [name] [kind label, right-aligned] - VS Code-style row, themed
    // to Kodo's live panel/text colors.
    private Control BuildContentVisual()
    {
        var (glyph, _) = GlyphAndColorFor(Kind);

        var iconChip = new Border
        {
            Width = 20,
            Height = 20,
            CornerRadius = new CornerRadius(6),
            Background = GlyphBrushes[Kind],
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
            Foreground = PanelForeground,
            FontFamily = MonoFontFamily,
            FontSize = 13,
            Margin = new Thickness(8, 0, 16, 0),
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            TextWrapping = TextWrapping.NoWrap,
        };

        var kindBlock = new TextBlock
        {
            Text = KindLabel(Kind),
            Foreground = MutedForeground,
            FontSize = 11,
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Right,
            TextTrimming = TextTrimming.CharacterEllipsis,
            TextWrapping = TextWrapping.NoWrap,
            MaxWidth = 160,
        };

        var row = new Grid
        {
            VerticalAlignment = VerticalAlignment.Center,
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

/// <summary>Predictive CodePredict engine: language candidates from the active .kox profile, plus per-file declared variables.</summary>
public sealed class CodePredictEngine
{
    private readonly Dictionary<string, HashSet<string>> _variablesByFile = new(StringComparer.OrdinalIgnoreCase);

    private const string NotCompoundOrArrow = @"(?<![=!<>+\-*/%&|^~])=(?![=>])";

    private static readonly Regex TypedOrKeywordDeclaration = new(
        @"\b(?:var|let|val|const|int|long|short|byte|sbyte|uint|ulong|ushort|float|double|decimal|bool|char|string|object|dynamic|auto)\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:[:,][^=]*)?" + NotCompoundOrArrow,
        RegexOptions.Compiled);

    private static readonly Regex ModifiedBinding = new(
        @"\b(?:let|var|val)\s+(?:mut\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*(?::[^=]*)?" + NotCompoundOrArrow,
        RegexOptions.Compiled);

    private static readonly Regex BareAssignment = new(
        @"^\s*([A-Za-z_][A-Za-z0-9_]*)\s*(?::\s*[A-Za-z_][A-Za-z0-9_.<>\[\],\s]*)?" + NotCompoundOrArrow,
        RegexOptions.Compiled);

    private static readonly Regex LoopOrHandlerBinding = new(
        @"\b(?:foreach|for|catch|using)\s*\(\s*(?:var|[A-Za-z_][A-Za-z0-9_<>,.\[\]\s]*)\s+([A-Za-z_][A-Za-z0-9_]*)\s*(?:in\b|=|\))",
        RegexOptions.Compiled);

    private static readonly Regex[] DeclarationPatterns =
    {
        TypedOrKeywordDeclaration,
        ModifiedBinding,
        BareAssignment,
        LoopOrHandlerBinding,
    };

    private static readonly HashSet<string> ReservedWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "if", "elif", "elseif", "else", "for", "foreach", "while", "switch", "match", "case",
        "return", "try", "catch", "finally", "throw", "raise", "new", "class", "struct",
        "interface", "trait", "enum", "namespace", "module", "package", "using", "import",
        "from", "require", "include", "export", "public", "private", "protected", "internal",
        "static", "readonly", "sealed", "abstract", "virtual", "override", "async", "await",
        "void", "null", "nil", "none", "undefined", "true", "false", "this", "self", "base",
        "super", "def", "function", "fn", "lambda", "in", "is", "as", "out", "ref", "params",
        "yield", "break", "continue", "do", "default", "goto", "mut", "let", "var", "val",
        "const", "end", "then", "done", "fi", "esac", "not", "and", "or",
        "echo", "print", "println", "printf", "console", "exit", "call", "cls", "pause", "rem",
        "cd", "dir", "mkdir", "rmdir", "del", "copy", "move", "ren", "shift", "setlocal",
        "endlocal", "errorlevel", "start", "taskkill", "set", "unset", "source",
        "sudo", "chmod", "chown", "curl", "wget", "ssh", "scp", "tar", "zip", "unzip",
        "git", "npm", "npx", "yarn", "pnpm", "pip", "python", "python3", "node", "dotnet",
        "cargo", "go", "make", "cmake", "gradle", "mvn", "docker", "kubectl", "brew", "apt",
        "build", "publish", "install", "run", "test", "clean", "deploy", "restore", "release",
        "debug",
    };

    private static readonly Regex TrailingLineComment = new(
        @"(?:(?<=\s)|^)(?://|\#)(?!\S).*$|(?:(?<=\s)|^)(?://|\#)\s.*$",
        RegexOptions.Compiled);

    private static string MaskStringLiterals(string lineText)
    {
        if (lineText.IndexOfAny(['"', '\'']) < 0)
            return lineText;

        var chars = lineText.ToCharArray();
        char? quote = null;
        for (var i = 0; i < chars.Length; i++)
        {
            var c = chars[i];
            if (quote is null)
            {
                if (c is '"' or '\'')
                    quote = c;
                continue;
            }

            if (c == '\\' && i + 1 < chars.Length)
            {
                chars[i] = ' ';
                chars[i + 1] = ' ';
                i++;
                continue;
            }

            if (c == quote)
            {
                quote = null;
                continue;
            }

            chars[i] = ' ';
        }

        return new string(chars);
    }

    public static IEnumerable<string> IdentifyVariableInitializations(string lineText)
    {
        if (string.IsNullOrWhiteSpace(lineText))
            return [];

        var trimmed = lineText.TrimStart();
        if (trimmed.StartsWith("//") || trimmed.StartsWith('#') ||
            trimmed.StartsWith('*') || trimmed.StartsWith("'''") || trimmed.StartsWith("\"\"\"") ||
            trimmed.StartsWith("--") || trimmed.StartsWith(';') ||
            trimmed.StartsWith("REM ", StringComparison.OrdinalIgnoreCase) ||
            trimmed.Equals("REM", StringComparison.OrdinalIgnoreCase))
            return [];

        var scanText = TrailingLineComment.Replace(MaskStringLiterals(lineText), string.Empty);
        if (string.IsNullOrWhiteSpace(scanText))
            return [];

        List<string>? found = null;
        foreach (var pattern in DeclarationPatterns)
        {
            foreach (Match match in pattern.Matches(scanText))
            {
                var name = match.Groups[1].Value;
                if (string.IsNullOrEmpty(name) || ReservedWords.Contains(name))
                    continue;
                (found ??= []).Add(name);
            }
        }

        return found is null ? [] : found.Distinct(StringComparer.Ordinal);
    }

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

    public void ForgetFile(string fileKey) => _variablesByFile.Remove(fileKey);

    public IReadOnlyCollection<string> GetVariables(string fileKey) =>
        _variablesByFile.TryGetValue(fileKey, out var vars) ? vars : Array.Empty<string>();

    public static bool IsWordChar(char c) => char.IsLetterOrDigit(c) || c == '_';

    public static int FindWordStart(string documentText, int caretOffset)
    {
        var start = caretOffset;
        while (start > 0 && IsWordChar(documentText[start - 1]))
            start--;
        return start;
    }

    private static IReadOnlyCollection<string> GetEnclosingCallNames(string documentText, int caretOffset)
    {
        if (string.IsNullOrEmpty(documentText) || caretOffset <= 0)
            return Array.Empty<string>();

        var callStack = new Stack<string?>();
        var token = new List<char>();
        var inString = false;
        var stringDelimiter = '\0';
        var escaped = false;
        var inLineComment = false;

        static string? FlushToken(List<char> chars)
        {
            if (chars.Count == 0)
                return null;
            var value = new string(chars.ToArray());
            chars.Clear();
            return value;
        }

        for (var i = 0; i < caretOffset; i++)
        {
            var c = documentText[i];

            if (inLineComment)
            {
                if (c == '\n')
                    inLineComment = false;
                continue;
            }

            if (inString)
            {
                if (escaped)
                {
                    escaped = false;
                    continue;
                }

                if (c == '\\')
                {
                    escaped = true;
                    continue;
                }

                if (c == stringDelimiter)
                {
                    inString = false;
                    stringDelimiter = '\0';
                }

                continue;
            }

            if (c == '/' && i + 1 < caretOffset && documentText[i + 1] == '/')
            {
                inLineComment = true;
                token.Clear();
                i++;
                continue;
            }

            if (c is '"' or '\'')
            {
                inString = true;
                stringDelimiter = c;
                token.Clear();
                continue;
            }

            if (IsWordChar(c))
            {
                token.Add(c);
                continue;
            }

            if (c == '(')
            {
                callStack.Push(FlushToken(token));
                continue;
            }

            token.Clear();

            if (c == ')')
            {
                if (callStack.Count > 0)
                    callStack.Pop();
            }
        }

        return callStack.Where(name => !string.IsNullOrWhiteSpace(name)).Select(name => name!).ToArray();
    }

    public List<CodePredictSuggestion> GetSuggestions(
        string prefix,
        string fileKey,
        LoadedExtension? languageExtension,
        string documentText,
        int caretOffset,
        int maxResults = 25)
    {
        var results = new List<CodePredictSuggestion>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var enclosingCalls = GetEnclosingCallNames(documentText, caretOffset);
        var blacklist = languageExtension?.Blacklist is { Length: > 0 }
            ? new HashSet<string>(languageExtension.Blacklist, StringComparer.OrdinalIgnoreCase)
            : null;

        void AddCandidates(IEnumerable<string> names, CodePredictKind kind)
        {
            foreach (var name in names)
            {
                if (string.IsNullOrEmpty(name)) continue;
                if (!string.IsNullOrEmpty(prefix) &&
                    !name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (blacklist is not null && enclosingCalls.Any(call => blacklist.Contains(call) && string.Equals(call, name, StringComparison.OrdinalIgnoreCase)))
                    continue;
                if (!seen.Add(name)) continue;
                results.Add(new CodePredictSuggestion(name, kind));
            }
        }

        AddCandidates(GetVariables(fileKey), CodePredictKind.Variable);

        if (languageExtension is not null)
        {
            AddCandidates(languageExtension.Functions, CodePredictKind.Function);
            AddCandidates(languageExtension.Properties, CodePredictKind.Property);
            AddCandidates(languageExtension.Types, CodePredictKind.Type);
            AddCandidates(languageExtension.Namespaces, CodePredictKind.Namespace);
            AddCandidates(languageExtension.Keywords, CodePredictKind.Keyword);
        }

        return results
            .OrderByDescending(s => s.Priority)
            .ThenBy(s => s.Text, StringComparer.OrdinalIgnoreCase)
            .Take(maxResults)
            .ToList();
    }
}


