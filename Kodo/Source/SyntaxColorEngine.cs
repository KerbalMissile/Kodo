using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Avalonia.Media;

namespace Kodo;

public sealed class CompiledSyntaxProfile
{
    private const string MarkupTextToken = "markupText";
    private const string MarkupTextFallback = "#F4F4F4";
    private const string VariableIdentifierBodyPattern = "[\\p{L}_][\\p{L}\\p{Nd}_]*";
    private const string CommonStringPrefixPattern =
        "(?i)(?<![\\p{L}\\p{Nd}_])(?:fr|rf|br|rb|ur|ru|cr|rc|f|r|u|b|c)(?=(?:\\\"\\\"\\\"|'''|\\\"|'|#+\\\"))";

    public LoadedExtension Extension { get; }
    public IReadOnlyList<CompiledSyntaxRule> TokenRules { get; }
    public IReadOnlyList<Regex> StringRegexes { get; }
    public Regex? SingleLineCommentRegex { get; }

    private CompiledSyntaxProfile(
        LoadedExtension extension,
        IReadOnlyList<CompiledSyntaxRule> tokenRules,
        IReadOnlyList<Regex> stringRegexes,
        Regex? singleLineCommentRegex)
    {
        Extension = extension;
        TokenRules = tokenRules;
        StringRegexes = stringRegexes;
        SingleLineCommentRegex = singleLineCommentRegex;
    }

    public static CompiledSyntaxProfile Create(LoadedExtension extension)
    {
        var rules = new List<CompiledSyntaxRule>();
        var traits = SyntaxLanguageTraits.From(extension);

        if (extension.Keywords.Length > 0)
            rules.Add(CreateTokenRule(extension.Keywords, traits, SyntaxTokenKind.Keyword, "keyword", "#569CD6"));
        if (extension.Types.Length > 0)
            rules.Add(CreateTokenRule(extension.Types, traits, SyntaxTokenKind.Type, "type", "#4EC9B0"));
        if (extension.Namespaces.Length > 0)
            rules.Add(CreateTokenRule(extension.Namespaces, traits, SyntaxTokenKind.Namespace, "namespace", "#4FC1FF"));
        if (extension.Properties.Length > 0)
            rules.Add(CreateTokenRule(extension.Properties, traits, SyntaxTokenKind.Property, "property", "#9CDCFE"));
        if (extension.Functions.Length > 0)
            rules.Add(CreateTokenRule(extension.Functions, traits, SyntaxTokenKind.Function, "function", "#DCDCAA"));

        if (extension.StringDelimiters.Contains("\"") ||
            extension.StringDelimiters.Contains("'") ||
            extension.MultiLineStringDelimiters.Contains("\"\"\"") ||
            extension.MultiLineStringDelimiters.Contains("'''"))
        {
            rules.Add(new(new Regex(CommonStringPrefixPattern, RegexOptions.Compiled), "string", "#CE9178"));
        }

        if (traits.IsMarkupLike)
            rules.Add(new(new Regex(@"(?<=>)\s*v?\d[\d._-]*[A-Za-z0-9]*\s*(?=<)", RegexOptions.Compiled), "number", "#B5CEA8"));

        rules.Add(new(new Regex(@"(?<![\p{L}\p{Nd}_])(?:0[xX][0-9A-Fa-f]+|0[bB][01]+|0[oO][0-7]+|\d+(?:\.\d+)?(?:[eE][+\-]?\d+)?)(?![\p{L}\p{Nd}_])", RegexOptions.Compiled), "number", "#B5CEA8"));

        if (traits.IsCssLike)
        {
            rules.Add(new(new Regex(@"(?<![\p{L}\p{Nd}_-])#[0-9A-Fa-f]{3}(?:[0-9A-Fa-f]{1}|[0-9A-Fa-f]{3}|[0-9A-Fa-f]{5})?(?![\p{L}\p{Nd}_-])", RegexOptions.Compiled), "number", "#B5CEA8"));
            rules.Add(new(new Regex(@"(?<![\p{L}\p{Nd}_-])(?:\d+(?:\.\d+)?|\.\d+)(?:px|r?em|ex|ch|lh|rlh|vw|vh|vmin|vmax|vb|vi|svw|svh|lvw|lvh|dvw|dvh|cqw|cqh|cqi|cqb|cqmin|cqmax|cm|mm|Q|in|pc|pt|deg|grad|rad|turn|s|ms|Hz|kHz|dpi|dpcm|dppx|fr|%)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase), "number", "#B5CEA8"));
            rules.Add(new(new Regex(@"(?<![\p{L}\p{Nd}_-])--[\p{L}_-][\p{L}\p{Nd}_-]*(?![\p{L}\p{Nd}_-])", RegexOptions.Compiled), "property", "#9CDCFE"));
            rules.Add(new(new Regex(@"(?<=:):?[\p{L}_-][\p{L}\p{Nd}_-]*(?![\p{L}\p{Nd}_-])", RegexOptions.Compiled), "keyword", "#569CD6"));
            rules.Add(new(new Regex(@"(?<![\p{L}\p{Nd}_-])@[\p{L}_-][\p{L}\p{Nd}_-]*(?![\p{L}\p{Nd}_-])", RegexOptions.Compiled), "preprocessor", "#C586C0"));
        }

        if (!traits.IsMarkupLike)
        {
            rules.Add(new(new Regex(@"(?<=\.|->|::)[\p{L}_][\p{L}\p{Nd}_:-]*", RegexOptions.Compiled), "property", "#9CDCFE"));
            rules.Add(new(new Regex(@"(?<![\p{L}\p{Nd}_])[\p{L}_][\p{L}\p{Nd}_]*(?=\.)", RegexOptions.Compiled), "namespace", "#4FC1FF"));
        }

        if (traits.IsMarkupLike)
        {
            rules.Add(new(new Regex(@"</?|/?>|=", RegexOptions.Compiled), MarkupTextToken, MarkupTextFallback));
        }
        else
        {
            rules.Add(new(new Regex(
                traits.IsCssLike
                    ? @"(?<![\p{L}\p{Nd}_-])@[\p{L}_-][\p{L}\p{Nd}_-]*(?![\p{L}\p{Nd}_-])"
                    : @"(?<![\p{L}\p{Nd}_])[@#][\p{L}_][\p{L}\p{Nd}_-]*",
                RegexOptions.Compiled), "preprocessor", "#C586C0"));
            rules.Add(new(new Regex(@"(?<=\[)[\p{L}_][\p{L}\p{Nd}_:.]*(?=[,\]\(])|(?<=<)[\p{L}_][\p{L}\p{Nd}_:-]*(?=[^>]*>)", RegexOptions.Compiled), "attribute", "#C586C0"));
            rules.Add(new(new Regex(@"(?<![\p{L}\p{Nd}_])[\p{L}_][\p{L}\p{Nd}_]*(?=\s*\()", RegexOptions.Compiled), "function", "#DCDCAA"));
            rules.Add(new(new Regex(@"=>|->|::|\+\+|--|\+=|-=|\*=|/=|%=|&&|\|\||<<|>>|<=|>=|==|!=|=|\+|-|\*|/|%|!|\?|:|<|>|&|\||\^|~", RegexOptions.Compiled), "operator", "#D4D4D4"));
            rules.Add(new(new Regex(@"[{}\[\]();,.]", RegexOptions.Compiled), "punctuation", "#D4D4D4"));
            rules.Add(new(new Regex(@"(?<=\b(?:using|import|include|require|use|from)\b\s+(?:[\p{L}_][\p{L}\p{Nd}_./\\]*\s*[./\\]\s*)?)[\p{L}_][\p{L}\p{Nd}_]*(?=\s*(?:;|$))", RegexOptions.Compiled), "namespace", "#4FC1FF"));
            rules.Add(new(BuildVariableRegex(extension.Keywords.Concat(extension.Types).Concat(extension.Functions).Concat(extension.Properties).Concat(extension.Namespaces)), "variable", "#A0DBFD"));
        }

        if (traits.IsCssLike)
            rules.Add(new(new Regex(@"(?<![\p{L}\p{Nd}_-])#[0-9A-Fa-f]{3}(?:[0-9A-Fa-f]{1}|[0-9A-Fa-f]{3}|[0-9A-Fa-f]{5})?(?![\p{L}\p{Nd}_-])", RegexOptions.Compiled), "number", "#B5CEA8"));

        if (extension.DisableSingleQuoteStrings)
            rules.Add(new(new Regex(@"'(?:\\(?:u[0-9A-Fa-f]{4}|U[0-9A-Fa-f]{8}|x[0-9A-Fa-f]{1,4}|[0-7]{1,3}|[abfnrtv\\""'0])|[^\\'])'", RegexOptions.Compiled), "charLiteral", "#CE9178"));

        var stringRegexes = new List<Regex>();
        foreach (var delimiter in extension.MultiLineStringDelimiters.Where(d => !string.IsNullOrWhiteSpace(d)).Distinct())
        {
            var escaped = Regex.Escape(delimiter);
            stringRegexes.Add(new Regex($@"{escaped}.*?{escaped}", RegexOptions.Compiled));
        }
        foreach (var delimiter in extension.StringDelimiters.Where(d => !string.IsNullOrWhiteSpace(d)).Distinct())
        {
            if (delimiter == "'" && extension.DisableSingleQuoteStrings)
                continue;
            stringRegexes.Add(BuildSingleLineStringRegex(delimiter));
        }

        var singleLineCommentRegex = string.IsNullOrWhiteSpace(extension.CommentLine)
            ? null
            : new Regex(Regex.Escape(extension.CommentLine) + @".*$", RegexOptions.Compiled);

        return new CompiledSyntaxProfile(extension, rules, stringRegexes, singleLineCommentRegex);
    }

    private static CompiledSyntaxRule CreateTokenRule(IEnumerable<string> tokens, SyntaxLanguageTraits traits, SyntaxTokenKind kind, string tokenName, string fallback) =>
        new(traits.IsMarkupLike ? BuildMarkupTokenRegex(tokens, kind) : BuildTokenRegex(tokens, traits), tokenName, fallback);

    private static Regex BuildTokenRegex(IEnumerable<string> tokens, SyntaxLanguageTraits traits)
    {
        var boundary = traits.IsCssLike ? @"[\p{L}\p{Nd}_-]" : @"[\p{L}\p{Nd}_]";
        return new Regex(@"(?<!" + boundary + @")(" + BuildTokenAlternation(tokens) + @")(?!" + boundary + @")", RegexOptions.Compiled);
    }

    private static Regex BuildMarkupTokenRegex(IEnumerable<string> tokens, SyntaxTokenKind kind)
    {
        var alternation = BuildTokenAlternation(tokens);
        var pattern = kind switch
        {
            SyntaxTokenKind.Keyword => @"(?<=</?|<!)(" + alternation + @")(?![\p{L}\p{Nd}_:-])",
            SyntaxTokenKind.Property => @"(?<=\s)(" + alternation + @")(?=\s*=)",
            SyntaxTokenKind.Namespace => @"(?<=\s)(" + alternation + @")(?=:[\p{L}_-][\p{L}\p{Nd}_-]*\s*=)",
            _ => @"(?!)"
        };

        return new Regex(pattern, RegexOptions.Compiled);
    }

    private static string BuildTokenAlternation(IEnumerable<string> tokens) =>
        string.Join("|", tokens
            .Where(token => !string.IsNullOrWhiteSpace(token))
            .Distinct()
            .OrderByDescending(token => token.Length)
            .Select(Regex.Escape));

    private static Regex BuildVariableRegex(IEnumerable<string> reservedTokens)
    {
        var reserved = reservedTokens
            .Where(token => !string.IsNullOrWhiteSpace(token))
            .Distinct(StringComparer.Ordinal)
            .OrderByDescending(token => token.Length)
            .Select(Regex.Escape)
            .ToArray();
        var reservedPrefix = reserved.Length > 0
            ? $"(?!{string.Join("|", reserved.Select(r => r + "(?![\\p{L}\\p{Nd}_])"))})"
            : string.Empty;
        return new Regex($"(?<![.\\p{{L}}\\p{{Nd}}_]){reservedPrefix}{VariableIdentifierBodyPattern}(?!\\s*[\\.(\"'`]|[\\p{{L}}\\p{{Nd}}_])", RegexOptions.Compiled);
    }

    private static Regex BuildSingleLineStringRegex(string delimiter)
    {
        var escaped = Regex.Escape(delimiter);
        return delimiter switch
        {
            "\"" => new Regex("\"(?:\\\\.|[^\"\\\\])*\"", RegexOptions.Compiled),
            "'" => new Regex("'(?:\\\\.|[^'\\\\])*'", RegexOptions.Compiled),
            "`" => new Regex("`(?:\\\\.|[^`\\\\])*`", RegexOptions.Compiled),
            _ => new Regex($@"{escaped}.*?{escaped}", RegexOptions.Compiled)
        };
    }

    private readonly record struct SyntaxLanguageTraits(bool IsCssLike, bool IsMarkupLike)
    {
        public static SyntaxLanguageTraits From(LoadedExtension extension) =>
            new(
                extension.Extensions.Any(ext => string.Equals(ext, ".css", StringComparison.OrdinalIgnoreCase)),
                extension.CommentBlockStart == "<!--" ||
                extension.Extensions.Any(ext =>
                    string.Equals(ext, ".html", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(ext, ".htm", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(ext, ".xml", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(ext, ".svg", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(ext, ".xaml", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(ext, ".axaml", StringComparison.OrdinalIgnoreCase)));

    }

    private enum SyntaxTokenKind
    {
        Keyword,
        Type,
        Namespace,
        Property,
        Function
    }
}

public readonly record struct CompiledSyntaxRule(Regex Regex, string ColorTokenName, string FallbackHex);

public enum EmbeddedSyntaxScanMode
{
    Normal,
    LineComment,
    BlockComment,
    String,
    MultiLineString
}

public readonly record struct EmbeddedSyntaxState(
    string BracketStack,
    EmbeddedSyntaxScanMode Mode,
    string? Delimiter,
    bool IsVerbatimString,
    bool CanSpanMultipleLines)
{
    public static EmbeddedSyntaxState Empty { get; } = new(string.Empty, EmbeddedSyntaxScanMode.Normal, null, false, false);
}

public sealed class EmbeddedSyntaxProfile
{
    private static readonly string[] CommonStringPrefixes =
    [
        "fr", "rf", "br", "rb", "ur", "ru", "cr", "rc",
        "f", "r", "u", "b", "c"
    ];

    public LoadedExtension Extension { get; }
    public IBrush CommentBrush { get; }
    public IBrush StringBrush { get; }
    public IReadOnlyList<(Regex Regex, IBrush Brush)> TokenRules { get; }

    private EmbeddedSyntaxProfile(
        LoadedExtension extension,
        IBrush commentBrush,
        IBrush stringBrush,
        IReadOnlyList<(Regex Regex, IBrush Brush)> tokenRules)
    {
        Extension = extension;
        CommentBrush = commentBrush;
        StringBrush = stringBrush;
        TokenRules = tokenRules;
    }

    public static EmbeddedSyntaxProfile Create(CompiledSyntaxProfile syntaxProfile)
    {
        var tokenRules = syntaxProfile.TokenRules
            .Select(rule => (
                rule.Regex,
                Brush.Parse(syntaxProfile.Extension.ColorTokens.TryGetValue(rule.ColorTokenName, out var hex)
                    ? hex
                    : rule.FallbackHex)))
            .ToArray();

        return new EmbeddedSyntaxProfile(
            syntaxProfile.Extension,
            Brush.Parse(syntaxProfile.Extension.ColorTokens.TryGetValue("comment", out var commentHex) ? commentHex : "#6A9955"),
            Brush.Parse(syntaxProfile.Extension.ColorTokens.TryGetValue("string", out var stringHex) ? stringHex : "#CE9178"),
            tokenRules);
    }

    public EmbeddedSyntaxState Advance(string text, EmbeddedSyntaxState initialState) =>
        Process(text, initialState, applyBrush: null, rainbowBrushResolver: null);

    public void Colorize(
        string text,
        int lineOffset,
        EmbeddedSyntaxState initialState,
        Action<int, int, IBrush> applyBrush,
        Func<int, IBrush> rainbowBrushResolver)
    {
        _ = Process(text, initialState, (start, end, brush) => applyBrush(lineOffset + start, lineOffset + end, brush), rainbowBrushResolver);
    }

    private EmbeddedSyntaxState Process(
        string text,
        EmbeddedSyntaxState initialState,
        Action<int, int, IBrush>? applyBrush,
        Func<int, IBrush>? rainbowBrushResolver)
    {
        if (string.IsNullOrEmpty(text))
            return initialState;

        var protectedRanges = new bool[text.Length];
        var stack = new Stack<char>(initialState.BracketStack.Reverse());
        var mode = initialState.Mode;
        var activeDelimiter = initialState.Delimiter;
        var isVerbatimString = initialState.IsVerbatimString;
        var canSpanMultipleLines = initialState.CanSpanMultipleLines;
        var segmentStart = mode == EmbeddedSyntaxScanMode.Normal ? -1 : 0;

        for (var index = 0; index < text.Length; index++)
        {
            if (mode == EmbeddedSyntaxScanMode.LineComment)
            {
                if (text[index] is '\r' or '\n')
                {
                    ReserveAndColor(segmentStart, index, CommentBrush);
                    mode = EmbeddedSyntaxScanMode.Normal;
                    activeDelimiter = null;
                    segmentStart = -1;
                }
                continue;
            }

            if (mode == EmbeddedSyntaxScanMode.BlockComment)
            {
                if (!string.IsNullOrEmpty(activeDelimiter) && MatchesAt(text, index, activeDelimiter))
                {
                    var end = index + activeDelimiter.Length;
                    ReserveAndColor(segmentStart, end, CommentBrush);
                    index = end - 1;
                    mode = EmbeddedSyntaxScanMode.Normal;
                    activeDelimiter = null;
                    segmentStart = -1;
                }
                continue;
            }

            if (mode is EmbeddedSyntaxScanMode.String or EmbeddedSyntaxScanMode.MultiLineString)
            {
                if (!string.IsNullOrEmpty(activeDelimiter) &&
                    IsStringTerminator(text, index, activeDelimiter, isVerbatimString))
                {
                    var end = index + activeDelimiter.Length;
                    ReserveAndColor(segmentStart, end, StringBrush);
                    index = end - 1;
                    mode = EmbeddedSyntaxScanMode.Normal;
                    activeDelimiter = null;
                    isVerbatimString = false;
                    canSpanMultipleLines = false;
                    segmentStart = -1;
                    continue;
                }

                if (!canSpanMultipleLines && text[index] is '\r' or '\n')
                {
                    ReserveAndColor(segmentStart, index, StringBrush);
                    mode = EmbeddedSyntaxScanMode.Normal;
                    activeDelimiter = null;
                    isVerbatimString = false;
                    canSpanMultipleLines = false;
                    segmentStart = -1;
                }
                continue;
            }

            if (!string.IsNullOrEmpty(Extension.CommentLine) && MatchesAt(text, index, Extension.CommentLine))
            {
                mode = EmbeddedSyntaxScanMode.LineComment;
                activeDelimiter = null;
                segmentStart = index;
                index += Extension.CommentLine.Length - 1;
                continue;
            }

            if (!string.IsNullOrEmpty(Extension.CommentBlockStart) && MatchesAt(text, index, Extension.CommentBlockStart))
            {
                mode = EmbeddedSyntaxScanMode.BlockComment;
                activeDelimiter = Extension.CommentBlockEnd;
                segmentStart = index;
                index += Extension.CommentBlockStart.Length - 1;
                continue;
            }

            if (TryMatchStringStart(text, index, out var stringStart))
            {
                mode = stringStart.CanSpanMultipleLines ? EmbeddedSyntaxScanMode.MultiLineString : EmbeddedSyntaxScanMode.String;
                activeDelimiter = stringStart.Delimiter;
                isVerbatimString = stringStart.IsVerbatim;
                canSpanMultipleLines = stringStart.CanSpanMultipleLines;
                segmentStart = index;
                index += stringStart.MatchedLength - 1;
                continue;
            }
        }

        if (mode == EmbeddedSyntaxScanMode.LineComment)
        {
            ReserveAndColor(segmentStart, text.Length, CommentBrush);
            mode = EmbeddedSyntaxScanMode.Normal;
            activeDelimiter = null;
        }
        else if (mode is EmbeddedSyntaxScanMode.String or EmbeddedSyntaxScanMode.BlockComment or EmbeddedSyntaxScanMode.MultiLineString)
        {
            if (segmentStart >= 0)
            {
                ReserveAndColor(segmentStart, text.Length, mode == EmbeddedSyntaxScanMode.BlockComment ? CommentBrush : StringBrush);
            }

            if (mode == EmbeddedSyntaxScanMode.String)
            {
                mode = EmbeddedSyntaxScanMode.Normal;
                activeDelimiter = null;
                isVerbatimString = false;
                canSpanMultipleLines = false;
            }
        }

        foreach (var (regex, brush) in TokenRules)
        {
            foreach (Match match in regex.Matches(text))
            {
                if (IsProtected(protectedRanges, match.Index, match.Index + match.Length))
                    continue;

                applyBrush?.Invoke(match.Index, match.Index + match.Length, brush);
            }
        }

        if (rainbowBrushResolver is not null)
        {
            for (var index = 0; index < text.Length; index++)
            {
                if (protectedRanges[index])
                    continue;

                var ch = text[index];
                if (TryGetClosingBracket(ch, out _))
                {
                    applyBrush?.Invoke(index, index + 1, rainbowBrushResolver(stack.Count));
                    stack.Push(ch);
                    continue;
                }

                if (TryGetOpeningBracket(ch, out var opening) &&
                    stack.Count > 0 &&
                    stack.Peek() == opening)
                {
                    applyBrush?.Invoke(index, index + 1, rainbowBrushResolver(stack.Count - 1));
                    stack.Pop();
                }
            }
        }

        return new EmbeddedSyntaxState(new string(stack.Reverse().ToArray()), mode, activeDelimiter, isVerbatimString, canSpanMultipleLines);

        void ReserveAndColor(int start, int end, IBrush brush)
        {
            if (!TryReserveRange(protectedRanges, start, end))
                return;

            applyBrush?.Invoke(start, end, brush);
        }
    }

    private bool TryMatchStringStart(string text, int index, out EmbeddedStringStart stringStart)
    {
        foreach (var candidate in Extension.MultiLineStringDelimiters
                     .Where(value => !string.IsNullOrWhiteSpace(value))
                     .Distinct()
                     .OrderByDescending(value => value.Length))
        {
            if (MatchesStringStart(text, index, candidate, out var matchedLength, out var isVerbatim, out var canSpanMultipleLines))
            {
                stringStart = new EmbeddedStringStart(candidate, matchedLength, isVerbatim, true);
                return true;
            }
        }

        foreach (var candidate in Extension.StringDelimiters
                     .Where(value => !string.IsNullOrWhiteSpace(value))
                     .Distinct()
                     .Where(value => !(Extension.DisableSingleQuoteStrings && value == "'"))
                     .OrderByDescending(value => value.Length))
        {
            if (MatchesStringStart(text, index, candidate, out var matchedLength, out var isVerbatim, out var canSpanMultipleLines))
            {
                stringStart = new EmbeddedStringStart(candidate, matchedLength, isVerbatim, canSpanMultipleLines);
                return true;
            }
        }

        stringStart = default;
        return false;
    }

    private bool MatchesStringStart(string text, int index, string delimiter, out int matchedLength, out bool isVerbatim, out bool canSpanMultipleLines)
    {
        if (MatchesAt(text, index, delimiter))
        {
            matchedLength = delimiter.Length;
            isVerbatim = false;
            canSpanMultipleLines = false;
            return true;
        }

        if (MatchesPrefixedDelimiter(text, index, delimiter, out matchedLength))
        {
            isVerbatim = false;
            canSpanMultipleLines = false;
            return true;
        }

        if (delimiter == "\"" && TryMatchCSharpStringPrefix(text, index, out matchedLength, out isVerbatim, out canSpanMultipleLines))
            return true;

        matchedLength = 0;
        isVerbatim = false;
        canSpanMultipleLines = false;
        return false;
    }

    private static bool MatchesPrefixedDelimiter(string text, int index, string delimiter, out int matchedLength)
    {
        foreach (var prefix in CommonStringPrefixes)
        {
            var candidate = prefix + delimiter;
            if (MatchesAt(text, index, candidate))
            {
                matchedLength = candidate.Length;
                return true;
            }
        }

        matchedLength = 0;
        return false;
    }

    private bool TryMatchCSharpStringPrefix(string text, int index, out int matchedLength, out bool isVerbatim, out bool canSpanMultipleLines)
    {
        if (!Extension.DisableSingleQuoteStrings)
        {
            matchedLength = 0;
            isVerbatim = false;
            canSpanMultipleLines = false;
            return false;
        }

        foreach (var candidate in new[] { "$@\"", "@$\"", "$\"", "@\"" })
        {
            if (MatchesAt(text, index, candidate))
            {
                matchedLength = candidate.Length;
                isVerbatim = candidate.Contains('@');
                canSpanMultipleLines = isVerbatim;
                return true;
            }
        }

        matchedLength = 0;
        isVerbatim = false;
        canSpanMultipleLines = false;
        return false;
    }

    private static bool IsStringTerminator(string text, int index, string delimiter, bool isVerbatimString)
    {
        if (!MatchesAt(text, index, delimiter))
            return false;

        if (!isVerbatimString || delimiter != "\"")
            return !IsEscaped(text, index);

        return !(index + 1 < text.Length && text[index + 1] == '"');
    }

    private static bool MatchesAt(string text, int index, string token) =>
        index >= 0 &&
        index + token.Length <= text.Length &&
        string.CompareOrdinal(text, index, token, 0, token.Length) == 0;

    private static bool IsEscaped(string text, int index)
    {
        var slashCount = 0;
        for (var cursor = index - 1; cursor >= 0 && text[cursor] == '\\'; cursor--)
            slashCount++;

        return (slashCount & 1) == 1;
    }

    private static bool TryReserveRange(bool[] protectedRanges, int start, int end)
    {
        if (start < 0 || end > protectedRanges.Length || start >= end)
            return false;

        if (IsProtected(protectedRanges, start, end))
            return false;

        for (var index = start; index < end; index++)
            protectedRanges[index] = true;

        return true;
    }

    private static bool IsProtected(bool[] protectedRanges, int start, int end)
    {
        for (var index = Math.Max(0, start); index < end && index < protectedRanges.Length; index++)
        {
            if (protectedRanges[index])
                return true;
        }

        return false;
    }

    private static bool TryGetClosingBracket(char ch, out char closing)
    {
        switch (ch)
        {
            case '(':
                closing = ')';
                return true;
            case '[':
                closing = ']';
                return true;
            case '{':
                closing = '}';
                return true;
            default:
                closing = default;
                return false;
        }
    }

    private static bool TryGetOpeningBracket(char ch, out char opening)
    {
        switch (ch)
        {
            case ')':
                opening = '(';
                return true;
            case ']':
                opening = '[';
                return true;
            case '}':
                opening = '{';
                return true;
            default:
                opening = default;
                return false;
        }
    }

    private readonly record struct EmbeddedStringStart(string Delimiter, int MatchedLength, bool IsVerbatim, bool CanSpanMultipleLines);
}
