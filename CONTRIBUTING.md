# Contributing to Kodo

Thanks for your interest in contributing! Kodo is a small project built by two people, so every contribution genuinely matters. This document covers everything you need to know to get started.

---

## Getting Set Up

To run Kodo locally, you'll need:

1. [.NET](https://dotnet.microsoft.com/en-us/download) minimum version 10
2. Run `dotnet new install Avalonia.Templates`
3. Change your directory to Kodo's source folder: `cd path\to\Kodo\Source`
4. Run `dotnet run` - it'll take a few seconds then open up

That's it. No complicated build pipeline, no extra tools.

---

## How to Contribute

### Reporting Bugs

If something's broken, open an [Issue](https://github.com/KerbalMissile/Kodo/issues) and describe:
- What you were doing when it happened
- What you expected to happen
- What actually happened
- Your OS and .NET version if relevant

### Submitting Changes

1. Fork the repo
2. Make your changes on a new branch
3. Open a Pull Request with a clear description of what you changed and why

Please keep PRs focused, one thing per PR makes it much easier to review. If you're planning something large, open an Issue first so we can discuss it before you put in the work.

### Code Style

Look at the existing code and match it. A few things to keep consistent:

- Instead of a per-file changelog comment, make a good PR - a clear title and description of what you changed and why goes a long way for review.
- Keep comments descriptive but not excessive - explain *why*, not just *what*
- Don't leave dead code or commented-out blocks behind

---

## Making Extensions

Extensions are the best way to contribute without touching the core app. A `.kox` file is just a renamed `.zip` containing:

```
manifest.json     ← required
language.json     ← for language extensions
language2.json    ← optional, a second language profile in the same package
theme.json        ← for theme extensions
icon.png/icon.svg ← optional, must be square
```

### manifest.json

Required for all extensions:

```json
{
  "id": "mylang-kodo-extension",
  "version": "v1.0.0",
  "name": "MyLang Language Support",
  "type": "language",
  "author": "Your Name",
  "description": "Syntax highlighting for MyLang files.",
  "extensions": [".myl"]
}
```

### language.json

All fields are optional unless noted. Below is a full example followed by field descriptions.

```json
{
  "keywords": ["if", "else", "return", "while"],
  "types": ["int", "string", "bool", "MyClass"],
  "functions": [],
  "properties": [],
  "namespaces": [],
  "commentLine": "//",
  "commentBlockStart": "/*",
  "commentBlockEnd": "*/",
  "stringDelimiters": ["\"", "'"],
  "multiLineStringDelimiters": ["\"\"\""],
  "disableSingleQuoteStrings": false,
  "colorTokens": {
    "keyword":      "#569CD6",
    "type":         "#4EC9B0",
    "string":       "#CE9178",
    "comment":      "#6A9955",
    "number":       "#B5CEA8",
    "operator":     "#D4D4D4",
    "punctuation":  "#D4D4D4",
    "function":     "#DCDCAA",
    "property":     "#9CDCFE",
    "namespace":    "#4FC1FF",
    "attribute":    "#C586C0",
    "preprocessor": "#C586C0",
    "variable":     "#A0DBFD",
    "charLiteral":  "#CE9178"
  }
}
```

**Token lists**

| Field | Description |
|---|---|
| `keywords` | Reserved words colored with `keyword` color (e.g. `if`, `return`, `class`) |
| `types` | Type names colored with `type` color (e.g. `int`, `string`, built-in classes) |
| `functions` | Function/method names colored with `function` color |
| `properties` | Property/field names colored with `property` color |
| `namespaces` | Namespace/module names colored with `namespace` color |

All five accept a JSON array of strings. Word-boundary matching is applied automatically, so `"int"` won't match inside `"integer"`.

**Comment delimiters**

| Field | Description |
|---|---|
| `commentLine` | Single-character or string that starts a line comment (e.g. `"//"`, `"#"`, `">"`) |
| `commentBlockStart` | Opening delimiter for block comments (e.g. `"/*"`) |
| `commentBlockEnd` | Closing delimiter for block comments (e.g. `"*/"`) |

**String delimiters**

| Field | Description |
|---|---|
| `stringDelimiters` | Single-line string delimiters (e.g. `["\"", "'"]`). The span ends at the closing delimiter or end of line. |
| `multiLineStringDelimiters` | Multi-line string delimiters (e.g. `["\"\"\"", "'''", "` ``` `"]`). The span continues across lines until the closing delimiter is found. List longer delimiters first. |
| `disableSingleQuoteStrings` | Set to `true` to replace the open-ended `'…'` span with a precise char-literal regex. Use for languages like C# where `'` appears in non-string contexts. Defaults to `false`. |

**colorTokens**

Overrides the default colors for any highlighting category. All values are hex color strings. Available keys:

`keyword`, `type`, `string`, `comment`, `number`, `operator`, `punctuation`, `function`, `property`, `namespace`, `attribute`, `preprocessor`, `variable`, `charLiteral`

Any key you omit falls back to the built-in default for that category.

### theme.json

Supports: `themeId`, `displayName`, `baseTheme` (`"Dark"` or `"Light"`), and color keys for `windowBackground`, `topBar`, `sidebar`, `button`, `buttonHover`, `editorBackground`, `card`, `primaryText`, `mutedText`, `surfaceBorder`, `accent`, `previewBackground`, `previewBorder`.

---

## Submitting an Extension

If you want your extension in the official marketplace, open a PR adding your `.kox` to `Official_Extensions/` and the relevant entry to `Indexs/ExtensionsIndex.json`.

---

## License

By contributing, you agree that your contributions are licensed under the [Kodo Public License v1.1](https://github.com/KerbalMissile/Kodo/blob/main/LICENSE.md). In short: credit must be given, modifications must stay under KPL-v1.1, and if you distribute anything containing Kodo's code you need to let us know within 30 days.

---

## Questions?

Jump into the [Discord](https://discord.gg/cUQ6C88Z9C), we're always here to help.