# Contributing to Kodo

Thanks for your interest in contributing! Kodo is a small project built by two people, so every contribution genuinely matters. This document covers everything you need to know to get started.

---

## Getting Set Up

To run Kodo locally, you'll need:

1. [.NET](https://dotnet.microsoft.com/en-us/download) minimum version 8
2. Run `dotnet new install Avalonia.Templates`
3. Change your directory to Kodo's source folder: `cd path\to\Kodo\Kodo`
4. Run `dotnet run` — it'll take a few seconds then open up

That's it. No complicated build pipeline, no extra tools.

---

## How to Contribute

### Reporting Bugs

If something's broken, open an [Issue](https://github.com/KerbalMissile/Kodo/issues) and describe:
- What you were doing when it happened
- What you expected to happen
- What actually happened
- Your OS and .NET version if relevant

We're in Beta, so expect bugs — but that also means bug reports are especially valuable right now.

### Submitting Changes

1. Fork the repo
2. Make your changes on a new branch
3. Open a Pull Request with a clear description of what you changed and why

Please keep PRs focused — one thing per PR makes it much easier to review. If you're planning something large, open an Issue first so we can discuss it before you put in the work.

### Code Style

Look at the existing code and match it. A few things to keep consistent:

- Add a comment at the top of any file you modify in the format the codebase already uses:
  ```
  // April 24th, 2026 - YourName - Brief description of what you changed
  ```
- Keep comments descriptive but not excessive — explain *why*, not just *what*
- Don't leave dead code or commented-out blocks behind

---

## Making Extensions

Extensions are the best way to contribute without touching the core app. A `.kox` file is just a renamed `.zip` containing:

```
manifest.json    ← required
language.json    ← for language extensions
theme.json       ← for theme extensions
icon.png         ← optional, must be square
```

**manifest.json** example:
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

**language.json** supports: `keywords`, `types`, `commentLine`, `commentBlockStart`, `commentBlockEnd`, and `colorTokens` (`keyword`, `type`, `string`, `comment`, `number`).

**theme.json** supports: `themeId`, `displayName`, `baseTheme` (`"Dark"` or `"Light"`), and color keys for `windowBackground`, `topBar`, `sidebar`, `button`, `buttonHover`, `editorBackground`, `card`, `primaryText`, `mutedText`, `surfaceBorder`, `accent`, `previewBackground`, `previewBorder`.

If you want your extension in the official marketplace, open a PR adding your `.kox` to `Official_Extensions/` and the relevant entry to `Indexs/ExtensionsIndex.json`.

---

## License

By contributing, you agree that your contributions are licensed under the [Kodo Public License v1.1](https://github.com/KerbalMissile/Kodo/blob/main/LICENSE.md). In short: credit must be given, modifications must stay under KPL-v1.1, and if you distribute anything containing Kodo's code you need to let us know within 30 days.

---

## Questions?

Jump into the [Discord](https://discord.gg/cUQ6C88Z9C), we're always here to help.