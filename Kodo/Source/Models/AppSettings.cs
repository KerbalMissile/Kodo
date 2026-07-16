using System;
using System.Collections.Generic;

namespace Kodo;

/// <summary>
/// The full schema Kodo persists to <c>kodosettings.json</c>. Pure data - no I/O, no
/// defaults resolution logic beyond simple property initializers. MainWindow owns
/// reading/writing this to disk (LoadSettings / BuildSettingsSnapshot / PersistSettingsSnapshot);
/// this type only defines the shape.
/// </summary>
internal sealed class AppSettings
{
    // Kept here (rather than on MainWindow) because it's a settings value, not just a
    // UI bound - this is the single source of truth other code should reference.
    public const double DefaultTerminalPanelHeight = 300;

    public string ThemeName { get; set; } = "Dark";
    public bool AutoSaveEnabled { get; set; }
    public bool DiscordRichPresenceEnabled { get; set; }
    public bool DiscordImprovedRpcEnabled { get; set; }
    public bool DeveloperOptionsVisible { get; set; }
    public bool VerboseLoggingEnabled { get; set; }
    public bool StatusBarFilePathVisible { get; set; } = true;
    public bool WordWrapEnabled { get; set; }
    // Predictive completion (CodePredict). Defaults to true - on unless the user disables it.
    public bool CodePredictEnabled { get; set; } = true;
    public int TabSize { get; set; } = 4;
    public int EditorFontSize { get; set; } = 14;
    public bool ConfirmBeforeClosingUnsavedTabsEnabled { get; set; } = true;
    public bool RestoreOpenTabsOnLaunchEnabled { get; set; }
    public bool AutoUpdateExtensionsEnabled { get; set; }
    // Sub-setting under AutoUpdateExtensionsEnabled - see
    // IsAutoUpdateExtensionsInBackgroundEnabled for what it controls.
    public bool AutoUpdateExtensionsInBackgroundEnabled { get; set; }
    // Defaults to true - most users want Kodo to stay current without thinking about it.
    public bool AutoUpdateAppEnabled { get; set; } = true;
    // Sub-setting: defaults to false so Update Now/Later still shows.
    public bool AutoUpdateAppInBackgroundEnabled { get; set; }
    public string? PreferredTerminalShellId { get; set; }
    public bool TerminalVisible { get; set; }
    public double TerminalPanelHeight { get; set; } = DefaultTerminalPanelHeight;
    public List<string> OpenTabPaths { get; set; } = [];
    public string? ActiveTabPath { get; set; }
    public List<RecentFileEntry> RecentFiles { get; set; } = [];
    // False on first launch (settings file didn't exist yet); set to true after the
    // tutorial is dismissed so it never shows again on subsequent launches.
    public bool HasCompletedTutorial { get; set; }
    public string AccentColorMode { get; set; } = "kodo";
    public string CustomAccentHex { get; set; } = "#8C00FF";
    // Personalization - optional; empty/0 means "use OS defaults".
    public string? UserCountry { get; set; }
    public int UserHemisphere { get; set; }
    public string? UserTimezoneOffset { get; set; }
    public string? UserName { get; set; }
    public string? LastSeenVersion { get; set; }
    // Anonymous usage-analytics opt-in. False (no tracking) until the user
    // has explicitly responded to the consent prompt at least once.
    public bool AllowDataTracking { get; set; }
    public bool HasRespondedToDataTrackingPrompt { get; set; }
    // Acknowledgment of the embedded Privacy Policy text - separate from the data-tracking
    // opt-in above. There's no decline path; this only tracks whether the user has scrolled
    // through and accepted the terms at least once.
    public bool HasAcceptedPrivacyPolicy { get; set; }
}

/// <summary>One entry in AppSettings.RecentFiles - a recently opened file or folder.</summary>
public sealed class RecentFileEntry
{
    public string Path { get; set; } = string.Empty;
    public bool IsFolder { get; set; }
    public DateTime LastOpened { get; set; } = DateTime.Now;
    public bool IsPinned { get; set; }
}