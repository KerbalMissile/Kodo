// Licensed under GPL-v3.0
using System;

namespace Kodo;

internal static class GitHubRepoInfo
{
    private const string OrganizationOwner = "Kodo-IDE";
    private const string PersonalOwner = "KerbalMissile";
    private const string RepoName = "Kodo";
    private const string ExtensionsRepoName = "Kodo-Extensions";
    private const string OfficialExtensionsFolder = "Official_Extensions";
    private const string SharedExtensionsFolder = "Extensions";

    public static readonly string[] Owners = [OrganizationOwner, PersonalOwner];
    public static string OrganizationRepoUrl => BuildRepoUrl(OrganizationOwner);
    public static string PersonalRepoUrl => BuildRepoUrl(PersonalOwner);
    public static string ReleaseNotesUrl => BuildReleasesUrl(OrganizationOwner);
    public static string IssuesUrl => BuildIssuesUrl(OrganizationOwner);
    public static string MarketplaceIndexUrl => BuildExtensionsContentsUrl(OrganizationOwner, "Indexs/ExtensionsIndex.json");
    public static string AnnouncementsUrl => BuildContentsUrl(OrganizationOwner, "Announcements/ANNOUNCEMENTS.md");
    public static string ExtensionsRepoUrl => BuildExtensionsRepoUrl(OrganizationOwner);
    public static string OfficialExtensionsRepoUrl => BuildExtensionsRepoUrl(PersonalOwner);
    public static string ExtensionsFolderName => SharedExtensionsFolder;
    public static string OfficialExtensionsFolderName => OfficialExtensionsFolder;
    public static string PrivacyPolicyUrl => BuildBlobUrl(OrganizationOwner, "main", "Policies/PRIVACY%20POLICY.txt");
    public static string UserAgent => $"Kodo/1.0.0-DEV (https://github.com/{OrganizationOwner}/{RepoName})";

    public static string LatestReleaseApiUrl => BuildLatestReleaseApiUrl(OrganizationOwner);
    public static string ExtensionsIndexUrl => BuildExtensionsContentsUrl(OrganizationOwner, "Indexs/ExtensionsIndex.json");
    public static string FallbackExtensionsIndexUrl => BuildContentsUrl(PersonalOwner, "Indexs/ExtensionsIndex.json");

    public static string GetLatestReleaseApiUrl(string owner) => BuildLatestReleaseApiUrl(owner);
    public static string GetReleasesApiUrl(string owner) => BuildReleasesApiUrl(owner);
    public static string GetExtensionsIndexUrl(string owner) =>
        owner.Equals(PersonalOwner, StringComparison.OrdinalIgnoreCase)
            ? BuildContentsUrl(PersonalOwner, "Indexs/ExtensionsIndex.json")
            : BuildExtensionsContentsUrl(owner, "Indexs/ExtensionsIndex.json");

    public static string GetFallbackReleaseNotesUrl() => BuildReleasesUrl(PersonalOwner);

    public static string GetIssueUrl(string title, string body) =>
        BuildIssuesUrl(OrganizationOwner) + $"?title={Uri.EscapeDataString(title)}&body={Uri.EscapeDataString(body)}";

    private static string BuildRepoUrl(string owner) => $"https://github.com/{owner}/{RepoName}";

    private static string BuildReleasesUrl(string owner) => $"{BuildRepoUrl(owner)}/releases";

    private static string BuildIssuesUrl(string owner) => $"{BuildRepoUrl(owner)}/issues/new";

    private static string BuildLatestReleaseApiUrl(string owner) =>
        $"https://api.github.com/repos/{owner}/{RepoName}/releases/latest";

    private static string BuildReleasesApiUrl(string owner) =>
        $"https://api.github.com/repos/{owner}/{RepoName}/releases";

    private static string BuildContentsUrl(string owner, string path) =>
        $"https://api.github.com/repos/{owner}/{RepoName}/contents/{path}";

    private static string BuildExtensionsRepoUrl(string owner) =>
        $"https://github.com/{owner}/{ExtensionsRepoName}";

    private static string BuildExtensionsContentsUrl(string owner, string path) =>
        $"https://api.github.com/repos/{owner}/{ExtensionsRepoName}/contents/{path}";

    private static string BuildBlobUrl(string owner, string branch, string path) =>
        $"https://github.com/{owner}/{RepoName}/blob/{branch}/{path}";
}
