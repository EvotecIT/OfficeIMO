using System;
using System.Collections.Generic;

namespace OfficeIMO.Markdown;

/// <summary>
/// Helpers to scaffold common Markdown files such as README.
/// </summary>
public static class Scaffold {
    /// <summary>
    /// Builds a README document with common sections using provided options.
    /// </summary>
    public static MarkdownDoc Readme(string projectName, Action<ReadmeOptions> configure) {
        var opts = new ReadmeOptions(projectName);
        configure(opts);

        var md = MarkdownDoc.Create()
            .H1(projectName)
            .P(opts.DescriptionText ?? "");

        if (opts.BadgeList.Count > 0) {
            md.P(p => {
                for (int i = 0; i < opts.BadgeList.Count; i++) {
                    var b = opts.BadgeList[i];
                    p.Link(b.Text, b.Url, b.Title);
                }
            });
        }

        if (!string.IsNullOrWhiteSpace(opts.NuGetPackageId)) {
            md.H2("Install")
              .Code("bash", string.IsNullOrWhiteSpace(opts.InstallCommand)
                    ? $"dotnet add package {opts.NuGetPackageId}"
                    : opts.InstallCommand!);
        }

        if (!string.IsNullOrWhiteSpace(opts.GettingStartedCode)) {
            md.H2("Getting started")
              .Code("csharp", opts.GettingStartedCode!);
        }

        if (opts.LinkList.Count > 0) {
            md.H2("Links").Ul(ul => {
                foreach (var l in opts.LinkList) ul.ItemLink(l.Text, l.Url);
            });
        }

        if (opts.LicenseMITRequested) {
            md.H2("License").P("MIT");
        }

        return md;
    }
}

/// <summary>
/// Options for README scaffolding.
/// </summary>
public sealed class ReadmeOptions {
    internal string ProjectName { get; }
    internal string? NuGetPackageId { get; private set; }
    internal string? DescriptionText { get; private set; }
    internal string? InstallCommand { get; private set; }
    internal string? GettingStartedCode { get; private set; }
    internal bool LicenseMITRequested { get; private set; }
    internal List<(string Text, string Url, string? Title)> BadgeList { get; } = new();
    internal List<(string Text, string Url)> LinkList { get; } = new();

    /// <summary>Create options for project.</summary>
    public ReadmeOptions(string projectName) { ProjectName = projectName; }

    /// <summary>Sets the NuGet package id.</summary>
    public ReadmeOptions NuGet(string packageId) { NuGetPackageId = packageId; return this; }
    /// <summary>Sets the README description paragraph.</summary>
    public ReadmeOptions Description(string description) { DescriptionText = description; return this; }
    /// <summary>Sets install command and getting-started code snippet.</summary>
    public ReadmeOptions GettingStarted(string installCommand, string codeBlock) { InstallCommand = installCommand; GettingStartedCode = codeBlock; return this; }
    /// <summary>Adds MIT license section.</summary>
    public ReadmeOptions LicenseMIT() { LicenseMITRequested = true; return this; }
    /// <summary>Adds links section entries.</summary>
    public ReadmeOptions Links(params (string Text, string Url)[] links) { LinkList.AddRange(links); return this; }
    /// <summary>Adds badges via builder.</summary>
    public ReadmeOptions Badges(Action<BadgeBuilder> build) { var b = new BadgeBuilder(BadgeList); build(b); return this; }
}

/// <summary>
/// Helper to add common badges to README.
/// </summary>
public sealed class BadgeBuilder {
    private readonly List<(string Text, string Url, string? Title)> _badges;
    internal BadgeBuilder(List<(string Text, string Url, string? Title)> list) { _badges = list; }
    /// <summary>Adds a custom badge label linked to a URL.</summary>
    public BadgeBuilder Custom(string text, string url, string? title = null) { _badges.Add((text, url, title)); return this; }
    // Minimal helpers; users can use Custom to inject their own
    /// <summary>Shortcut for a NuGet badge link.</summary>
    public BadgeBuilder NuGet(string? id = null) {
        if (!string.IsNullOrWhiteSpace(id)) _badges.Add(("NuGet", $"https://www.nuget.org/packages/{id}", null));
        else _badges.Add(("NuGet", "https://www.nuget.org/", null));
        return this;
    }
    /// <summary>Adds a CI build badge link.</summary>
    public BadgeBuilder Build(string? url = null) { _badges.Add(("Build", url ?? "https://github.com/EvotecIT/OfficeIMO/actions", null)); return this; }
    /// <summary>Adds a code coverage badge link.</summary>
    public BadgeBuilder Coverage(string? url = null) { _badges.Add(("Coverage", url ?? "https://codecov.io/", null)); return this; }
}
