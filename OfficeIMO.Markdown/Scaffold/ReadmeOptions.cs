namespace OfficeIMO.Markdown;

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
    internal List<(string Alt, string LinkUrl, string ImageUrl)> BadgeImageList { get; } = new();
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
    public ReadmeOptions Badges(Action<BadgeBuilder> build) { var b = new BadgeBuilder(BadgeList, BadgeImageList); build(b); return this; }
}
