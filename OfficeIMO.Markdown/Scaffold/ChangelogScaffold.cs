namespace OfficeIMO.Markdown;

/// <summary>
/// Scaffolds a simple Keep a Changelog-compliant CHANGELOG.md skeleton.
/// </summary>
public static class MarkdownChangelogScaffold {
    /// <summary>Builds a basic CHANGELOG document.</summary>
    /// <param name="projectName">Optional project name for the header.</param>
    public static MarkdownDocument Changelog(string? projectName = null) {
        var title = string.IsNullOrWhiteSpace(projectName) ? "Changelog" : projectName + " Changelog";
        var md = MarkdownDocument.Create()
            .H1(title)
            .P("All notable changes to this project will be documented in this file.")
            .P("The format is based on Keep a Changelog and this project adheres to Semantic Versioning.")
            .H2("Unreleased")
            .Ul(ul => ul.Item("Added").Item("Changed").Item("Fixed"));
        return md;
    }
}
