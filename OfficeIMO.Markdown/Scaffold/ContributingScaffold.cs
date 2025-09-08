namespace OfficeIMO.Markdown;

/// <summary>
/// Scaffolds a CONTRIBUTING.md skeleton.
/// </summary>
public static class ContributingScaffold {
    /// <summary>Builds a basic CONTRIBUTING document.</summary>
    /// <param name="projectName">Optional project name used in the intro sentence.</param>
    public static MarkdownDoc Contributing(string? projectName = null) {
        var name = string.IsNullOrWhiteSpace(projectName) ? "this project" : projectName;
        var md = MarkdownDoc.Create()
            .H1("Contributing")
            .P($"Thanks for considering contributing to {name}!")
            .H2("How to contribute")
            .Ol(ol => ol
                .Item("Fork the repository")
                .Item("Create a feature branch")
                .Item("Make your changes with tests")
                .Item("Open a pull request"))
            .H2("Development setup")
            .Code("bash", "dotnet build\ndotnet test")
            .H2("Code of Conduct")
            .P("By participating, you are expected to uphold our Code of Conduct.");
        return md;
    }
}
