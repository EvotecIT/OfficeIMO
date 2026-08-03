using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_GitHubFlavoredMarkdown_Inventory_Tests {
    [Fact]
    public void Gfm_TrackedExtensionInventory_Report_Is_Current() {
        string fixturePath = GetTestProjectPath("Markdown", "Fixtures", "GitHubFlavoredMarkdown", "cmark-gfm-extensions-smoke.json");
        string reportPath = GetRepositoryPath("Docs", "officeimo.markdown.gfm-inventory.md");

        var report = GfmInventory.Build(fixturePath);
        string markdown = GfmInventoryMarkdownWriter.Write(report);

        if (string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_GFM_INVENTORY"), "1", StringComparison.Ordinal)) {
            File.WriteAllText(reportPath, markdown);
        }

        Assert.True(File.Exists(reportPath), "GFM inventory report is missing: " + reportPath);
        Assert.Equal(NormalizeLineEndings(File.ReadAllText(reportPath)), NormalizeLineEndings(markdown));
    }

    private static string GetTestProjectPath(params string[] segments) {
        string path = Path.Combine(AppContext.BaseDirectory, "..", "..", "..");
        foreach (string segment in segments) {
            path = Path.Combine(path, segment);
        }

        return Path.GetFullPath(path);
    }

    private static string GetRepositoryPath(params string[] segments) {
        string path = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..");
        foreach (string segment in segments) {
            path = Path.Combine(path, segment);
        }

        return Path.GetFullPath(path);
    }

    private static string NormalizeLineEndings(string value) =>
        value.Replace("\r\n", "\n").Replace("\r", "\n");

}
