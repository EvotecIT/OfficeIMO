using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_CommonMark_Inventory_Tests {
    [Fact]
    public void CommonMark_FullInventory_Report_Is_Current() {
        string officialSpecPath = GetTestProjectPath("Markdown", "Fixtures", "CommonMark", "commonmark-0.31.2-spec.json");
        string pinnedFixturePath = GetTestProjectPath("Markdown", "Fixtures", "CommonMark", "commonmark-0.31.2-smoke.json");
        string reportPath = GetRepositoryPath("Docs", "officeimo.markdown.commonmark-inventory.md");

        var report = CommonMarkInventory.Build(officialSpecPath, pinnedFixturePath);
        string markdown = CommonMarkInventoryMarkdownWriter.Write(report);

        if (string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_COMMONMARK_INVENTORY"), "1", StringComparison.Ordinal)) {
            File.WriteAllText(reportPath, markdown);
        }

        Assert.True(File.Exists(reportPath), "CommonMark inventory report is missing: " + reportPath);
        Assert.Equal(NormalizeLineEndings(File.ReadAllText(reportPath)), NormalizeLineEndings(markdown));
        AssertDocsTrackInventoryCounts(report);
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

    private static void AssertDocsTrackInventoryCounts(CommonMarkInventoryReport report) {
        string inventoryText = $"{report.PassingPinned + report.PassingUnpinned} of {report.Total} official CommonMark `0.31.2` examples";
        string failureText = $"{report.Failing} failures";

        string compatibilityMatrix = File.ReadAllText(GetRepositoryPath("Docs", "officeimo.markdown.compatibility-matrix.md"));
        Assert.Contains(inventoryText, compatibilityMatrix, StringComparison.Ordinal);
        Assert.Contains(failureText, compatibilityMatrix, StringComparison.Ordinal);

        string parityGapPlan = File.ReadAllText(GetRepositoryPath("Docs", "officeimo.markdown.markdig-parity-gap-plan.md"));
        Assert.Contains(inventoryText, parityGapPlan, StringComparison.Ordinal);
        Assert.Contains($"{report.Failing} are failing", parityGapPlan, StringComparison.Ordinal);
    }
}
