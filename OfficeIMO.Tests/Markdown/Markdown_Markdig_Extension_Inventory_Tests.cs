using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_Markdig_Extension_Inventory_Tests {
    [Fact]
    public void Markdig_ExtensionInventory_Report_Is_Current() {
        string repositoryRoot = GetRepositoryRoot();
        string reportPath = Path.Combine(repositoryRoot, "Docs", "officeimo.markdown.markdig-extension-inventory.md");
        string matrixPath = Path.Combine(repositoryRoot, "Docs", "officeimo.markdown.markdig-compatibility-matrix.md");

        var report = MarkdigExtensionInventory.Build(repositoryRoot);
        string markdown = MarkdigExtensionInventoryMarkdownWriter.Write(report);
        string matrix = MarkdigExtensionCompatibilityMatrixWriter.Write(report);

        Assert.Empty(report.MissingTrackedUseMethods);
        Assert.Empty(report.ObsoleteTrackedUseMethods);
        Assert.All(report.Rows, row => Assert.False(string.IsNullOrWhiteSpace(row.Route), row.MethodName + " route is missing."));
        Assert.All(report.Rows, row => Assert.NotEqual(MarkdigExtensionScopeDecision.Unknown, row.ScopeDecision));
        Assert.All(report.Rows.Where(static row => row.Status == MarkdigExtensionInventoryStatus.Gap), row =>
            Assert.True(
                row.ScopeDecision is MarkdigExtensionScopeDecision.OptionalExtension
                    or MarkdigExtensionScopeDecision.RendererHostPolicy
                    or MarkdigExtensionScopeDecision.Deferred
                    or MarkdigExtensionScopeDecision.IntentionalDifference
                    or MarkdigExtensionScopeDecision.CoreEngine,
                row.MethodName + " gap row must have an explicit scope decision."));
        Assert.All(report.Rows, row => Assert.False(string.IsNullOrWhiteSpace(row.PromotionBar), row.MethodName + " promotion bar is missing."));

        if (string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_MARKDIG_INVENTORY"), "1", StringComparison.Ordinal)) {
            File.WriteAllText(reportPath, markdown);
            File.WriteAllText(matrixPath, matrix);
        }

        Assert.True(File.Exists(reportPath), "Markdig extension inventory report is missing: " + reportPath);
        Assert.Equal(NormalizeLineEndings(File.ReadAllText(reportPath)), NormalizeLineEndings(markdown));
        Assert.True(File.Exists(matrixPath), "Markdig extension compatibility matrix is missing: " + matrixPath);
        Assert.Equal(NormalizeLineEndings(File.ReadAllText(matrixPath)), NormalizeLineEndings(matrix));
        AssertDocsTrackInventoryCounts(report, repositoryRoot);
    }

    private static void AssertDocsTrackInventoryCounts(MarkdigExtensionInventoryReport report, string repositoryRoot) {
        string rowText = $"{report.Total} Markdig extension-family rows";
        string statusText = $"{report.Partial} partial";

        string compatibilityMatrix = File.ReadAllText(Path.Combine(repositoryRoot, "Docs", "officeimo.markdown.compatibility-matrix.md"));
        Assert.Contains(rowText, compatibilityMatrix, StringComparison.Ordinal);
        Assert.Contains(statusText, compatibilityMatrix, StringComparison.Ordinal);

        string markdigCompatibilityMatrix = File.ReadAllText(Path.Combine(repositoryRoot, "Docs", "officeimo.markdown.markdig-compatibility-matrix.md"));
        Assert.Contains(rowText, markdigCompatibilityMatrix, StringComparison.Ordinal);
        Assert.Contains(statusText, markdigCompatibilityMatrix, StringComparison.Ordinal);
        Assert.Contains("Engine parser", markdigCompatibilityMatrix, StringComparison.Ordinal);
        Assert.Contains("AST/source", markdigCompatibilityMatrix, StringComparison.Ordinal);
        Assert.Contains("Writer/render", markdigCompatibilityMatrix, StringComparison.Ordinal);
        Assert.Contains("semantic HeadingBlock level/text source spans", markdigCompatibilityMatrix, StringComparison.Ordinal);
        Assert.Contains("semantic ImageBlock source spans", markdigCompatibilityMatrix, StringComparison.Ordinal);
        Assert.Contains("semantic CodeBlock and SemanticFencedBlock info/content source spans", markdigCompatibilityMatrix, StringComparison.Ordinal);

        string parityGapPlan = File.ReadAllText(Path.Combine(repositoryRoot, "Docs", "officeimo.markdown.markdig-parity-gap-plan.md"));
        Assert.Contains(rowText, parityGapPlan, StringComparison.Ordinal);
        Assert.Contains("Markdig extension inventory", parityGapPlan, StringComparison.Ordinal);
        Assert.Contains("Markdig extension compatibility matrix", parityGapPlan, StringComparison.Ordinal);
        Assert.Contains("Route", parityGapPlan, StringComparison.Ordinal);
        Assert.Contains("Scope decision", parityGapPlan, StringComparison.Ordinal);
        Assert.Contains("semantic HeadingBlock level/text source spans", parityGapPlan, StringComparison.Ordinal);
        Assert.Contains("semantic ImageBlock source spans", parityGapPlan, StringComparison.Ordinal);
        Assert.Contains("semantic CodeBlock and SemanticFencedBlock info/content source spans", parityGapPlan, StringComparison.Ordinal);
    }

    private static string GetRepositoryRoot() {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null) {
            if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.sln"))) {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Unable to locate OfficeIMO repository root from test runtime base directory.");
    }

    private static string NormalizeLineEndings(string value) =>
        value.Replace("\r\n", "\n").Replace("\r", "\n");
}
