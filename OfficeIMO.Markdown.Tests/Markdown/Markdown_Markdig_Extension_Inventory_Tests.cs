using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_Markdig_Extension_Inventory_Tests {
    [Fact]
    public void Markdig_ExtensionInventory_Contract_Is_Current() {
        string repositoryRoot = GetRepositoryRoot();

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
        Assert.Contains($"| Markdig extension-family rows | {report.Total} |", markdown, StringComparison.Ordinal);
        Assert.Contains($"| Partial | {report.Partial} |", markdown, StringComparison.Ordinal);
        Assert.Contains("Engine parser", matrix, StringComparison.Ordinal);
        Assert.Contains("AST/source", matrix, StringComparison.Ordinal);
        Assert.Contains("Writer/render", matrix, StringComparison.Ordinal);

        string publishedMatrix = File.ReadAllText(Path.Combine(
            repositoryRoot,
            "Docs",
            "officeimo.markdown.compatibility-matrix.md"));
        Assert.Contains($"| Extension-family rows | {report.Total} |", publishedMatrix, StringComparison.Ordinal);
        Assert.Contains($"| Covered | {report.Covered} |", publishedMatrix, StringComparison.Ordinal);
        Assert.Contains($"| Partial | {report.Partial} |", publishedMatrix, StringComparison.Ordinal);
        Assert.Contains($"| Intentional | {report.Intentional} |", publishedMatrix, StringComparison.Ordinal);
        Assert.Contains($"| Gap | {report.Gap} |", publishedMatrix, StringComparison.Ordinal);
        Assert.All(report.Rows, row => Assert.Contains(
            $"| {row.Family} | `{row.Status}` | {row.Route} |",
            publishedMatrix,
            StringComparison.Ordinal));
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

}
