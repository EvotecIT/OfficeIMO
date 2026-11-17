using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfColumnFlowProseTableTests {
    [Fact]
    public void TwoColumnPaperBodyStaysProseWhileItsTablesRemainTables() {
        // A LaTeX two-column paper: citation years give its body lines digits, and both page columns share baselines.
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Pdf", "IndependentCorpus",
            "acl-2021-mt5-multilingual.pdf");
        PdfDocumentReadResult result = PdfDocument.Load(path, new PdfLoadOptions {
            PageSelection = PdfPageSelection.From(1, 2, 3, 4)
        }).Read();

        PdfLogicalPage first = result.Pages[0];
        Assert.Empty(first.Tables);
        string markdown = Regex.Replace(first.ExtractMarkdown(), @"\s+", " ");
        // Abstract (left column) and introduction (right column) sentences read through instead of alternating per line.
        Assert.Contains("state-of-the-art re- sults on a wide variety of English-language NLP tasks.", markdown, StringComparison.Ordinal);
        Assert.Contains("has been to release dozens of models, each pre-trained on a single non-English language", markdown, StringComparison.Ordinal);
        Assert.Contains("Popular models of this type are mBERT (Devlin, 2018), mBART (Liu et al., 2020a), and XLM-R", markdown, StringComparison.Ordinal);
        // The left column's introduction precedes the right column's continuation.
        int left = markdown.IndexOf("Current natural language processing", StringComparison.Ordinal);
        int right = markdown.IndexOf("In this paper, we continue this tradition", StringComparison.Ordinal);
        Assert.True(left >= 0 && right > left, markdown);

        // The paper's own tables, including the borderless numeric result tables, are still recovered.
        Assert.Contains(result.Pages[2].Tables, table => table.Rows.Any(row => row.Contains("Architecture")));
        Assert.Contains(result.Pages[3].Tables, table => table.Rows.Any(row => row.Contains("XLM-R") && row.Contains("82.6")));
    }
}
