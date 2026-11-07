using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPaintedTableRowTests {
    [Fact]
    public void ComparisonReportKeepsShadedSummaryPairsAndWrappedFindingsInTheirCells() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "Website", "static", "downloads", "showcase",
            "features", "word-comparison", "comparison-report.pdf");
        using var pdf = PdfDocument.Load(path);
        PdfDocumentReadResult result = pdf.Read();

        PdfLogicalTable summary = Assert.Single(result.Pages[0].Tables, table =>
            table.Rows.Any(row => row.Count == 2 && row[0] == "Findings"));
        Assert.Equal(6, summary.Rows.Count);
        Assert.Contains(summary.Rows, row => row.SequenceEqual(new[] { "Findings", "18" }));
        Assert.Contains(summary.Rows, row => row.SequenceEqual(new[] { "Limitations", "1" }));
        Assert.Contains("Findings\t18", result.Text, StringComparison.Ordinal);
        Assert.Contains("Limitations\t1", result.Text, StringComparison.Ordinal);

        PdfLogicalTable findings = Assert.Single(result.Pages[1].Tables);
        Assert.Equal(11, findings.Rows.Count);
        Assert.Equal(new[] { "#", "Scope", "Change", "Location", "Detailed Location", "Source", "Target", "Message" },
            findings.Rows[0]);
        IReadOnlyList<string> supplier = Assert.Single(findings.Rows, row => row[0] == "2");
        Assert.Equal("The supplier must respond within 48 hours.", supplier[5]);
        Assert.Equal("The supplier must respond within 24 hours.", supplier[6]);
        Assert.Contains("| The supplier must respond within 48 hours. | The supplier must respond within 24 hours. |",
            result.ExtractMarkdown(), StringComparison.Ordinal);
    }
}
