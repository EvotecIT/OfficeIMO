using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfLogicalTableTextExportTests {
    [Fact]
    public void PdfTables_ExtractMarkdownTables_AppliesPageRangesAndRowCaps() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .KeyValueTable(new[] {
                PdfKeyValueRow.Text("InvoiceId", "INV-001"),
                PdfKeyValueRow.Text("Customer", "Evotec"),
                PdfKeyValueRow.Text("Due", "2026-06-30")
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 120, 170 },
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .PageBreak()
            .Paragraph(p => p.Text("No table on this page."))
            .ToBytes();

        string markdown = PdfLogicalTableTextExportExtensions.ExtractMarkdownTables(
            pdf,
            new PdfLogicalTableTextExportOptions {
                LayoutOptions = new PdfTextLayoutOptions {
                    ForceSingleColumn = true
                },
                PageRanges = new[] { PdfPageRange.From(1, 1) },
                MaxRows = 2
            });

        Assert.Contains("### PDF page 1, table 1", markdown, StringComparison.Ordinal);
        Assert.Contains("| Key | Value |", markdown, StringComparison.Ordinal);
        Assert.Contains("| InvoiceId | INV-001 |", markdown, StringComparison.Ordinal);
        Assert.Contains("| Customer | Evotec |", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("2026-06-30", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("No table on this page", markdown, StringComparison.Ordinal);

        string empty = PdfLogicalTableTextExportExtensions.ExtractMarkdownTables(
            pdf,
            new PdfLogicalTableTextExportOptions {
                LayoutOptions = new PdfTextLayoutOptions {
                    ForceSingleColumn = true
                },
                PageRanges = new[] { PdfPageRange.From(2, 2) },
                EmptyTableMessage = "Nothing tabular was detected."
            });

        Assert.Equal("Nothing tabular was detected.", empty);
    }

    [Fact]
    public void PdfTables_ExtractHtmlTables_ExportsSemanticTablesAndEscapesCells() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Code", "Label", "Amount" },
                new[] { "A&B", "<Alpha>", "125.50" },
                new[] { "B|C", "Beta", "14" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();

        string html = PdfLogicalTableTextExportExtensions.ExtractHtmlTables(
            pdf,
            new PdfLogicalTableTextExportOptions {
                LayoutOptions = new PdfTextLayoutOptions {
                    ForceSingleColumn = true
                },
                HtmlDocumentTitle = "Detected PDF Tables"
            });

        Assert.StartsWith("<!doctype html>", html, StringComparison.Ordinal);
        Assert.Contains("<title>Detected PDF Tables</title>", html, StringComparison.Ordinal);
        Assert.Contains("<figure class=\"pdf-table\" data-page-number=\"1\" data-table-index=\"0\"", html, StringComparison.Ordinal);
        Assert.Contains("<figcaption>PDF page 1, table 1</figcaption>", html, StringComparison.Ordinal);
        Assert.Contains("<th class=\"pdf-numeric\" style=\"text-align:right\">Amount</th>", html, StringComparison.Ordinal);
        Assert.Contains("<td>A&amp;B</td>", html, StringComparison.Ordinal);
        Assert.Contains("<td>&lt;Alpha&gt;</td>", html, StringComparison.Ordinal);
        Assert.Contains("<td>B|C</td>", html, StringComparison.Ordinal);
        Assert.Contains("<td class=\"pdf-numeric\" style=\"text-align:right\">125.50</td>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfTables_ToMarkdownTables_EscapesMarkdownTableSyntax() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 260,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Key", "Value" },
                new[] { "A|B", "# literal" },
                new[] { "2)", "<tag>" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 110, 170 },
                HeaderRowCount = 1
            })
            .ToBytes();

        string markdown = PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
                ForceSingleColumn = true
            })
            .ToMarkdownTables(new PdfLogicalTableTextExportOptions {
                IncludeSourceCaptions = false
            });

        Assert.Contains("| A\\|B | \\# literal |", markdown, StringComparison.Ordinal);
        Assert.Contains("| 2\\) | \\<tag\\> |", markdown, StringComparison.Ordinal);
    }
}
