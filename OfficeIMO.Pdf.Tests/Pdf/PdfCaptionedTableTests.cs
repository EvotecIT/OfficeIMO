using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCaptionedTableTests {
    [Fact]
    public void RepeatedHeadersKeepCaptionedTablesSeparateAndPreserveWrappedCells() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "Website", "static", "downloads", "showcase",
            "features", "pdf-table-styles", "pdf.tablestylegallery.pdf");
        using var pdf = PdfDocument.Load(path);
        PdfDocumentReadResult result = pdf.Read();

        // The gallery has five independently captioned tables on each of its
        // first six pages and one on its last page. Each style uses the same rows.
        Assert.Equal(new[] { 5, 5, 5, 5, 5, 5, 1 }, result.Pages.Select(page => page.Tables.Count));
        Assert.All(result.Tables, table => {
            Assert.Equal(3, table.Rows.Count);
            Assert.Equal(new[] { "Signal", "State", "Notes" }, table.Rows[0]);
            Assert.Equal(new[] { "Header", "Repeated", "The first row should stay readable without relying on a domain-specific preset." },
                table.Rows[1]);
            Assert.Equal(new[] { "Flow", "Generic", "Borders, row separators, and spacing should reveal the preset shape at raster level." },
                table.Rows[2]);
        });

        string[] captions = {
            "TableNormal", "TableGrid", "TableGridLight", "PlainTable1", "GridTable1Light",
            "GridTable1LightAccent1", "GridTable1LightAccent2", "GridTable1LightAccent3", "GridTable1LightAccent4", "GridTable1LightAccent5",
            "GridTable1LightAccent6", "ListTable1Light", "ListTable1LightAccent1", "ListTable1LightAccent2", "ListTable1LightAccent3",
            "ListTable1LightAccent4", "ListTable1LightAccent5", "ListTable1LightAccent6", "GridTableLight", "GridTable1Light-Accent1",
            "GridTable1Light-Accent2", "GridTable1Light-Accent3", "GridTable1Light-Accent4", "GridTable1Light-Accent5", "GridTable1Light-Accent6",
            "ListTable1Light-Accent1", "ListTable1Light-Accent2", "ListTable1Light-Accent3", "ListTable1Light-Accent4", "ListTable1Light-Accent5",
            "ListTable1Light-Accent6"
        };
        var captionSet = new HashSet<string>(captions, StringComparer.Ordinal);
        string[] textLines = result.Text.Split('\n').Select(line => line.Trim()).ToArray();
        Assert.Equal(captions, textLines.Where(captionSet.Contains));
        foreach (string caption in captions) {
            Assert.Equal("Signal\tState\tNotes", textLines[Array.IndexOf(textLines, caption) + 1]);
        }
        Assert.Equal(captions, result.ExtractMarkdown().Split('\n')
            .Select(line => line.Trim().TrimStart('#', ' ')).Where(captionSet.Contains));

        IReadOnlyList<PdfLogicalTableContinuationGroup> groups = result.GetTableContinuationGroups();
        Assert.Equal(31, groups.Count);
        Assert.All(groups, group => Assert.Single(group.Segments));
    }

    [Theory]
    [InlineData("Quality register")]
    [InlineData("Page {page} of {pages}")]
    public void RunningHeadersKeepOneWrappedTableAcrossPages(string headerFormat) {
        var options = new PdfCreateOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = OfficeLength.Points(9),
            PageSetup = {
                Size = new PdfPageSize(OfficeLength.Points(380), OfficeLength.Points(215)),
                Margins = { Left = OfficeLength.Points(30), Right = OfficeLength.Points(30), Top = OfficeLength.Points(34), Bottom = OfficeLength.Points(10) },
                Header = {
                    Format = headerFormat,
                    Show = true,
                    Alignment = OfficeAlignment.Left,
                    StandardFont = PdfStandardFont.Helvetica,
                    FontSize = OfficeLength.Points(8),
                    Offset = OfficeLength.Points(12)
                }
            }
        };
        var rows = new List<string[]> { new[] { "No.", "State", "Result" } };
        foreach (string number in new[] { "1", "2", "3", "4" }) {
            rows.Add(new[] { number, "Active", "Check completed after coordinated review and inventory validation by the maintenance team." });
        }
        var style = new PdfTableStyle {
            BorderWidth = OfficeLength.Points(0),
            HeaderBold = true,
            HeaderRowCount = 1,
            RepeatHeaderRowCount = 1,
            AllowRowBreakAcrossPages = false,
            CellPaddingX = OfficeLength.Points(5),
            CellPaddingY = OfficeLength.Points(3),
            ColumnWidths = new List<OfficeLength?> { OfficeLength.Points(70), OfficeLength.Points(90), OfficeLength.Points(120) }
        };
        using var generated = PdfDocument.Create(options).Table(rows, style: style);
        using var pdf = PdfDocument.Load(generated.ToBytes());
        PdfDocumentReadResult result = pdf.Read();

        Assert.Equal(2, result.Pages.Count);
        PdfLogicalTableContinuationGroup group = Assert.Single(result.GetTableContinuationGroups());
        Assert.Equal(2, group.Segments.Count);
        Assert.Equal(new[] { "1", "2", "3", "4" }, group.Rows.Select(row => row[0]));
    }
}
