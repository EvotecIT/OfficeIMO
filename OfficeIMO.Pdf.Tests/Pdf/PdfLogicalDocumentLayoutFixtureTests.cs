using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfLogicalDocumentTests {
    [Fact]
    public void Load_ReadsTwoPageStatementFixtureAsLogicalPagesAndTables() {
        byte[] pdf = PdfDocumentRasterVisualBaselineTests.CreateLineItemsTwoPage();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf);

        Assert.Equal(2, logical.PageCount);
        Assert.True(logical.HasSourcePage(1));
        Assert.True(logical.HasSourcePage(2));
        Assert.Equal(new[] { 1, 2 }, logical.Pages.Select(page => page.PageNumber).ToArray());
        Assert.Contains(logical.TextBlocks, block => Normalize(block.Text).Contains("Statement#4048", StringComparison.Ordinal));
        Assert.Contains(logical.Pages[0].TextBlocks, block => Normalize(block.Text).Contains("Experientiamnostrum", StringComparison.Ordinal));
        Assert.Contains(logical.Pages[1].TextBlocks, block => Normalize(block.Text).Contains("Subtotal", StringComparison.Ordinal));
        Assert.Contains(logical.Pages[1].TextBlocks, block => Normalize(block.Text).Contains("Documentnote:", StringComparison.Ordinal));

        Assert.True(logical.HasElementKind(PdfLogicalElementKind.Table));
        Assert.True(logical.Pages[0].HasElementKind(PdfLogicalElementKind.Table));
        Assert.True(logical.Pages[1].HasElementKind(PdfLogicalElementKind.Table));
        Assert.Contains(logical.Tables, table =>
            table.PageNumber == 1 &&
            table.Columns.Count >= 4 &&
            table.Rows.Any(row => RowContains(row, "Experientiamnostrum", "31,80PLN", "2", "63,60PLN")));
        Assert.Contains(logical.Tables, table =>
            table.PageNumber == 2 &&
            table.Columns.Count >= 2 &&
            table.Rows.Any(row => RowContains(row, "Subtotal", "5201,32PLN")) &&
            table.Rows.Any(row => RowContains(row, "Total", "6397,62PLN")));

        PdfLogicalDocument selected = PdfLogicalDocument.LoadPageRanges(pdf, PdfPageRange.ParseMany("2,1"));

        Assert.Equal(new[] { 2, 1 }, selected.Pages.Select(page => page.PageNumber).ToArray());
        Assert.Contains(selected.Pages[0].TextBlocks, block => Normalize(block.Text).Contains("Subtotal", StringComparison.Ordinal));
        Assert.Contains(selected.Pages[1].TextBlocks, block => Normalize(block.Text).Contains("Statement#4048", StringComparison.Ordinal));
    }

    [Fact]
    public void Load_GroupsWrappedTextLinesIntoParagraphs() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("This logical paragraph should wrap across multiple nearby PDF text lines so wrappers can start from paragraph-like objects."))
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "P-100", "Paragraph table text", "2" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 50, 100, 30 },
                HeaderRowCount = 1
            })
            .ToBytes();

        PdfLogicalPage page = Assert.Single(PdfLogicalDocument.Load(pdf, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        }).Pages);

        PdfLogicalParagraph paragraph = Assert.Single(page.Paragraphs, item => item.Text.Contains("logical paragraph", StringComparison.Ordinal));
        Assert.True(paragraph.Lines.Count > 1);
        Assert.Contains("logical paragraph", paragraph.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("P-100", paragraph.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void Load_ReadsCompleteSeekableStreamAndRestoresPosition() {
        byte[] source = PdfDocument.Create()
            .Paragraph(p => p.Text("Logical stream marker."))
            .ToBytes();
        using var stream = new MemoryStream(source);
        stream.Position = source.Length / 2;
        long originalPosition = stream.Position;

        PdfLogicalDocument logical = PdfLogicalDocument.Load(stream);

        Assert.Single(logical.Pages);
        Assert.Contains(logical.Pages[0].TextBlocks, block => block.Text.Contains("Logical stream marker", StringComparison.Ordinal));
        Assert.Equal(originalPosition, stream.Position);
    }
}
