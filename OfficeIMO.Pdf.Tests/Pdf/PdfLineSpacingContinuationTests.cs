using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfLineSpacingContinuationTests {
    [Theory]
    [InlineData(8D, false)]
    [InlineData(24D, false)]
    [InlineData(8D, true)]
    [InlineData(24D, true)]
    public void AutomaticColumnWidthUsesEffectiveParagraphFontSize(double size, bool defaultStyle) {
        double Render(bool explicitRunSize) {
            var options = Options();
            var style = new PdfParagraphStyle { FontSize = size };
            if (defaultStyle) options.DefaultParagraphStyle = style;
            var document = PdfDocument.Create(options);
            document.Content.Row(row => row.Gap(10)
                .AutoColumn(column => column.Paragraph(p => {
                    if (explicitRunSize) p.FontSize(size);
                    p.Text("WWWW");
                }, style: defaultStyle ? null : style))
                .RelativeColumn(column => column.Text("Marker")));
            using var pdf = PdfPigDocument.Open(document.ToBytes());
            return Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "Marker").BoundingBox.Left;
        }
        Assert.Equal(Render(true), Render(false), 3);
    }

    [Theory]
    [InlineData(false, false, 52D)]
    [InlineData(true, false, 40D)]
    [InlineData(false, true, 32D)]
    [InlineData(true, true, 32D)]
    public void BalancedColumnsRetainStyledBlankLines(bool multiple, bool atBoundary, double expected) {
        var document = PdfDocument.Create(Options());
        document.Content.Columns(columns => columns.Paragraph(p => {
            p.FontSize(8).Text(atBoundary ? "A\nB\nC\nD\n" : "A\n");
            p.FontSize(32).LineBreak();
            p.FontSize(8).Text(atBoundary ? "E\nF\nG\nH\nI" : "B\nC\nD\nE\nF\nG\nH\nI\nJ");
        }, style: new PdfParagraphStyle {
            FontSize = 8, SpacingBefore = 0, SpacingAfter = 0,
            LineSpacing = multiple ? PdfLineSpacing.Multiple(1) : PdfLineSpacing.AtLeast(20, 1)
        }), new PdfMultiColumnOptions { ColumnCount = 2, Gap = 12, BalanceParagraphLines = true, BalanceLastPage = true });
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var letters = pdf.GetPage(1).Letters;
        var first = Assert.Single(letters, letter => letter.Value == "A");
        var other = Assert.Single(letters, letter => letter.Value == (atBoundary ? "E" : "B"));
        if (atBoundary) Assert.True(other.StartBaseLine.X > first.StartBaseLine.X + 100);
        else Assert.Equal(first.StartBaseLine.X, other.StartBaseLine.X, 3);
        Assert.Equal(expected, first.StartBaseLine.Y - other.StartBaseLine.Y, 3);
    }

    private static PdfOptions Options() => new() {
        PageWidth = 400, PageHeight = 400, DefaultFontSize = 12,
        MarginTop = 30, MarginBottom = 30, MarginLeft = 30, MarginRight = 30
    };
}
