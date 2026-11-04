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
            var style = new PdfParagraphStyle { FontSize = global::OfficeIMO.OfficeLength.Points(size) };
            if (defaultStyle) options.DefaultParagraphStyle = style;
            var document = PdfDocument.Create(options);
            document.Content.Row(row => row.Gap(global::OfficeIMO.OfficeLength.Points(10))
                .AutoColumn(column => column.AddParagraph(p => {
                    if (explicitRunSize) p.RunDefaults.FontSize = global::OfficeIMO.OfficeLength.Points(size);
                    p.AddText("WWWW");
                }, style: defaultStyle ? null : style))
                .RelativeColumn(column => column.AddParagraph("Marker")));
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
        var options = Options();
        if (atBoundary) options.PageSetup.Size = new PdfPageSize(options.PageSetup.Size.Width, global::OfficeIMO.OfficeLength.Points(60D + 5D * (multiple ? 8D : 20D)));
        var document = PdfDocument.Create(options);
        document.Content.Columns(columns => columns.AddParagraph(p => {
            p.AddText(atBoundary ? "A\nB\nC\nD\n" : "A\n").SetFontSize(global::OfficeIMO.OfficeLength.Points(8));
            p.AddLineBreak().SetFontSize(global::OfficeIMO.OfficeLength.Points(32));
            p.AddText(atBoundary ? "E" : "B\nC\nD\nE\nF\nG\nH\nI\nJ").SetFontSize(global::OfficeIMO.OfficeLength.Points(8));
        }, style: new PdfParagraphStyle {
            FontSize = global::OfficeIMO.OfficeLength.Points(8), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0),
            LineSpacing = multiple ? PdfLineSpacing.Multiple(1) : PdfLineSpacing.AtLeast(global::OfficeIMO.OfficeLength.Points(20), 1)
        }), new PdfMultiColumnOptions {
            ColumnCount = 2, Gap = global::OfficeIMO.OfficeLength.Points(12), BalanceParagraphLines = true, BalanceLastPage = !atBoundary,
            ColumnDefinitions = atBoundary ? new[] {
                new PdfFlowColumn(PdfColumnWidth.Fixed(global::OfficeIMO.OfficeLength.Points(148))), new PdfFlowColumn(PdfColumnWidth.Fixed(global::OfficeIMO.OfficeLength.Points(180)))
            } : Array.Empty<PdfFlowColumn>()
        });
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(atBoundary ? "ABCDE" : "ABCDEFGHIJ", string.Concat(letters.Select(letter => letter.Value)));
        var first = Assert.Single(letters, letter => letter.Value == "A");
        var other = Assert.Single(letters, letter => letter.Value == (atBoundary ? "E" : "B"));
        if (atBoundary) Assert.True(other.StartBaseLine.X > first.StartBaseLine.X + 100);
        else Assert.Equal(first.StartBaseLine.X, other.StartBaseLine.X, 3);
        Assert.Equal(expected, first.StartBaseLine.Y - other.StartBaseLine.Y, 3);
    }

    private static PdfCreateOptions Options() => new() {
        PageSetup = {
            Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(400), global::OfficeIMO.OfficeLength.Points(400)),
            Margins = { Top = global::OfficeIMO.OfficeLength.Points(30), Bottom = global::OfficeIMO.OfficeLength.Points(30), Left = global::OfficeIMO.OfficeLength.Points(30), Right = global::OfficeIMO.OfficeLength.Points(30) }
        }, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12)
    };
}
