using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRichLineBaselineTests {
    [Theory]
    [InlineData("flow", false)]
    [InlineData("column", false)]
    [InlineData("table", false)]
    [InlineData("flow", true)]
    [InlineData("column", true)]
    [InlineData("table", true)]
    [InlineData("canvas", false)]
    [InlineData("canvas", true)]
    [InlineData("linked-table", false)]
    [InlineData("linked-table", true)]
    public void FirstBaselineUsesTheVisibleRunFontMetrics(string frame, bool namedFont) {
        using var pdf = PdfPigDocument.Open(Render(frame, namedFont, mixed: false));
        var letter = Assert.Single(pdf.GetPage(1).Letters);
        double ascent = namedFont ? .8D : .74D;
        Assert.Equal(270D - 24D * ascent, letter.StartBaseLine.Y, 3);
        Assert.Equal(24D, letter.FontSize, 3);
    }

    [Theory]
    [InlineData("flow", false)]
    [InlineData("column", false)]
    [InlineData("table", false)]
    [InlineData("flow", true)]
    [InlineData("column", true)]
    [InlineData("table", true)]
    [InlineData("canvas", false)]
    [InlineData("canvas", true)]
    [InlineData("linked-table", false)]
    [InlineData("linked-table", true)]
    public void MixedSizeLinesUseTheirOwnAscentWithinEachLineBox(string frame, bool namedFont) {
        using var pdf = PdfPigDocument.Open(Render(frame, namedFont, mixed: true));
        var letters = pdf.GetPage(1).Letters;
        var first = Assert.Single(letters, letter => letter.Value == "A");
        var middle = Assert.Single(letters, letter => letter.Value == "B");
        var last = Assert.Single(letters, letter => letter.Value == "C");
        double ascent = namedFont ? .8D : .74D;
        Assert.Equal(270D - 8D * ascent, first.StartBaseLine.Y, 3);
        Assert.Equal(8D * 1.4D + 16D * ascent, first.StartBaseLine.Y - middle.StartBaseLine.Y, 3);
        Assert.Equal(24D * 1.4D - 16D * ascent, middle.StartBaseLine.Y - last.StartBaseLine.Y, 3);
    }

    private static byte[] Render(string frame, bool namedFont, bool mixed) {
        const string family = "Line Box Test";
        var options = new PdfOptions { PageWidth = 300, PageHeight = 300,
            MarginTop = 30, MarginBottom = 30, MarginLeft = 30, MarginRight = 30, DefaultFontSize = 8 };
        if (namedFont) options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(family,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A', 'B', 'C')));
        var style = new PdfParagraphStyle { FontSize = 8, SpacingBefore = 0, SpacingAfter = 0,
            LineSpacing = PdfLineSpacing.Multiple(1.4) };
        Action<PdfParagraphBuilder> text = p => {
            if (namedFont) p.FontFamily(family);
            if (mixed) p.FontSize(8).Text("A\n").FontSize(24).Text("B\n").FontSize(8).Text("C");
            else p.FontSize(24).Text("A");
        };
        var document = PdfDocument.Create(options);
        if (frame == "flow") document.Paragraph(text, style: style);
        else if (frame == "column") document.Content.Row(row => row.RelativeColumn(column => column.Paragraph(text, style: style)));
        else if (frame == "canvas") {
            var runs = mixed ? new[] { Run("A\n", 8), Run("B\n", 24), Run("C", 8) } : new[] { Run("A", 24) };
            document.Canvas(canvas => canvas.TextBox(runs, 30, 30, 240, 240, new PdfCanvasTextBoxStyle {
                Font = PdfStandardFont.TimesRoman, FontSize = 8, LineHeight = 8 * 1.4,
                PaddingX = 0, PaddingY = 0, Background = null, BorderColor = null
            }));
        }
        else {
            var runs = mixed
                ? new[] { Run("A\n", 8), Run("B\n", 24), Run("C", 8) }
                : new[] { Run("A", 24) };
            var cell = new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs, fontSize: 8, lineSpacing: style.LineSpacing) },
                linkUri: frame == "linked-table" ? "https://example.com/cell" : null);
            var tableStyle = TableStyles.Minimal();
            tableStyle.HeaderRowCount = 0; tableStyle.FontSize = 8;
            tableStyle.CellPaddingX = 0; tableStyle.CellPaddingY = 0; tableStyle.SpacingBefore = 0;
            document.Table(new[] { new[] { cell } }, style: tableStyle);
        }
        return document.ToBytes();

        PdfTextRun Run(string value, double size) => new(value, fontSize: size, font: PdfStandardFont.Helvetica,
            fontFamily: namedFont ? family : null, linkUri: frame == "linked-table" ? "https://example.com/run" : null);
    }
}
