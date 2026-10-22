using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDecimalRichFieldTests {
    [Theory]
    [InlineData(PdfTextWhitespaceMode.Collapse, '.', "1,234.56", true)]
    [InlineData(PdfTextWhitespaceMode.Preserve, ',', "1.234,56", true)]
    [InlineData(PdfTextWhitespaceMode.Preserve, '.', "1234", false)]
    public void DecimalTabsAnchorCompleteMixedFontFieldsAtTheSelectedSeparatorOrFieldEnd(
        PdfTextWhitespaceMode whitespace, char separator, string number, bool hasSeparator) {
        var style = new PdfParagraphStyle { TextWhitespaceMode = whitespace, SpacingAfter = 0D };
        style.AddTabStop(200D, PdfTabAlignment.DecimalSeparator).DecimalSeparator = separator;
        var options = new PdfOptions {
            PageWidth = 420D, PageHeight = 180D, MarginLeft = 20D, MarginRight = 20D,
            MarginTop = 20D, MarginBottom = 20D, DefaultFontSize = 12D
        };
        byte[] bytes = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Runs(new[] {
            PdfTextRun.Normal("LABEL\t"),
            PdfTextRun.Normal("USD ", fontSize: 9D, font: PdfStandardFont.HelveticaBold),
            PdfTextRun.Normal(number, fontSize: 16D, font: PdfStandardFont.Courier)
                .WithHorizontalTextScaling(125D).WithCharacterSpacing(.5D),
            PdfTextRun.Normal(" net", fontSize: 8D, font: PdfStandardFont.TimesRoman)
        }), style: style).ToBytes();

        using var pdf = PdfPigDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        var visible = page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("LABELUSD" + number + "net", string.Concat(visible.Select(letter => letter.Value)));
        Assert.All(visible, letter => Assert.Equal(visible[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3));
        double anchor = hasSeparator
            ? Assert.Single(visible, letter => letter.Value == separator.ToString()).StartBaseLine.X
            : visible[visible.Length - 1].EndBaseLine.X;
        Assert.Equal(220D, anchor, 2);
    }

    [Theory]
    [InlineData(PdfTabAlignment.Right)]
    [InlineData(PdfTabAlignment.Center)]
    [InlineData(PdfTabAlignment.DecimalSeparator)]
    public void TabFieldsRemeasureWhenWrappingSelectsAnEarlierNonLeftStop(PdfTabAlignment alignment) {
        var style = new PdfParagraphStyle { SpacingAfter = 0D };
        style.AddTabStop(90D, alignment).DecimalSeparator = '.';
        style.AddTabStop(170D, PdfTabAlignment.Left);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
            PageWidth = 220D, PageHeight = 180D, MarginLeft = 20D, MarginRight = 20D,
            MarginTop = 20D, MarginBottom = 20D, DefaultFont = PdfStandardFont.Courier, DefaultFontSize = 10D
        }).Paragraph(paragraph => paragraph.Text("PREFIXPREFIXPREFIX\tUSD 12.50 net"), style: style).ToBytes();

        using var pdf = PdfPigDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        var visible = page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("PREFIXPREFIXPREFIXUSD12.50net", string.Concat(visible.Select(letter => letter.Value)));
        var prefix = visible.Take("PREFIXPREFIXPREFIX".Length).ToArray();
        var field = visible.Skip(prefix.Length).ToArray();
        Assert.All(prefix, letter => Assert.Equal(prefix[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3));
        Assert.True(field[0].StartBaseLine.Y < prefix[0].StartBaseLine.Y);
        Assert.All(field, letter => Assert.Equal(field[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3));
        double anchor = alignment switch {
            PdfTabAlignment.Center => (field[0].StartBaseLine.X + field[field.Length - 1].EndBaseLine.X) / 2D,
            PdfTabAlignment.DecimalSeparator => Assert.Single(field, letter => letter.Value == ".").StartBaseLine.X,
            _ => field[field.Length - 1].EndBaseLine.X
        };
        // Courier10 makes the prefix108pt and first field word18pt. The170pt
        // left stop cannot admit that word in180pt, so the new line selects90pt.
        Assert.Equal(110D, anchor, 2);
    }
}
