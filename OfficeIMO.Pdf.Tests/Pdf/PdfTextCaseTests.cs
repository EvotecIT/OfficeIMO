using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using System.Text;
using System.Reflection;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public class PdfTextCaseTests {
    [Fact]
    public void PreTypographyTabAlignedConstructorRemainsBinaryDiscoverable() {
        ConstructorInfo? constructor = typeof(PdfTextRun).GetConstructor(new[] {
            typeof(string), typeof(bool), typeof(bool), typeof(PdfColor?), typeof(bool), typeof(bool),
            typeof(double?), typeof(PdfStandardFont?), typeof(string), typeof(string), typeof(PdfTextBaseline),
            typeof(string), typeof(PdfTabLeaderStyle), typeof(PdfTabAlignment), typeof(PdfColor?), typeof(string)
        });

        Assert.NotNull(constructor);
    }

    [Fact]
    public void PreTypographyLinkFactoriesRemainBinaryDiscoverable() {
        Type[] parameters = {
            typeof(string), typeof(string), typeof(PdfColor?), typeof(bool), typeof(string),
            typeof(PdfTextBaseline), typeof(double?), typeof(PdfColor?), typeof(PdfStandardFont?), typeof(string)
        };

        Assert.NotNull(typeof(PdfTextRun).GetMethod(nameof(PdfTextRun.Link), parameters));
        Assert.NotNull(typeof(PdfTextRun).GetMethod(nameof(PdfTextRun.LinkToBookmark), parameters));
    }

    [Fact]
    public void WithTextCasePreservesImmutableRunFormatting() {
        PdfTextRun source = new("Styled", bold: true,
            color: PdfColor.FromRgb(51, 102, 153), italic: true,
            fontSize: 14, baseline: PdfTextBaseline.Superscript,
            backgroundColor: PdfColor.FromRgb(240, 240, 240), fontFamily: "Aptos",
            underlineStyle: OfficeTextDecorationStyle.Dashed,
            strikeStyle: OfficeTextDecorationStyle.Double,
            decorationColor: PdfColor.FromRgb(200, 10, 20));

        PdfTextRun actual = source.WithTextCase(OfficeTextCase.ToggleCase);

        Assert.Equal("sTYLED", actual.Text);
        Assert.True(actual.Bold);
        Assert.True(actual.Italic);
        Assert.True(actual.Underline);
        Assert.True(actual.Strike);
        Assert.Equal(PdfTextBaseline.Superscript, actual.Baseline);
        Assert.Equal(14D, actual.FontSize);
        Assert.Equal("Aptos", actual.FontFamily);
        Assert.Equal(source.Color, actual.Color);
        Assert.Equal(source.BackgroundColor, actual.BackgroundColor);
        Assert.Equal(OfficeTextDecorationStyle.Dashed, actual.UnderlineStyle);
        Assert.Equal(OfficeTextDecorationStyle.Double, actual.StrikeStyle);
        Assert.Equal(source.DecorationColor, actual.DecorationColor);
    }

    [Fact]
    public void PdfWriterEmitsNativeDecorationPatterns() {
        PdfTextRun run = new(
            "Decorated",
            color: PdfColor.FromRgb(10, 20, 30),
            underlineStyle: OfficeTextDecorationStyle.Dashed,
            strikeStyle: OfficeTextDecorationStyle.Double,
            decorationColor: PdfColor.FromRgb(255, 0, 0));

        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Header(header => header.Text(text => text.Run(run)))
            .Paragraph(paragraph => paragraph
                .Underline(OfficeTextDecorationStyle.Dashed)
                .Strike(OfficeTextDecorationStyle.Double)
                .Text("Body"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("] 0 d", raw, System.StringComparison.Ordinal);
        Assert.Contains("1 0 0 RG", raw, System.StringComparison.Ordinal);
        Assert.True(raw.Split(new[] { " RG" }, System.StringSplitOptions.None).Length >= 4,
            "Expected one dashed underline plus two lines for the double strikethrough.");
    }

    [Fact]
    public void PdfTextRejectsUndefinedDecorationStyles() {
        Assert.Throws<System.ArgumentOutOfRangeException>(() => new PdfTextRun(
            "Invalid",
            underlineStyle: (OfficeTextDecorationStyle)99));
        Assert.Throws<System.ArgumentOutOfRangeException>(() => new PdfTextRun(
            "Invalid",
            strikeStyle: (OfficeTextDecorationStyle)99));
    }

    [Fact]
    public void PdfWriterRendersSharedDrawingRichTextWithStylesAndFrameLayout() {
        var drawing = new OfficeDrawing(220D, 90D)
            .AddRichText(
                new[] {
                    new OfficeRichTextRun(
                        "Styled ", 16D, OfficeColor.DarkBlue,
                        bold: true, italic: true,
                        underlineStyle: OfficeTextDecorationStyle.Dashed),
                    new OfficeRichTextRun(
                        "H2O", 16D, OfficeColor.DarkRed,
                        strikethroughStyle: OfficeTextDecorationStyle.Double,
                        baseline: OfficeTextBaseline.Subscript)
                },
                10D, 10D, 200D, 70D,
                alignment: OfficeTextAlignment.Center,
                verticalAlignment: OfficeTextVerticalAlignment.Center,
                rotationDegrees: 4D,
                wrapText: true,
                shrinkToFit: true,
                padding: new OfficeTextPadding(4D, 4D, 4D, 4D),
                paragraphIndent: OfficeTextParagraphIndent.FirstLine(6D));

        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Drawing(drawing)
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains(" cm", raw, System.StringComparison.Ordinal);
        Assert.Contains("] 0 d", raw, System.StringComparison.Ordinal);
        Assert.Contains(" Ts", raw, System.StringComparison.Ordinal);
        Assert.True(raw.Split(new[] { " RG" }, System.StringSplitOptions.None).Length >= 4,
            "Expected a dashed underline and both lines of a double strikethrough.");
    }
}
