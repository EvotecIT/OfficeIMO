using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using System.Text;
using System.Reflection;
using System.Globalization;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public class PdfTextCaseTests {
    [Theory]
    [InlineData(true, true, 3)]
    [InlineData(true, false, 2)]
    [InlineData(false, true, 1)]
    public void RichParagraphUnderlinesSpacesUsingTheirSourceRun(bool firstUnderlined, bool secondUnderlined, int expectedLineCount) {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Paragraph(paragraph => {
                if (firstUnderlined) paragraph.Underlined("Elite ");
                else paragraph.Text("Elite ");
                if (secondUnderlined) paragraph.Underlined("Performance");
                else paragraph.Text("Performance");
            })
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        MatchCollection strokes = Regex.Matches(raw, @"(?<start>\d+(?:\.\d+)?) (?<y>\d+(?:\.\d+)?) m (?<end>\d+(?:\.\d+)?) \k<y> l S");
        Assert.Equal(expectedLineCount, strokes.Count);
        if (firstUnderlined) {
            double firstEnd = double.Parse(strokes[0].Groups["end"].Value, CultureInfo.InvariantCulture);
            double spaceStart = double.Parse(strokes[1].Groups["start"].Value, CultureInfo.InvariantCulture);
            Assert.InRange(Math.Abs(firstEnd - spaceStart), 0, 0.01);
            if (secondUnderlined) {
                double spaceEnd = double.Parse(strokes[1].Groups["end"].Value, CultureInfo.InvariantCulture);
                double secondStart = double.Parse(strokes[2].Groups["start"].Value, CultureInfo.InvariantCulture);
                Assert.InRange(Math.Abs(spaceEnd - secondStart), 0, 0.01);
            }
        }
    }

    [Fact]
    public void RichParagraphWordsUnderlineLeavesSpacesClear() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Paragraph(paragraph => paragraph.Underlined("Elite Performance", OfficeTextDecorationStyle.Words))
            .ToBytes();
        MatchCollection strokes = Regex.Matches(Encoding.ASCII.GetString(bytes), @"(?<start>\d+(?:\.\d+)?) (?<y>\d+(?:\.\d+)?) m (?<end>\d+(?:\.\d+)?) \k<y> l S");
        Assert.Equal(2, strokes.Count);
        double firstEnd = double.Parse(strokes[0].Groups["end"].Value, CultureInfo.InvariantCulture);
        double secondStart = double.Parse(strokes[1].Groups["start"].Value, CultureInfo.InvariantCulture);
        Assert.True(secondStart - firstEnd > 1);
    }

    [Theory]
    [InlineData(PdfTextBaseline.Superscript)]
    [InlineData(PdfTextBaseline.Subscript)]
    public void UnderlinedScriptSpaceUsesTheSourceRunBaseline(PdfTextBaseline baseline) {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Paragraph(paragraph => paragraph.Runs(new[] {
                new PdfTextRun("Elite ", underlineStyle: OfficeTextDecorationStyle.Single, baseline: baseline, fontSize: 18),
                new PdfTextRun("Performance", fontSize: 18)
            }))
            .ToBytes();

        MatchCollection strokes = Regex.Matches(Encoding.ASCII.GetString(bytes), @"(?<start>\d+(?:\.\d+)?) (?<y>\d+(?:\.\d+)?) m (?<end>\d+(?:\.\d+)?) \k<y> l S");
        Assert.Equal(2, strokes.Count);
        double wordY = double.Parse(strokes[0].Groups["y"].Value, CultureInfo.InvariantCulture);
        double gapY = double.Parse(strokes[1].Groups["y"].Value, CultureInfo.InvariantCulture);
        Assert.InRange(Math.Abs(wordY - gapY), 0, 0.02);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DrawingWordsUnderlineLeavesSpacesClear(bool wrapText) {
        var drawing = new OfficeDrawing(240, 45).AddStyledText(
            "Elite Performance", 5, 4, 220, 34, new OfficeFontInfo("Arial", 15), OfficeColor.Black,
            OfficeTextAlignment.Left, null, OfficeTextVerticalAlignment.Top, 0, null, null,
            wrapText, false, false, false, false, null, null,
            OfficeTextDecorationStyle.Words, OfficeTextDecorationStyle.None, OfficeTextBaseline.Normal);
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Canvas(canvas => canvas.Drawing(drawing, 10, 10, 240, 45))
            .ToBytes();

        MatchCollection strokes = Regex.Matches(Encoding.ASCII.GetString(bytes), @"(?<start>\d+(?:\.\d+)?) (?<y>\d+(?:\.\d+)?) m (?<end>\d+(?:\.\d+)?) \k<y> l S");
        Assert.Equal(2, strokes.Count);
        double firstEnd = double.Parse(strokes[0].Groups["end"].Value, CultureInfo.InvariantCulture);
        double secondStart = double.Parse(strokes[1].Groups["start"].Value, CultureInfo.InvariantCulture);
        Assert.True(secondStart - firstEnd > 1);
    }

    [Fact]
    public void ShapedRightToLeftWordsUnderlineFollowsVisualWordOrder() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) return;
        const string text = "שלום ים";
        byte[] fontData = File.ReadAllBytes(fontPath);
        PdfTrueTypeFontProgram font = PdfTrueTypeFontProgram.Parse(fontData, "RTL underline test");
        if (PdfTextDiagnostics.AnalyzeEmbeddedFontText(text, font).Count > 0) return;

        var drawing = new OfficeDrawing(240, 45).AddStyledText(
            text, 5, 4, 220, 34, new OfficeFontInfo("Arial", 18), OfficeColor.Black,
            OfficeTextAlignment.Left, null, OfficeTextVerticalAlignment.Top, 0, null, null,
            false, false, false, false, false, null, null,
            OfficeTextDecorationStyle.Words, OfficeTextDecorationStyle.None, OfficeTextBaseline.Normal);
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false }
                .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "RTL underline test")
                .SetTextShapingProvider(OfficeManagedTextShapingProvider.Instance))
            .Canvas(canvas => canvas.Drawing(drawing, 10, 10, 240, 45))
            .ToBytes();

        MatchCollection strokes = Regex.Matches(Encoding.ASCII.GetString(bytes), @"(?<start>\d+(?:\.\d+)?) (?<y>\d+(?:\.\d+)?) m (?<end>\d+(?:\.\d+)?) \k<y> l S");
        Assert.Equal(2, strokes.Count);
        double firstStart = double.Parse(strokes[0].Groups["start"].Value, CultureInfo.InvariantCulture);
        double firstEnd = double.Parse(strokes[0].Groups["end"].Value, CultureInfo.InvariantCulture);
        double secondStart = double.Parse(strokes[1].Groups["start"].Value, CultureInfo.InvariantCulture);
        double secondEnd = double.Parse(strokes[1].Groups["end"].Value, CultureInfo.InvariantCulture);
        Assert.True(firstEnd - firstStart < secondEnd - secondStart);
        Assert.True(secondStart - firstEnd > 1);
    }

    [Fact]
    public void HeaderWordsUnderlineLeavesSpacesClear() {
        byte[] bytes = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Header(header => header.Text(text => text.Run(new PdfTextRun("Elite Performance", underlineStyle: OfficeTextDecorationStyle.Words))))
            .Paragraph(paragraph => paragraph.Text("Body"))
            .ToBytes();
        MatchCollection strokes = Regex.Matches(Encoding.ASCII.GetString(bytes), @"(?<start>\d+(?:\.\d+)?) (?<y>\d+(?:\.\d+)?) m (?<end>\d+(?:\.\d+)?) \k<y> l S");
        Assert.Equal(2, strokes.Count);
    }

    [Fact]
    public void HeaderNewlinesStopAtTheLayoutLineLimit() {
        PdfDocument document = PdfDocument.Create(new PdfOptions {
                ShowHeader = true,
                HeaderFormat = new string('\n', 100_001)
            })
            .Paragraph(paragraph => paragraph.Text("Body"));

        var error = Assert.Throws<System.IO.InvalidDataException>(() => document.ToBytes());
        Assert.Contains("100,000", error.Message);
    }

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
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = parsed.GetPage(1).Letters;
        var normal = letters.Single(letter => letter.Value == "S");
        var subscript = letters.Single(letter => letter.Value == "H");
        Assert.True(subscript.StartBaseLine.Y < normal.StartBaseLine.Y);
        Assert.True(subscript.FontSize < normal.FontSize);
        Assert.True(raw.Split(new[] { " RG" }, System.StringSplitOptions.None).Length >= 4,
            "Expected a dashed underline and both lines of a double strikethrough.");
    }
}
