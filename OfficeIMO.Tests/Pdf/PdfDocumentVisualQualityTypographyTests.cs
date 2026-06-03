using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void Paragraph_UsesNaturalWordSpacingForProportionalFonts() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 12
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("This sample uses proportional Helvetica text and should not stretch spaces between every word."))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var firstLineLetters = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .OrderByDescending(group => group.Key)
            .First()
            .OrderBy(letter => letter.StartBaseLine.X)
            .ToList();

        var gaps = firstLineLetters
            .Zip(firstLineLetters.Skip(1), (left, right) => right.StartBaseLine.X - left.EndBaseLine.X)
            .Where(gap => gap > 1)
            .ToList();

        Assert.NotEmpty(gaps);
        Assert.True(gaps.Max() < 9, $"Expected natural word spacing, but found a {gaps.Max():0.##}pt gap.");
    }

    [Fact]
    public void Paragraph_UsesProportionalGlyphWidthsForWrapping() {
        var options = new PdfOptions {
            PageWidth = 100,
            PageHeight = 160,
            MarginLeft = 35,
            MarginRight = 35,
            MarginTop = 25,
            MarginBottom = 25,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] narrowBytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("Illii Illii"))
            .ToBytes();

        byte[] wideBytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("WWWW"))
            .ToBytes();

        using var narrowPdf = PdfPigDocument.Open(new MemoryStream(narrowBytes));
        using var widePdf = PdfPigDocument.Open(new MemoryStream(wideBytes));

        Assert.Equal(1, CountTextLines(narrowPdf.GetPage(1)));
        Assert.True(CountTextLines(widePdf.GetPage(1)) >= 2, "Expected wide Helvetica glyphs to wrap instead of overrunning the text frame.");
    }

    [Fact]
    public void EmbeddedStandardFonts_UseTrueTypeMetricsForParagraphWrapping() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        PdfTrueTypeFontProgram fontProgram;
        try {
            fontProgram = PdfTrueTypeFontProgram.Parse(fontData, "OfficeIMOMetricsFont");
        } catch (NotSupportedException) {
            return;
        }

        string text = "iiii iiii iiii iiii";
        double embeddedProbeWidth = fontProgram.MeasureWinAnsiTextWidth("iiii iiii iiii", 10);
        double standardProbeWidth = PdfWriter.EstimateSimpleTextWidth("iiii iiii iiii", PdfStandardFont.Courier, 10);
        if (embeddedProbeWidth >= standardProbeWidth * 0.75D) {
            return;
        }

        var standardOptions = new PdfOptions {
            PageWidth = 92,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 25,
            MarginBottom = 25,
            DefaultFont = PdfStandardFont.Courier,
            DefaultFontSize = 10
        };
        var embeddedOptions = standardOptions.Clone()
            .EmbedStandardFont(PdfStandardFont.Courier, fontData, "OfficeIMOMetricsFont");

        byte[] standardBytes = PdfDocument.Create(standardOptions)
            .Paragraph(p => p.Text(text))
            .ToBytes();
        byte[] embeddedBytes = PdfDocument.Create(embeddedOptions)
            .Paragraph(p => p.Text(text))
            .ToBytes();

        using var standardPdf = PdfPigDocument.Open(new MemoryStream(standardBytes));
        using var embeddedPdf = PdfPigDocument.Open(new MemoryStream(embeddedBytes));
        int standardLineCount = CountTextLines(standardPdf.GetPage(1));
        int embeddedLineCount = CountTextLines(embeddedPdf.GetPage(1));
        string embeddedRaw = Encoding.ASCII.GetString(embeddedBytes);

        Assert.True(
            embeddedLineCount < standardLineCount,
            $"Expected embedded TrueType metrics to wrap fewer narrow-glyph lines than Courier metrics. Standard: {standardLineCount}; embedded: {embeddedLineCount}.");
        Assert.Contains("/BaseFont /OfficeIMOMetricsFont", embeddedRaw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Type0", embeddedRaw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /CIDFontType2", embeddedRaw, StringComparison.Ordinal);
        Assert.Contains("/Encoding /Identity-H", embeddedRaw, StringComparison.Ordinal);
        Assert.Contains("/CIDToGIDMap /Identity", embeddedRaw, StringComparison.Ordinal);
    }

    [Fact]
    public void Paragraph_JustifyExpandsWrappedLinesButNotFinalLine() {
        var options = new PdfOptions {
            PageWidth = 220,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 12
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("Alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau."), PdfAlign.Justify)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var lines = GetNonWhitespaceLetterLines(page);

        Assert.True(lines.Count >= 2, "Expected the justified paragraph to wrap onto at least two lines.");

        var firstLineGaps = GetInterWordGaps(lines[0]);
        var lastLineGaps = GetInterWordGaps(lines[lines.Count - 1]);

        Assert.NotEmpty(firstLineGaps);
        Assert.NotEmpty(lastLineGaps);
        Assert.True(firstLineGaps.Max() > 9, $"Expected justification to expand wrapped-line gaps, but the largest gap was {firstLineGaps.Max():0.##}pt.");
        Assert.True(lastLineGaps.Max() < 9, $"Expected the final justified paragraph line to keep natural spacing, but found a {lastLineGaps.Max():0.##}pt gap.");

        string extracted = page.Text;
        Assert.Contains("Alpha", extracted, StringComparison.Ordinal);
        Assert.Contains("omicron", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void Paragraph_JustifyDoesNotStretchExplicitLineBreaks() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 12
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p
                .Text("Alpha beta gamma")
                .LineBreak()
                .Text("Second line continues with enough words to wrap naturally."), PdfAlign.Justify)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var lines = GetNonWhitespaceLetterLines(page);

        Assert.True(lines.Count >= 2, "Expected the explicit line break to create a second rendered line.");

        var hardBreakLineGaps = GetInterWordGaps(lines[0]);

        Assert.NotEmpty(hardBreakLineGaps);
        Assert.True(hardBreakLineGaps.Max() < 9, $"Expected explicit line-break text to keep natural spacing, but found a {hardBreakLineGaps.Max():0.##}pt gap.");
    }

    [Fact]
    public void Heading_RightAlignsUsingProportionalTextWidth() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .H1("Illi", PdfAlign.Right)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double headingEndX = FindWordEndX(page, "Illi");
        double expectedRightEdge = options.PageWidth - options.MarginRight;

        Assert.InRange(Math.Abs(expectedRightEdge - headingEndX), 0, 5);
    }

    [Fact]
    public void ComposeItemHeading_AppliesExplicitAlignmentAndColor() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Column(column =>
                            column.Item(item =>
                                item.H2("ComposeHead", PdfAlign.Center, PdfColor.FromRgb(10, 20, 30)))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var line = GetVisualTextLines(page, 0, options.PageWidth)
            .Single(line => line.Text.Contains("ComposeHead", StringComparison.Ordinal));
        double contentCenter = (options.MarginLeft + options.PageWidth - options.MarginRight) / 2;
        double lineCenter = (line.X1 + line.X2) / 2;
        string rawPdf = Encoding.ASCII.GetString(bytes);

        Assert.InRange(Math.Abs(contentCenter - lineCenter), 0, 5);
        Assert.Contains("0.039 0.078 0.118 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnHeading_AppliesExplicitAlignmentAndColor() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.H3("ColumnHead", PdfAlign.Right, PdfColor.FromRgb(10, 20, 30)))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double headingEndX = FindWordEndX(page, "ColumnHead");
        double expectedRightEdge = options.PageWidth - options.MarginRight;
        string rawPdf = Encoding.ASCII.GetString(bytes);

        Assert.InRange(Math.Abs(expectedRightEdge - headingEndX), 0, 5);
        Assert.Contains("0.039 0.078 0.118 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void Heading_UsesProportionalGlyphWidthsForWideWrapping() {
        var options = new PdfOptions {
            PageWidth = 140,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .H3("WWWWWWWW")
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int wideLineCount = page.Letters
            .Where(letter => letter.Value == "W")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();
        double contentRight = options.PageWidth - options.MarginRight;
        double rightMostWideGlyph = page.Letters
            .Where(letter => letter.Value == "W")
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(wideLineCount > 1, "Expected heading wide glyphs to wrap using their real Helvetica advance instead of an average character width.");
        Assert.InRange(rightMostWideGlyph, double.NegativeInfinity, contentRight + 1);
    }

    [Fact]
    public void Heading_UsesProportionalGlyphWidthsWithoutOverWrappingNarrowText() {
        var options = new PdfOptions {
            PageWidth = 140,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .H3(new string('i', 20))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int narrowLineCount = page.Letters
            .Where(letter => letter.Value == "i")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();

        Assert.Equal(1, narrowLineCount);
    }

    [Fact]
    public void RowColumnHeading_UsesProportionalGlyphWidthsForWideWrapping() {
        var options = new PdfOptions {
            PageWidth = 140,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.H3("WWWWWWWW"))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int wideLineCount = page.Letters
            .Where(letter => letter.Value == "W")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();
        double contentRight = options.PageWidth - options.MarginRight;
        double rightMostWideGlyph = page.Letters
            .Where(letter => letter.Value == "W")
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(wideLineCount > 1, "Expected row-column heading wide glyphs to wrap using their real Helvetica advance instead of an average character width.");
        Assert.InRange(rightMostWideGlyph, double.NegativeInfinity, contentRight + 1);
    }

    [Fact]
    public void RowColumnHeading_UsesProportionalGlyphWidthsWithoutOverWrappingNarrowText() {
        var options = new PdfOptions {
            PageWidth = 140,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.H3(new string('i', 20)))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int narrowLineCount = page.Letters
            .Where(letter => letter.Value == "i")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();

        Assert.Equal(1, narrowLineCount);
    }

    [Fact]
    public void Footer_RightAlignsUsingProportionalTextWidth() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10,
            ShowPageNumbers = true,
            FooterFormat = "Illi",
            FooterFont = PdfStandardFont.Helvetica,
            FooterFontSize = 10,
            FooterAlign = PdfAlign.Right,
            FooterOffsetY = 12
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("Body"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double footerEndX = FindWordEndX(page, "Illi");
        double expectedRightEdge = options.PageWidth - options.MarginRight;

        Assert.InRange(Math.Abs(expectedRightEdge - footerEndX), 0, 5);
    }


}
