using System.Globalization;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfInlineElementTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TableRowShrinkingReservesFixedInlineWidth(bool canvas) {
        global::OfficeIMO.OfficeColor fill = global::OfficeIMO.OfficeColor.FromNormalizedRgb(.2D, .5D, .8D);
        PdfTableCell cell = new PdfTableCell(new[] {
            PdfTextRun.Inline(new PdfInlineBox(30, 10, background: fill)),
            PdfTextRun.Normal("ABCDE")
        }).WithNoWrap();
        PdfTableStyle style = new PdfTableStyle {
            FontSize = 12, HeaderRowCount = 0, HeaderBold = false,
            ColumnWidthPoints = new() { 50 }, MaxWidth = 50, PreserveWidth = true,
            CellPaddingX = 0, CellPaddingY = 0, MinimumShrinkFontSize = 4,
            FixedRowHeights = new() { 30 }, BorderColor = null, ShrinkTextToFit = true
        };
        PdfDocument document = PdfDocument.Create(new PdfCreateOptions {
            DefaultFont = PdfStandardFont.Courier, PageWidth = global::OfficeIMO.OfficeLength.Points(200), PageHeight = global::OfficeIMO.OfficeLength.Points(150),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(20), MarginRight = global::OfficeIMO.OfficeLength.Points(20), MarginTop = global::OfficeIMO.OfficeLength.Points(20), MarginBottom = global::OfficeIMO.OfficeLength.Points(20),
            CompressContentStreams = false
        });
        PdfTableCell[][] rows = new[] { new[] { cell } };
        if (canvas) document.Canvas(surface => surface.Table(rows, 20, 20, 50, 30, style));
        else document.Table(rows, style: style);
        byte[] bytes = document.ToBytes();
        using UglyToad.PdfPig.PdfDocument pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        UglyToad.PdfPig.Content.Word word = Assert.Single(pdf.GetPage(1).GetWords());
        Assert.Equal("ABCDE", word.Text);
        Assert.All(word.Letters, letter => Assert.InRange(letter.PointSize, 6.66D, 6.67D));
        Assert.True(word.BoundingBox.Right <= 70.01D);
        Assert.Equal(20D, FindFilledBoxX(PdfEncoding.Latin1GetString(bytes), fill, 30, 10), 3);
    }

    [Fact]
    public void RichParagraph_RendersInlineImageAndBoxInContentOrderWithTags() {
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false,
                TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers
            })
            .Paragraph(paragraph => paragraph
                .Text("Before ")
                .InlineImage(PdfPngTestImages.CreateRgbPng(2, 1), 24, 14, "Inline status image")
                .Text(" middle ")
                .InlineBox(
                    18,
                    12,
                    background: global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.2D, 0.7D, 0.3D),
                    borderColor: global::OfficeIMO.OfficeColor.Black,
                    borderWidth: 1D,
                    alternativeText: "Inline status box")
                .Text(" after"))
            .ToBytes();

        string raw = PdfEncoding.Latin1GetString(bytes);
        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Contains("Before", text, StringComparison.Ordinal);
        Assert.Contains("middle", text, StringComparison.Ordinal);
        Assert.Contains("after", text, StringComparison.Ordinal);
        Assert.Single(PdfImageExtractor.ExtractImages(bytes));
        Assert.Contains("/Figure << /Alt <496E6C696E652073746174757320696D616765>", raw, StringComparison.Ordinal);
        Assert.Contains("/Figure << /Alt <496E6C696E652073746174757320626F78>", raw, StringComparison.Ordinal);
        Assert.Contains("0.2 0.7 0.3 rg", raw, StringComparison.Ordinal);
        Assert.Contains("1 w", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void RichParagraph_InlineHeightAdvancesFollowingFlow() {
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(240),
                PageHeight = global::OfficeIMO.OfficeLength.Points(180),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(20),
                MarginRight = global::OfficeIMO.OfficeLength.Points(20),
                MarginTop = global::OfficeIMO.OfficeLength.Points(20),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(20),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
            })
            .Paragraph(paragraph => paragraph.Text("Prior flow marker"))
            .Paragraph(paragraph => paragraph
                .Text("Tall inline start ")
                .InlineBox(20, 34, background: global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.8D, 0.8D, 0.8D))
                .Text(" end"))
            .Paragraph(paragraph => paragraph.Text("Following flow marker"))
            .ToBytes();

        IReadOnlyList<PdfLogicalTextBlock> blocks = PdfDocument.Load(bytes).Reader.TextBlocks();
        PdfLogicalTextBlock prior = Assert.Single(blocks, block => block.Text.Contains("Prior flow marker", StringComparison.Ordinal));
        PdfLogicalTextBlock first = Assert.Single(blocks, block => block.Text.Contains("Tall inline start", StringComparison.Ordinal));
        PdfLogicalTextBlock following = Assert.Single(blocks, block => block.Text.Contains("Following flow marker", StringComparison.Ordinal));

        Assert.True(prior.BaselineY - following.BaselineY >= 44D);
        Assert.True(first.BaselineY > following.BaselineY);
    }

    [Fact]
    public void RichParagraph_RejectsInlineElementWiderThanItsFrame() {
        PdfDocument document = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(160),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30)
            })
            .Paragraph(paragraph => paragraph.InlineBox(101, 12));

        ArgumentException exception = Assert.Throws<ArgumentException>(() => document.ToBytes());

        Assert.Contains("Inline element width", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RichParagraph_DecorativeInlineImageDrawsOnceAsArtifact() {
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false,
                TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers
            })
            .Paragraph(paragraph => paragraph
                .Text("Before ")
                .InlineImage(PdfPngTestImages.CreateRgbPng(2, 1), 24, 14)
                .Text(" after"))
            .ToBytes();

        string raw = PdfEncoding.Latin1GetString(bytes);

        Assert.Single(PdfImageExtractor.ExtractImages(bytes));
        Assert.Equal(1, CountOccurrences(raw, "/Im1 Do"));
        Assert.Contains("/Artifact << /Type /Layout >> BDC", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/Figure << /Alt", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void RichParagraph_CentersInlineElementUsingItsWidth() {
        var fill = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.13D, 0.57D, 0.91D);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(240),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(20),
                MarginRight = global::OfficeIMO.OfficeLength.Points(20),
                CompressContentStreams = false
            })
            .Paragraph(paragraph => paragraph.InlineBox(40, 12, background: fill), PdfAlign.Center)
            .ToBytes();

        double x = FindFilledBoxX(PdfEncoding.Latin1GetString(bytes), fill, 40, 12);

        Assert.Equal(100D, x, 3);
    }

    [Fact]
    public void RichParagraph_LeadingTabAdvancesInlineElementToNextTabStop() {
        var fill = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.17D, 0.61D, 0.83D);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(240),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(20),
                MarginRight = global::OfficeIMO.OfficeLength.Points(20),
                CompressContentStreams = false
            })
            .Paragraph(paragraph => paragraph
                .Tab()
                .InlineBox(18, 12, background: fill))
            .ToBytes();

        double x = FindFilledBoxX(PdfEncoding.Latin1GetString(bytes), fill, 18, 12);

        Assert.Equal(56D, x, 3);
    }

    [Fact]
    public void RichParagraph_CollapsedSpaceRetainsTabBeforeInlineElement() {
        const double margin = 20D, stop = 72D;
        var fill = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.17D, 0.61D, 0.83D);
        var style = new PdfParagraphStyle { TextWhitespaceMode = PdfTextWhitespaceMode.Collapse, SpacingAfter = 0D };
        style.AddTabStop(stop, leader: PdfTabLeaderStyle.Dots);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(240D), MarginLeft = global::OfficeIMO.OfficeLength.Points(margin), MarginRight = global::OfficeIMO.OfficeLength.Points(margin),
            DefaultFont = PdfStandardFont.Courier, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12D), CompressContentStreams = false
        }).Paragraph(p => p.Text("A").Tab().Text(" ").InlineBox(18, 12, background: fill), style: style)
            .ToBytes();

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        var a = Assert.Single(page.Letters, letter => letter.Value == "A");
        Assert.Contains(page.Letters, letter => letter.Value == ".");
        double monospaceAdvance = a.EndBaseLine.X - a.StartBaseLine.X;
        Assert.True(monospaceAdvance > 0D);
        double x = FindFilledBoxX(PdfEncoding.Latin1GetString(bytes), fill, 18, 12);
        Assert.Equal(margin + stop + monospaceAdvance, x, 2);
    }

    [Fact]
    public void TableShrinkToFit_PreservesInlineElementRuns() {
        const string longValue = "ThisIdentifierShouldShrinkToFit";
        var fill = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.19D, 0.63D, 0.87D);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                CompressContentStreams = false
            })
            .Table(new[] {
                new[] {
                    PdfTableCell.RichTextCell(new[] {
                        PdfTextRun.Normal(longValue, fontSize: 30, font: PdfStandardFont.Courier),
                        PdfTextRun.Inline(new PdfInlineBox(12, 10, background: fill))
                    })
                }
            }, style: new PdfTableStyle {
                FontSize = 18,
                MinimumShrinkFontSize = 7,
                ShrinkTextToFit = true,
                ColumnWidthPoints = new List<double?> { 108 }
            })
            .ToBytes();

        string raw = PdfEncoding.Latin1GetString(bytes);
        PdfDocument document = PdfDocument.Load(bytes);
        IReadOnlyList<PdfLogicalTextBlock> blocks = document.Reader.TextBlocks();
        string compactText = document.Reader.Text()
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Replace(" ", string.Empty);

        Assert.Contains(longValue, compactText, StringComparison.Ordinal);
        Assert.Contains(blocks, block => block.FontSize < 30D);
        Assert.True(FindFilledBoxX(raw, fill, 12, 10) > 0D);
    }

    private static double FindFilledBoxX(string raw, global::OfficeIMO.OfficeColor fill, double width, double height) {
        string color = string.Join(" ", new[] { fill.ToNormalizedRgb().R, fill.ToNormalizedRgb().G, fill.ToNormalizedRgb().B }.Select(value => value.ToString("0.###", CultureInfo.InvariantCulture))) + " rg";
        string pattern = Regex.Escape(color) + @"\s+(?<x>-?\d+(?:\.\d+)?)\s+-?\d+(?:\.\d+)?\s+" +
            Regex.Escape(width.ToString("0.###", CultureInfo.InvariantCulture)) + @"\s+" +
            Regex.Escape(height.ToString("0.###", CultureInfo.InvariantCulture)) + @"\s+re\s+f";
        Match match = Regex.Match(raw, pattern, RegexOptions.CultureInvariant);
        Assert.True(match.Success, "Expected the inline box fill and rectangle in the generated content stream.");
        return double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture);
    }

    private static int CountOccurrences(string value, string token) {
        int count = 0;
        int offset = 0;
        while ((offset = value.IndexOf(token, offset, StringComparison.Ordinal)) >= 0) {
            count++;
            offset += token.Length;
        }

        return count;
    }
}
