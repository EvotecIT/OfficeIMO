using System.Globalization;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfInlineElementTests {
    [Fact]
    public void RichParagraph_RendersInlineImageAndBoxInContentOrderWithTags() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
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
                    background: new PdfColor(0.2D, 0.7D, 0.3D),
                    borderColor: PdfColor.Black,
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
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20,
                DefaultFontSize = 10
            })
            .Paragraph(paragraph => paragraph.Text("Prior flow marker"))
            .Paragraph(paragraph => paragraph
                .Text("Tall inline start ")
                .InlineBox(20, 34, background: new PdfColor(0.8D, 0.8D, 0.8D))
                .Text(" end"))
            .Paragraph(paragraph => paragraph.Text("Following flow marker"))
            .ToBytes();

        IReadOnlyList<PdfLogicalTextBlock> blocks = PdfDocument.Open(bytes).Read.TextBlocks();
        PdfLogicalTextBlock prior = Assert.Single(blocks, block => block.Text.Contains("Prior flow marker", StringComparison.Ordinal));
        PdfLogicalTextBlock first = Assert.Single(blocks, block => block.Text.Contains("Tall inline start", StringComparison.Ordinal));
        PdfLogicalTextBlock following = Assert.Single(blocks, block => block.Text.Contains("Following flow marker", StringComparison.Ordinal));

        Assert.True(prior.BaselineY - following.BaselineY >= 44D);
        Assert.True(first.BaselineY > following.BaselineY);
    }

    [Fact]
    public void RichParagraph_RejectsInlineElementWiderThanItsFrame() {
        PdfDocument document = PdfDocument.Create(new PdfOptions {
                PageWidth = 160,
                MarginLeft = 30,
                MarginRight = 30
            })
            .Paragraph(paragraph => paragraph.InlineBox(101, 12));

        ArgumentException exception = Assert.Throws<ArgumentException>(() => document.ToBytes());

        Assert.Contains("Inline element width", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RichParagraph_DecorativeInlineImageDrawsOnceAsArtifact() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
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
        Assert.Contains("/Artifact BMC", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/Figure << /Alt", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void RichParagraph_CentersInlineElementUsingItsWidth() {
        var fill = new PdfColor(0.13D, 0.57D, 0.91D);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                MarginLeft = 20,
                MarginRight = 20,
                CompressContentStreams = false
            })
            .Paragraph(paragraph => paragraph.InlineBox(40, 12, background: fill), PdfAlign.Center)
            .ToBytes();

        double x = FindFilledBoxX(PdfEncoding.Latin1GetString(bytes), fill, 40, 12);

        Assert.Equal(100D, x, 3);
    }

    [Fact]
    public void RichParagraph_LeadingTabAdvancesInlineElementToNextTabStop() {
        var fill = new PdfColor(0.17D, 0.61D, 0.83D);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                MarginLeft = 20,
                MarginRight = 20,
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
    public void TableShrinkToFit_PreservesInlineElementRuns() {
        const string longValue = "ThisIdentifierShouldShrinkToFit";
        var fill = new PdfColor(0.19D, 0.63D, 0.87D);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                CompressContentStreams = false
            })
            .Table(new[] {
                new[] {
                    PdfTableCell.RichTextCell(new[] {
                        TextRun.Normal(longValue, fontSize: 30, font: PdfStandardFont.Courier),
                        TextRun.Inline(new PdfInlineBox(12, 10, background: fill))
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
        PdfDocument document = PdfDocument.Open(bytes);
        IReadOnlyList<PdfLogicalTextBlock> blocks = document.Read.TextBlocks();
        string compactText = document.Read.Text()
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Replace(" ", string.Empty);

        Assert.Contains(longValue, compactText, StringComparison.Ordinal);
        Assert.Contains(blocks, block => block.FontSize < 30D);
        Assert.True(FindFilledBoxX(raw, fill, 12, 10) > 0D);
    }

    private static double FindFilledBoxX(string raw, PdfColor fill, double width, double height) {
        string color = string.Join(" ", new[] { fill.R, fill.G, fill.B }.Select(value => value.ToString("0.###", CultureInfo.InvariantCulture))) + " rg";
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
