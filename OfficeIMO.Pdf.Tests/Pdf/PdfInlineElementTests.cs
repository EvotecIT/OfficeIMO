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
        string text = PdfReadDocument.Load(bytes).ExtractText();

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

        IReadOnlyList<PdfLogicalTextBlock> blocks = PdfDocument.Load(bytes).Read.TextBlocks();
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
}
