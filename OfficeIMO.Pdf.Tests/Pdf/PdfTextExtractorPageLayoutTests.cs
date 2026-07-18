using System;
using System.IO;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfTextExtractorPageTests {
    [Fact]
    public void ExtractTextByPage_WithLayoutOptionsUsesColumnAwareReadingOrder() {
        byte[] pdf = BuildTwoColumnPdf();

        IReadOnlyList<string> pages = PdfTextExtractor.ExtractTextByPage(pdf, new PdfTextLayoutOptions {
            MarginLeft = 36,
            MarginRight = 36,
            MinGutterWidth = 24
        });

        string text = Normalize(pages[0]);
        int leftStart = text.IndexOf("LeftStart", StringComparison.Ordinal);
        int leftFinish = text.IndexOf("LeftFinish", StringComparison.Ordinal);
        int rightStart = text.IndexOf("RightStart", StringComparison.Ordinal);
        int rightFinish = text.IndexOf("RightFinish", StringComparison.Ordinal);

        Assert.Single(pages);
        Assert.True(leftStart >= 0, "Expected left column start marker to be extracted.");
        Assert.True(leftFinish > leftStart, "Expected left column markers to preserve top-to-bottom order.");
        Assert.True(rightStart >= 0, "Expected right column start marker to be extracted.");
        Assert.True(rightFinish > rightStart, "Expected right column markers to preserve top-to-bottom order.");
        Assert.True(leftFinish < rightStart,
            $"Expected column-aware extraction to finish the left column before reading the right column. Text: {pages[0]}");
    }

    [Fact]
    public void GetTextSpans_UsesStandardFontMetricsWhenWidthsAreOmitted() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.TimesRoman,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("WWWW"))
            .ToBytes();

        PdfTextSpan span = Assert.Single(
            PdfReadDocument.Open(pdf)
                .Pages[0]
                .GetTextSpans(),
            item => item.Text == "WWWW");

        Assert.Equal(37.76, span.Advance, 2);
    }

    [Fact]
    public void GetTextSpans_UsesWinAnsiPunctuationMetricsWhenWidthsAreOmitted() {
        const string text = "\u201CWait\u201D\u2014ok\u2026";
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.TimesRoman,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text(text))
            .ToBytes();

        PdfTextSpan span = Assert.Single(
            PdfReadDocument.Open(pdf)
                .Pages[0]
                .GetTextSpans(),
            item => item.Text == text);

        Assert.Equal(58.32, span.Advance, 2);
    }

    [Fact]
    public void GetTextSpans_UsesWinAnsiAccentedLetterMetricsWhenWidthsAreOmitted() {
        const string text = "r\u00E9sum\u00E9";
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.TimesRoman,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text(text))
            .ToBytes();

        PdfTextSpan span = Assert.Single(
            PdfReadDocument.Open(pdf)
                .Pages[0]
                .GetTextSpans(),
            item => item.Text == text);

        Assert.Equal(28.88, span.Advance, 2);
    }
}
