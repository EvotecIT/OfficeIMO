using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTextQualityReviewRegressionTests {
    [Theory]
    [InlineData("abc \u05D0\u05D1\u05D2", true)]
    [InlineData("abc\u05D0\u05D1\u05D2", true)]
    [InlineData("abc\u05D0\u05D1\u05D2123xyz", true)]
    [InlineData("\u05D0\u05D1\u05D2123", true)]
    [InlineData("\u05D0\u05D1\u05D2 123", true)]
    [InlineData("\u05D0\u05D1\u05D2", false)]
    public void ScalarBidiReorderingRetainsLogicalTextExceptForExactRtlReversal(string text, bool marked) {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(text.Select(character => (int)character).Distinct().ToArray());
        PdfTrueTypeFontProgram font = PdfTrueTypeFontProgram.Parse(data);
        PdfGlyphRun run = font.ShapeText(text, PdfTextShapingOptions.ForRendering("Bidi", PdfTextShapingMode.UnicodeScalar));
        Assert.Equal(marked ? text : null, run.ActualText);
        Assert.NotEqual(Enumerable.Range(0, text.Length), run.Glyphs.Select(glyph => glyph.TextIndex));
        byte[] pdf = PdfDocument.Create(new PdfOptions { TextShapingMode = PdfTextShapingMode.UnicodeScalar }
                .UseFontFamily(new PdfEmbeddedFontFamily("Bidi", data)))
            .Canvas(canvas => canvas.Text(text, 20, 20, 500, 30)).ToBytes();
        Assert.Equal(text, PdfReadDocument.Open(pdf).ExtractText().Trim());
        if (marked && text.IndexOf(' ') < 0) {
            if (text.Any(char.IsDigit)) Assert.Contains("/ActualText", PdfTestContentText.ReadDecoded(pdf), StringComparison.Ordinal);
            else Assert.DoesNotContain("/ActualText", PdfTestContentText.ReadDecoded(pdf), StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData("ff", 0xFB00)]
    [InlineData("fi", 0xFB01)]
    [InlineData("fl", 0xFB02)]
    [InlineData("ffi", 0xFB03)]
    [InlineData("ffl", 0xFB04)]
    public void MixedScalarLatinLigatureRetainsItsSourceLetters(string letters, int ligature) {
        string text = letters + "\u05D0\u05D1\u05D2";
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(text.Select(character => (int)character).Concat(new[] { ligature }).Distinct().ToArray());
        PdfTrueTypeFontProgram font = PdfTrueTypeFontProgram.Parse(data);
        PdfGlyphRun run = font.ShapeText(text, PdfTextShapingOptions.ForRendering("Bidi", PdfTextShapingMode.LatinLigatures));
        Assert.Equal(letters, run.Glyphs[0].UnicodeText);
        byte[] pdf = PdfDocument.Create(new PdfOptions { TextShapingMode = PdfTextShapingMode.LatinLigatures }
                .UseFontFamily(new PdfEmbeddedFontFamily("Bidi", data)))
            .Canvas(canvas => canvas.Text(text, 20, 20, 500, 30)).ToBytes();
        Assert.Equal(text, PdfReadDocument.Open(pdf).ExtractText().Trim());
        Assert.Contains("/ActualText", PdfTestContentText.ReadDecoded(pdf), StringComparison.Ordinal);
    }

    [Fact]
    public void MixedSimpleAndCompositeJustificationDoesNotLeaveTwOnCompositeRuns() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', ' ');
        var options = new PdfOptions { PageWidth = 160, PageHeight = 200, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, DefaultFontSize = 20, CompressContentStreams = false };
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Embedded", data));
        byte[] pdf = PdfDocument.Create(options).Paragraph(p => p.Runs(new[] {
            new PdfTextRun("A ", font: PdfStandardFont.Helvetica),
            new PdfTextRun("A A A A A A A A A A A A A A A A", fontFamily: "Embedded")
        }), align: PdfAlign.Justify).ToBytes();
        string raw = Encoding.ASCII.GetString(pdf);
        Assert.Contains("/Subtype /Type0", raw, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Type1", raw, StringComparison.Ordinal);
        Match[] spacing = Regex.Matches(raw, @"(-?\d+(?:\.\d+)?) Tw").Cast<Match>().ToArray();
        Assert.NotEmpty(spacing);
        Assert.All(spacing, match => Assert.Equal(0D, double.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture)));
        Assert.True(PdfReadDocument.Open(pdf).Pages[0].GetTextSpans().Select(span => span.Y).Distinct().Count() > 1);
    }
}
