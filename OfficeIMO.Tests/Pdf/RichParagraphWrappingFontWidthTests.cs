using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class RichParagraphWrappingTests {
        [Fact]
        public void WrapRichRuns_MixedStylesWrapLikePlainText() {
            var baseFont = PdfStandardFont.Helvetica;
            double fontSize = 12;
            double maxWidth = 70;
            var mixedRuns = new[] {
                new TextRun("Alpha ", bold: true),
                new TextRun("Beta "),
                new TextRun("Gamma ", bold: true),
                new TextRun("Delta")
            };
            var plainRuns = new[] { new TextRun("Alpha Beta Gamma Delta") };

            var mixedResult = InvokeWrapRichRuns(mixedRuns, maxWidth, fontSize, baseFont);
            var plainResult = InvokeWrapRichRuns(plainRuns, maxWidth, fontSize, baseFont);

            var mixedLines = ExtractLines(mixedResult);
            var plainLines = ExtractLines(plainResult);

            var expectedLines = new[] {
                new[] { "Alpha", "Beta" },
                new[] { "Gamma" },
                new[] { "Delta" }
            };

            Assert.Equal(expectedLines.Length, mixedLines.Count);
            Assert.Equal(expectedLines.Length, plainLines.Count);

            for (int i = 0; i < expectedLines.Length; i++) {
                var expectedTokens = expectedLines[i];
                var mixedTokens = mixedLines[i].ConvertAll(ExtractText).ToArray();
                var plainTokens = plainLines[i].ConvertAll(ExtractText).ToArray();
                Assert.Equal(expectedTokens, mixedTokens);
                Assert.Equal(expectedTokens, plainTokens);
            }

            var chooseBold = typeof(PdfWriter).GetMethod("ChooseBold", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(chooseBold);
            var expectedBoldFont = (PdfStandardFont)chooseBold!.Invoke(null, new object[] { baseFont })!;
            var expectedNormalFont = baseFont;

            foreach (var line in mixedLines) {
                foreach (var seg in line) {
                    var font = ExtractFont(seg);
                    if (ExtractBold(seg)) Assert.Equal(expectedBoldFont, font);
                    else Assert.Equal(expectedNormalFont, font);
                }
            }
        }

        [Fact]
        public void WrapRichRuns_UsesStandardGlyphWidthsForProportionalWrapping() {
            var narrowResult = InvokeWrapRichRuns(new[] {
                new TextRun("Illii Illii")
            }, 30, 10, PdfStandardFont.Helvetica);

            var narrowLine = Assert.Single(ExtractLines(narrowResult));
            Assert.Equal(new[] { "Illii", "Illii" }, narrowLine.ConvertAll(ExtractText).ToArray());

            var wideResult = InvokeWrapRichRuns(new[] {
                new TextRun("WWWW")
            }, 30, 10, PdfStandardFont.Helvetica);

            var wideLines = ExtractLines(wideResult);
            Assert.True(wideLines.Count >= 2, "Expected wide glyphs to wrap by measured width instead of average character count.");
            Assert.Equal("WWWW", string.Concat(wideLines.SelectMany(line => line.ConvertAll(ExtractText))));
        }

        [Fact]
        public void EstimateSimpleTextWidth_UsesHelveticaBoldGlyphWidths() {
            double regularWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "nnnnn", PdfStandardFont.Helvetica, 10.0);
            double boldWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "nnnnn", PdfStandardFont.HelveticaBold, 10.0);
            double boldObliqueWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "nnnnn", PdfStandardFont.HelveticaBoldOblique, 10.0);

            Assert.Equal(27.8, regularWidth, 2);
            Assert.Equal(30.55, boldWidth, 2);
            Assert.Equal(boldWidth, boldObliqueWidth, 2);
        }

        [Fact]
        public void WrapRichRuns_UsesHelveticaBoldGlyphWidthsForStyledRuns() {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("nnnnn", bold: true)
            }, 30, 10, PdfStandardFont.Helvetica);

            var lines = ExtractLines(result);

            Assert.True(lines.Count >= 2, "Expected bold Helvetica widths to drive wrapping for styled runs.");
            Assert.Equal("nnnnn", string.Concat(lines.SelectMany(line => line.ConvertAll(ExtractText))));
            Assert.All(lines.SelectMany(line => line), seg => Assert.True(ExtractBold(seg)));
        }

        [Fact]
        public void EstimateSimpleTextWidth_UsesTimesGlyphWidthsInsteadOfAverageFallback() {
            double narrowWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "Illii", PdfStandardFont.TimesRoman, 10.0);
            double wideWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "WWWW", PdfStandardFont.TimesRoman, 10.0);

            Assert.True(narrowWidth < 16, $"Expected Times narrow glyphs to measure narrowly, but width was {narrowWidth:0.##}pt.");
            Assert.True(wideWidth > 37, $"Expected Times W glyphs to measure widely, but width was {wideWidth:0.##}pt.");
        }

        [Fact]
        public void EstimateSimpleTextWidth_UsesTimesVariantGlyphWidths() {
            double romanW = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "W", PdfStandardFont.TimesRoman, 10.0);
            double boldW = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "W", PdfStandardFont.TimesBold, 10.0);
            double italicW = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "W", PdfStandardFont.TimesItalic, 10.0);
            double boldItalicW = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "W", PdfStandardFont.TimesBoldItalic, 10.0);

            Assert.Equal(9.44, romanW, 2);
            Assert.Equal(10.0, boldW, 2);
            Assert.Equal(8.33, italicW, 2);
            Assert.Equal(8.89, boldItalicW, 2);
        }

        [Fact]
        public void WrapRichRuns_UsesTimesGlyphWidthsForWideCharacters() {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("WWWW")
            }, 25, 10, PdfStandardFont.TimesRoman);

            var lines = ExtractLines(result);

            Assert.True(lines.Count >= 2, "Expected wide Times glyphs to wrap by measured width instead of average character count.");
            Assert.Equal("WWWW", string.Concat(lines.SelectMany(line => line.ConvertAll(ExtractText))));
        }

        [Fact]
        public void WrapRichRuns_UsesWinAnsiPunctuationWidthsForWrapping() {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("\u201CWait\u201D\u2014ok\u2026")
            }, 50, 10, PdfStandardFont.TimesRoman);

            var lines = ExtractLines(result);

            Assert.True(lines.Count >= 2, "Expected smart punctuation and em dash widths to drive wrapping.");
            Assert.Equal("\u201CWait\u201D\u2014ok\u2026", string.Concat(lines.SelectMany(line => line.ConvertAll(ExtractText))));
        }

        [Fact]
        public void WrapRichRuns_UsesWinAnsiAccentedLetterWidthsForWrapping() {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("r\u00E9sum\u00E9")
            }, 29, 10, PdfStandardFont.TimesRoman);

            var lines = ExtractLines(result);

            Assert.Single(lines);
            Assert.Equal("r\u00E9sum\u00E9", string.Concat(lines.SelectMany(line => line.ConvertAll(ExtractText))));
        }

        [Fact]
        public void WrapRichRuns_UsesTimesBoldGlyphWidthsForStyledRuns() {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("WWW", bold: true)
            }, 29, 10, PdfStandardFont.TimesRoman);

            var lines = ExtractLines(result);

            Assert.True(lines.Count >= 2, "Expected bold Times glyph widths to drive wrapping for styled runs.");
            Assert.Equal("WWW", string.Concat(lines.SelectMany(line => line.ConvertAll(ExtractText))));
            Assert.All(lines.SelectMany(line => line), seg => Assert.True(ExtractBold(seg)));
        }

        [Fact]
        public void WrapRichRuns_UsesScaledWidthsForSuperscriptAndSubscriptRuns() {
            var superscriptResult = InvokeWrapRichRuns(new[] {
                new TextRun("Wide"),
                TextRun.Superscript("9999")
            }, 45, 12, PdfStandardFont.Helvetica);

            var normalResult = InvokeWrapRichRuns(new[] {
                new TextRun("Wide"),
                new TextRun("9999")
            }, 45, 12, PdfStandardFont.Helvetica);

            var superscriptLine = Assert.Single(ExtractLines(superscriptResult));
            var normalLines = ExtractLines(normalResult);

            Assert.Equal(new[] { "Wide", "9999" }, superscriptLine.ConvertAll(ExtractText).ToArray());
            Assert.Equal(PdfTextBaseline.Superscript, ExtractBaseline(superscriptLine[1]));
            Assert.True(normalLines.Count > 1, "Expected unscaled text to wrap where superscript can remain on the same line.");
        }
    }
}
