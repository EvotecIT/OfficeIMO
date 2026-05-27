using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public class RichParagraphWrappingTests {
        private static object InvokeWrapRichRuns(IEnumerable<TextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont) {
            var method = typeof(PdfWriter).GetMethod("WrapRichRuns", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);
            return method!.Invoke(null, new object?[] { runs, maxWidthPts, fontSize, baseFont, fontSize * 1.4, null })!;
        }

        private static T InvokePrivateFontMethod<T>(string methodName, params object[] parameters) {
            var method = typeof(PdfWriter).GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);
            return (T)method!.Invoke(null, parameters)!;
        }

        private static TargetInvocationException InvokePrivateFontMethodExpectingFailure(string methodName, params object[] parameters) {
            var method = typeof(PdfWriter).GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);
            return Assert.Throws<TargetInvocationException>(() => method!.Invoke(null, parameters));
        }

        private static List<List<object>> ExtractLines(object wrapResult) {
            var item1Field = wrapResult.GetType().GetField("Item1");
            Assert.NotNull(item1Field);
            var item1 = item1Field!.GetValue(wrapResult)!;
            var lines = new List<List<object>>();
            foreach (var lineObj in (IEnumerable)item1) {
                var segs = new List<object>();
                foreach (var segObj in (IEnumerable)lineObj) segs.Add(segObj);
                lines.Add(segs);
            }
            return lines;
        }

        private static PdfStandardFont ExtractFont(object seg) {
            var prop = seg.GetType().GetProperty("Font");
            Assert.NotNull(prop);
            return (PdfStandardFont)prop!.GetValue(seg)!;
        }

        private static string ExtractText(object seg) {
            var prop = seg.GetType().GetProperty("Text");
            Assert.NotNull(prop);
            return (string)prop!.GetValue(seg)!;
        }

        private static bool ExtractBold(object seg) {
            var prop = seg.GetType().GetProperty("Bold");
            Assert.NotNull(prop);
            return (bool)prop!.GetValue(seg)!;
        }

        private static bool ExtractLeadingSpace(object seg) {
            var prop = seg.GetType().GetProperty("LeadingSpace");
            Assert.NotNull(prop);
            return (bool)prop!.GetValue(seg)!;
        }

        private static double ExtractLeadingAdvance(object seg) {
            var prop = seg.GetType().GetProperty("LeadingAdvance");
            Assert.NotNull(prop);
            return (double)prop!.GetValue(seg)!;
        }

        private static bool ExtractLeadingSpaceIsExpandable(object seg) {
            var prop = seg.GetType().GetProperty("LeadingSpaceIsExpandable");
            Assert.NotNull(prop);
            return (bool)prop!.GetValue(seg)!;
        }

        private static PdfTextBaseline ExtractBaseline(object seg) {
            var prop = seg.GetType().GetProperty("Baseline");
            Assert.NotNull(prop);
            return (PdfTextBaseline)prop!.GetValue(seg)!;
        }

        [Fact]
        public void WriterFontSelection_NormalizesEveryStandardFontVariantToItsFamily() {
            Assert.Equal(PdfStandardFont.Helvetica, InvokePrivateFontMethod<PdfStandardFont>("ChooseNormal", PdfStandardFont.HelveticaOblique));
            Assert.Equal(PdfStandardFont.Helvetica, InvokePrivateFontMethod<PdfStandardFont>("ChooseNormal", PdfStandardFont.HelveticaBoldOblique));
            Assert.Equal(PdfStandardFont.TimesRoman, InvokePrivateFontMethod<PdfStandardFont>("ChooseNormal", PdfStandardFont.TimesItalic));
            Assert.Equal(PdfStandardFont.TimesRoman, InvokePrivateFontMethod<PdfStandardFont>("ChooseNormal", PdfStandardFont.TimesBoldItalic));
            Assert.Equal(PdfStandardFont.Courier, InvokePrivateFontMethod<PdfStandardFont>("ChooseNormal", PdfStandardFont.CourierOblique));
            Assert.Equal(PdfStandardFont.Courier, InvokePrivateFontMethod<PdfStandardFont>("ChooseNormal", PdfStandardFont.CourierBoldOblique));

            Assert.Equal(PdfStandardFont.TimesBold, InvokePrivateFontMethod<PdfStandardFont>("ChooseBold", PdfStandardFont.TimesBoldItalic));
            Assert.Equal(PdfStandardFont.CourierOblique, InvokePrivateFontMethod<PdfStandardFont>("ChooseItalic", PdfStandardFont.CourierBoldOblique));
            Assert.Equal(PdfStandardFont.HelveticaBoldOblique, InvokePrivateFontMethod<PdfStandardFont>("ChooseBoldItalic", PdfStandardFont.HelveticaBoldOblique));
        }

        [Fact]
        public void WriterFontSelection_RejectsInvalidFontValuesInsteadOfFallingBack() {
            var chooseException = InvokePrivateFontMethodExpectingFailure("ChooseNormal", (PdfStandardFont)99);
            Assert.IsType<ArgumentOutOfRangeException>(chooseException.InnerException);

            var glyphException = InvokePrivateFontMethodExpectingFailure("GlyphWidthEmFor", (PdfStandardFont)99);
            Assert.IsType<ArgumentOutOfRangeException>(glyphException.InnerException);

            var spaceException = InvokePrivateFontMethodExpectingFailure("SpaceWidthEmFor", (PdfStandardFont)99);
            Assert.IsType<ArgumentOutOfRangeException>(spaceException.InnerException);

            var ascenderException = InvokePrivateFontMethodExpectingFailure("GetAscender", (PdfStandardFont)99, 12.0);
            Assert.IsType<ArgumentOutOfRangeException>(ascenderException.InnerException);

            var descenderException = InvokePrivateFontMethodExpectingFailure("GetDescender", (PdfStandardFont)99, 12.0);
            Assert.IsType<ArgumentOutOfRangeException>(descenderException.InnerException);
        }

        [Fact]
        public void WriterFontSelection_UsesRequestedFontFamilyForGeneratedPdfResources() {
            var pdf = PdfDoc.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.TimesItalic,
                    FooterFont = PdfStandardFont.TimesRoman
                })
                .Paragraph(p => p.Text("Times family should stay Times."))
                .ToBytes();

            string content = System.Text.Encoding.ASCII.GetString(pdf);

            Assert.Contains("/BaseFont /Times-Roman", content);
            Assert.DoesNotContain("/BaseFont /Courier", content);
        }

        [Fact]
        public void WinAnsiEncoding_EncodesSupportedWindows1252CharactersWithoutFallback() {
            var bytes = PdfWinAnsiEncoding.Encode("\u20AC\u2022\u201C\u201D\u0152\u0178");

            Assert.Equal(new byte[] { 0x80, 0x95, 0x93, 0x94, 0x8C, 0x9F }, bytes);
            Assert.True(PdfWinAnsiEncoding.CanEncode("Invoice \u20AC \u2022", out int unsupportedIndex));
            Assert.Equal(-1, unsupportedIndex);
        }

        [Fact]
        public void EstimateSimpleTextWidth_UsesWinAnsiPunctuationWidths() {
            double width = InvokePrivateFontMethod<double>(
                "EstimateSimpleTextWidth",
                "\u201CWait\u201D\u2014ok\u2026",
                PdfStandardFont.TimesRoman,
                10.0);

            Assert.Equal(58.32, width, 2);
        }

        [Fact]
        public void EstimateSimpleTextWidth_UsesWinAnsiAccentedLetterWidths() {
            double width = InvokePrivateFontMethod<double>(
                "EstimateSimpleTextWidth",
                "r\u00E9sum\u00E9",
                PdfStandardFont.TimesRoman,
                10.0);

            Assert.Equal(28.88, width, 2);
        }

        [Fact]
        public void WinAnsiEncoding_RejectsUnsupportedGeneratedTextWithClearDiagnostic() {
            var exception = Assert.Throws<ArgumentException>(() =>
                PdfDoc.Create()
                    .Paragraph(p => p.Text("Snowman \u2603"))
                    .ToBytes());

            Assert.Contains("U+2603", exception.Message, StringComparison.Ordinal);
            Assert.Contains("PDF WinAnsiEncoding", exception.Message, StringComparison.Ordinal);
            Assert.Contains("embedded Unicode fonts", exception.Message, StringComparison.Ordinal);
            Assert.False(PdfWinAnsiEncoding.CanEncode("Snowman \u2603", out int unsupportedIndex));
            Assert.Equal(8, unsupportedIndex);
        }

        [Fact]
        public void WinAnsiEncoding_RejectsControlCharactersWithLayoutDiagnostic() {
            var exception = Assert.Throws<ArgumentException>(() => PdfWinAnsiEncoding.Encode("Alpha\tBeta"));

            Assert.Contains("control character U+0009", exception.Message, StringComparison.Ordinal);
            Assert.Contains("layout", exception.Message, StringComparison.Ordinal);
            Assert.False(PdfWinAnsiEncoding.CanEncode("Alpha\tBeta", out int unsupportedIndex));
            Assert.Equal(5, unsupportedIndex);
        }

        [Fact]
        public void WrapMonospace_PreservesExplicitLineBreaksForSimpleTextBlocks() {
            var lines = InvokePrivateFontMethod<List<string>>("WrapMonospace", "Alpha\nBeta\r\nGamma", 200.0, 12.0, 0.55);

            Assert.Equal(new[] { "Alpha", "Beta", "Gamma" }, lines);
        }

        [Fact]
        public void WrapMonospace_PreservesBlankHardBreakLines() {
            var lines = InvokePrivateFontMethod<List<string>>("WrapMonospace", "Alpha\n\nBeta", 200.0, 12.0, 0.55);

            Assert.Equal(new[] { "Alpha", string.Empty, "Beta" }, lines);
        }

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

        [Fact]
        public void WrapRichRuns_TreatsTabsAsWordLikeSpacing() {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("Alpha\tBeta")
            }, 200, 12, PdfStandardFont.Helvetica);

            var line = Assert.Single(ExtractLines(result));
            Assert.Equal(new[] { "Alpha", "Beta" }, line.ConvertAll(ExtractText).ToArray());
            Assert.False(ExtractLeadingSpace(line[0]));
            Assert.True(ExtractLeadingSpace(line[1]));
            Assert.False(ExtractLeadingSpaceIsExpandable(line[1]));
            Assert.True(ExtractLeadingAdvance(line[1]) > 3);
        }

        [Fact]
        public void WrapRichRuns_AdvancesTabsToDefaultHalfInchStops() {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("A\tB")
            }, 200, 12, PdfStandardFont.Helvetica);

            var line = Assert.Single(ExtractLines(result));
            Assert.Equal(new[] { "A", "B" }, line.ConvertAll(ExtractText).ToArray());
            Assert.InRange(ExtractLeadingAdvance(line[1]), 27, 29);
        }

        [Fact]
        public void ParagraphTabs_RenderAsVisibleDefaultTabStopGap() {
            byte[] bytes = PdfDoc.Create(new PdfOptions {
                    DefaultFontSize = 12
                })
                .Paragraph(p => p.Text("A B"), style: new PdfParagraphStyle {
                    SpacingAfter = 0
                })
                .Paragraph(p => p.Text("A\tB"), style: new PdfParagraphStyle {
                    SpacingAfter = 0
                })
                .ToBytes();

            using var pdf = UglyToad.PdfPig.PdfDocument.Open(new MemoryStream(bytes));
            var page = pdf.GetPage(1);
            var lineGaps = page.Letters
                .Where(letter => letter.Value == "A" || letter.Value == "B")
                .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
                .OrderByDescending(group => group.Key)
                .Select(group => {
                    var ordered = group.OrderBy(letter => letter.StartBaseLine.X).ToList();
                    var a = ordered.First(letter => letter.Value == "A");
                    var b = ordered.First(letter => letter.Value == "B");
                    return b.StartBaseLine.X - a.EndBaseLine.X;
                })
                .ToArray();

            Assert.Equal(2, lineGaps.Length);
            Assert.True(lineGaps[1] > lineGaps[0] + 10, $"Expected a default tab-stop gap rather than a collapsed single space. Plain gap: {lineGaps[0]:0.##}, tab gap: {lineGaps[1]:0.##}.");
        }

        [Fact]
        public void WrapRichRuns_DoesNotInsertSpaceBeforePunctuationAcrossRuns() {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("Hello", bold: true),
                new TextRun(", world")
            }, 200, 12, PdfStandardFont.Helvetica);

            var line = Assert.Single(ExtractLines(result));
            Assert.Equal(new[] { "Hello", ",", "world" }, line.ConvertAll(ExtractText).ToArray());
            Assert.False(ExtractLeadingSpace(line[1]));
            Assert.True(ExtractLeadingSpace(line[2]));
        }

        [Fact]
        public void WrapRichRuns_HonorsExplicitLineBreakRuns() {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("Status", bold: true),
                TextRun.LineBreak(),
                new TextRun("Healthy")
            }, 200, 12, PdfStandardFont.Helvetica);

            var lines = ExtractLines(result);

            Assert.Equal(2, lines.Count);
            Assert.Equal(new[] { "Status" }, lines[0].ConvertAll(ExtractText).ToArray());
            Assert.Equal(new[] { "Healthy" }, lines[1].ConvertAll(ExtractText).ToArray());
        }

        [Fact]
        public void WrapRichRuns_NormalizesCarriageReturnLineBreaks() {
            var result = InvokeWrapRichRuns(new[] {
                new TextRun("Alpha\r\nBeta\rGamma")
            }, 200, 12, PdfStandardFont.Helvetica);

            var lines = ExtractLines(result);

            Assert.Equal(3, lines.Count);
            Assert.Equal(new[] { "Alpha" }, lines[0].ConvertAll(ExtractText).ToArray());
            Assert.Equal(new[] { "Beta" }, lines[1].ConvertAll(ExtractText).ToArray());
            Assert.Equal(new[] { "Gamma" }, lines[2].ConvertAll(ExtractText).ToArray());
        }
    }
}

