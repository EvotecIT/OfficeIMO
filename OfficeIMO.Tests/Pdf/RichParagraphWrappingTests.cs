using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public class RichParagraphWrappingTests {
        private static object InvokeWrapRichRuns(IEnumerable<TextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont) {
            var method = typeof(PdfWriter).GetMethod("WrapRichRuns", BindingFlags.NonPublic | BindingFlags.Static);
            Assert.NotNull(method);
            return method!.Invoke(null, new object[] { runs, maxWidthPts, fontSize, baseFont })!;
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
    }
}

