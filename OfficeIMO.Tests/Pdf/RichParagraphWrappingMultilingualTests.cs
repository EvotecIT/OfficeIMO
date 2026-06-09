using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class RichParagraphWrappingTests {
        [Fact]
        public void WrapRichRuns_UsesMultilingualBreakpointsForCjkClosingPunctuation() {
            const string text = "\u65E5\u672C\u6771\u4EAC\u3001\u5927\u962A\u4EAC\u90FD";
            double maxWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "\u65E5\u672C\u6771\u4EAC", PdfStandardFont.Helvetica, 10.0);

            var result = InvokeWrapRichRuns(new[] {
                new TextRun(text)
            }, maxWidth, 10, PdfStandardFont.Helvetica);

            var lines = ExtractLines(result);
            var lineTexts = lines.Select(line => string.Concat(line.Select(ExtractText))).ToArray();

            Assert.True(lineTexts.Length > 1, "Expected CJK text to wrap across multiple lines.");
            Assert.Equal(text, string.Concat(lineTexts));
            Assert.DoesNotContain(lineTexts.Skip(1), line => line.StartsWith("\u3001", System.StringComparison.Ordinal));
        }

        [Fact]
        public void WrapSimpleText_UsesMultilingualBreakpointsForCjkClosingPunctuation() {
            const string text = "\u65E5\u672C\u6771\u4EAC\u3001\u5927\u962A\u4EAC\u90FD";
            double maxWidth = InvokePrivateFontMethod<double>("EstimateSimpleTextWidth", "\u65E5\u672C\u6771\u4EAC", PdfStandardFont.Helvetica, 10.0);

            var lines = InvokePrivateFontMethod<List<string>>("WrapSimpleText", text, maxWidth, PdfStandardFont.Helvetica, 10.0);

            Assert.True(lines.Count > 1, "Expected CJK text to wrap across multiple lines.");
            Assert.Equal(text, string.Concat(lines));
            Assert.DoesNotContain(lines.Skip(1), line => line.StartsWith("\u3001", System.StringComparison.Ordinal));
        }
    }
}
