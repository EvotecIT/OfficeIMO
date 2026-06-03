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
