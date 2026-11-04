using System;
using System.Buffers.Binary;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class DrawingPublisherMetricsTests {
        private static OfficeRasterCanvas Canvas(string feature = "dlig", string script = "latn", bool positioning = false) =>
            new OfficeRasterCanvas(new OfficeRasterImage(1, 1), fonts: new OfficeFontFaceCollection()
                .Add("Publisher fixture", ManagedTextShapingTestAssets.CreatePublisherMetricsFont(feature, script, positioning)));

        private static OfficeRichTextParagraph Paragraph(string text, double after = 0) => new OfficeRichTextParagraph(
            new[] { new OfficeRichTextRun(text, 10, OfficeColor.Black, fontFamily: "Publisher fixture") },
            margins: new OfficeTextPadding(0, 0, 0, after), lineHeightFactor: 1.19D) { UsesPublisherFontMetrics = true };

        [Theory]
        [InlineData("dlig", "latn", false, 10D, 8D)]
        [InlineData("dlig", "latn", true, 10D, 8D)]
        [InlineData("kern", "latn", false, 13D, 9.96D)]
        [InlineData("dlig", "arab", false, 13D, 9.96D)]
        public void SelectedFontControlsContentExtentAndIndependentReferenceBaseline(string feature, string script,
            bool positioning, double expectedContent, double expectedBaseline) {
            OfficeRasterCanvas canvas = Canvas(feature, script, positioning);
            var layout = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { Paragraph("A") },
                100, 100, canvas.CreateDrawingTextMetrics());
            OfficeRichTextLine line = Assert.Single(layout.Lines);
            Assert.Equal(expectedContent, layout.Height, 8);
            Assert.Equal(expectedContent * 1.19D, line.LineHeight, 8);
            Assert.Equal(expectedBaseline, OfficeTextBlockRenderer.ResolveRichTextRenderBaseline(line, 0, line.LineHeight, true), 8);
        }

        [Theory]
        [InlineData(9.99D, false)]
        [InlineData(10D, true)]
        public void TerminalContentFitsWithoutItsUnusedAdvanceAndFlowRetainsTheNextLine(double height, bool fits) {
            var source = Paragraph("A\nB");
            var metrics = Canvas().CreateDrawingTextMetrics();
            var flow = new OfficeRichTextFlow(new[] { source }, _ => { });
            var first = flow.Take(100, height, metrics, default);
            Assert.Equal(fits ? 1 : 0, first.Count);
            Assert.Equal(fits ? 2 : 0, flow.CharacterPosition);
            Assert.True(flow.HasRemaining);
            if (fits) {
                Assert.Equal("A\n", string.Concat(first[0].Runs.Select(run => run.Text)));
                var second = Assert.Single(flow.Take(100, 10, metrics, default));
                Assert.Equal("B", string.Concat(second.Runs.Select(run => run.Text)));
                Assert.False(flow.HasRemaining); Assert.Equal(3, flow.CharacterPosition);
                Assert.Equal(10D, OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { second }, 100, 10, metrics).Height, 8);
            }
            Assert.Equal("A\nB", source.Runs[0].Text);
        }

        [Theory]
        [InlineData(26.89D, false)]
        [InlineData(26.9D, true)]
        public void FollowingParagraphStillRequiresFullAdvanceAndAfterSpacing(double height, bool bothFit) {
            var metrics = Canvas().CreateDrawingTextMetrics();
            var source = new[] { Paragraph("A", 5), Paragraph("B", 20) };
            var flow = new OfficeRichTextFlow(source, _ => { });
            var visible = flow.Take(100, height, metrics, default);
            Assert.Equal(bothFit ? 2 : 1, visible.Count);
            Assert.Equal(!bothFit, flow.HasRemaining);
            Assert.Equal(bothFit ? 3 : 2, flow.CharacterPosition);
            Assert.Equal(0, visible[visible.Count - 1].Margins.Bottom);
            var layout = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(visible, 100, height, metrics);
            Assert.Equal(bothFit ? 26.9D : 10D, layout.Height, 8);
            Assert.Equal(5, source[0].Margins.Bottom); Assert.Equal(20, source[1].Margins.Bottom);
        }

        [Fact]
        public void OrdinaryParagraphKeepsGenericGeometryWithTheSameScopedFont() {
            var paragraph = new OfficeRichTextParagraph(Paragraph("A").Runs, lineHeightFactor: 1.19D);
            var layout = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { paragraph }, 100, 100,
                Canvas().CreateDrawingTextMetrics());
            OfficeRichTextLine line = Assert.Single(layout.Lines);
            Assert.Equal(12, layout.Height);
            Assert.Equal(9.4D, OfficeTextBlockRenderer.ResolveRichTextRenderBaseline(line, 0, line.LineHeight, true), 8);
        }

        [Fact]
        public void EmptyNativeParagraphRetainsItsTypingFontAndContentHeight() {
            var metrics = Canvas().CreateDrawingTextMetrics();
            var flow = new OfficeRichTextFlow(new[] { Paragraph("") }, _ => { });
            var visible = flow.Take(100, 10, metrics, default);
            Assert.Single(visible); Assert.False(flow.HasRemaining);
            Assert.Equal(10, OfficeDrawingTextLayout.CreateParagraphsWithMetrics(visible, 100, 10, metrics).Height);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void FallbackFacesShareTheWidthPaintAndFlowGeometry(bool completeFallback) {
            Assert.True(OfficeFontUnicodeRangeSet.TryParseCss("U+0041", out var primaryRange));
            Assert.True(OfficeFontUnicodeRangeSet.TryParseCss("U+0042", out var fallbackRange));
            var fonts = new OfficeFontFaceCollection();
            Assert.True(fonts.TryAdd("Primary", ManagedTextShapingTestAssets.CreatePublisherMetricsFont("dlig", "latn", false, 1400),
                OfficeFontStyle.Regular, primaryRange));
            Assert.True(fonts.TryAdd("Fallback", ManagedTextShapingTestAssets.CreatePublisherMetricsFont("dlig", "latn", false),
                OfficeFontStyle.Regular, completeFallback ? OfficeFontUnicodeRangeSet.All : fallbackRange));
            fonts.AddFallbackFamily("Fallback");
            var canvas = new OfficeRasterCanvas(new OfficeRasterImage(1, 1), fonts: fonts);
            var metrics = canvas.CreateDrawingTextMetrics();
            var paragraph = new OfficeRichTextParagraph(new[] {
                new OfficeRichTextRun("AB", 10, OfficeColor.Black, fontFamily: "Primary")
            }, lineHeightFactor: 1.19D) { UsesPublisherFontMetrics = true };
            var flow = new OfficeRichTextFlow(new[] { paragraph }, _ => { });
            Assert.Empty(flow.Take(100, 16.99, metrics, default));
            Assert.Equal(0, flow.CharacterPosition); Assert.True(flow.HasRemaining);
            var visible = flow.Take(100, 17, metrics, default);
            Assert.Equal("AB", string.Concat(Assert.Single(visible).Runs.Select(run => run.Text)));
            Assert.False(flow.HasRemaining); Assert.Equal(2, flow.CharacterPosition);
            var layout = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(visible, 100, 17, metrics);
            OfficeRichTextLine line = Assert.Single(layout.Lines);
            Assert.Equal(17, layout.Height, 8); Assert.Equal(20.23, line.LineHeight, 8);
            Assert.Equal(15, OfficeTextBlockRenderer.ResolveRichTextRenderBaseline(line, 0, line.LineHeight, true), 8);
            canvas.FontMetricScale = 2;
            var scaled = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(visible, 200, 34, canvas.CreateDrawingTextMetrics(), scale: 2);
            Assert.Equal(34, scaled.Height, 8);
            Assert.Equal(30, OfficeTextBlockRenderer.ResolveRichTextRenderBaseline(Assert.Single(scaled.Lines), 0, scaled.LineHeight, true), 8);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void LabelTypingStyleAndScaledIntrinsicPlacementKeepResolvedGeometry(bool emptyBody) {
            var run = new OfficeRichTextRun("B", emptyBody ? 10 : 20, OfficeColor.Red, fontFamily: "Publisher fixture");
            var body = new[] { new OfficeRichTextRun(emptyBody ? "" : "A", emptyBody ? 20 : 10, OfficeColor.Black, fontFamily: "Publisher fixture") };
            var paragraph = new OfficeRichTextParagraph(body,
                OfficeTextParagraphLabel.InBox(run, 0, 15), lineHeightFactor: 1.19D) { UsesPublisherFontMetrics = true };
            var drawing = new OfficeDrawing(200, 100).AddRichTextParagraphs(new[] { paragraph },
                0, 0, 200, 100, OfficeTextAreaAlignment.Center);
            var canvas = Canvas();
            var text = (OfficeDrawingRichText)drawing.Elements[0];
            var layout = OfficeDrawingTextLayout.CreateWithRasterMetrics(text, 200, 100, canvas);
            OfficeRichTextLine line = Assert.Single(layout.Lines);
            Assert.Equal("B", line.Segments[0].Text);
            Assert.Equal(20, layout.Height, 8); Assert.Equal(23.8, line.LineHeight, 8);
            Assert.Equal(16, OfficeTextBlockRenderer.ResolveRichTextRenderBaseline(line, 0, line.LineHeight, true), 8);
            canvas.FontMetricScale = 2;
            var scaled = OfficeDrawingTextLayout.CreateWithRasterMetrics(text, 400, 200, canvas, 2);
            OfficeRichTextLine larger = Assert.Single(scaled.Lines);
            Assert.Equal(layout.Height * 2, scaled.Height, 8);
            Assert.Equal(line.OffsetX * 2, larger.OffsetX, 8);
            Assert.Equal(32, OfficeTextBlockRenderer.ResolveRichTextRenderBaseline(larger, 0, larger.LineHeight, true), 8);
        }
    }
}

namespace OfficeIMO.TestAssets {
    internal static partial class ManagedTextShapingTestAssets {
        internal static byte[] CreatePublisherMetricsFont(string feature, string script, bool positioning, short typoAscent = 700) {
            var os2 = new byte[78];
            BinaryPrimitives.WriteInt16BigEndian(os2.AsSpan(68), typoAscent);
            BinaryPrimitives.WriteInt16BigEndian(os2.AsSpan(70), -200);
            BinaryPrimitives.WriteInt16BigEndian(os2.AsSpan(72), 100);
            BinaryPrimitives.WriteUInt16BigEndian(os2.AsSpan(74), 1000);
            BinaryPrimitives.WriteUInt16BigEndian(os2.AsSpan(76), 300);
            byte[] features = CreateMultipleGsub(scriptTag: script, featureTag: feature);
            return CreateFontFromCmap(CreateFormat12Cmap(new[] { (int)'A', (int)'B', 32 }), glyphCount: 5,
                os2: os2, gsub: positioning ? null : features, gpos: positioning ? features : null);
        }
    }
}
