using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class RichParagraphWrappingTests {
        [Fact]
        public void WrapRichRuns_CachesWidthMeasuredWithExplicitOpenTypeFeatures() {
            string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
            Assert.NotNull(fontPath);

            byte[] fontData = File.ReadAllBytes(fontPath!);
            var fontProgram = PdfOpenTypeCffFontProgram.Parse(fontData, "OfficeIMO Feature Width Font");
            var provider = new ControlledAdvanceTextShapingProvider(fontProgram, featureAware: true);
            var options = new PdfOptions()
                .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Feature Width Font")
                .SetTextShapingProvider(provider);
            OfficeTextFeatureSettings features = OfficeTextFeatureSettings.Default.With("kern", 0);
            PdfTextRun run = PdfTextRun.Normal("AV").WithFeatureSettings(features);

            object result = InvokeWrapRichRunsWithOptions(new[] { run }, 200D, 12D, PdfStandardFont.Helvetica, options);
            object segment = Assert.Single(Assert.Single(ExtractLines(result)));

            Assert.Equal(24D, ExtractMeasuredWidth(segment), precision: 6);
            Assert.Contains(provider.Requests, request =>
                request.Text == "AV" &&
                request.FeatureSettings.TryGetValue("kern", out int value) &&
                value == 0);

            provider.Requests.Clear();
            PdfDocument.Create(options)
                .Paragraph(paragraph => paragraph.Runs(new[] {
                    PdfTextRun.Normal("AV AV").WithFeatureSettings(features)
                }))
                .ToBytes();
            List<OfficeTextShapingRequest> spaceRequests = provider.Requests.FindAll(request => request.Text == " ");
            Assert.NotEmpty(spaceRequests);
            Assert.All(spaceRequests, request =>
                Assert.True(request.FeatureSettings.TryGetValue("kern", out int value) && value == 0));
        }

        [Fact]
        public void WrapRichRuns_CachesWholeChunkWidthAfterLinearLongTokenPlanning() {
            string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
            Assert.NotNull(fontPath);

            byte[] fontData = File.ReadAllBytes(fontPath!);
            var fontProgram = PdfOpenTypeCffFontProgram.Parse(fontData, "OfficeIMO Chunk Width Font");
            var provider = new ControlledAdvanceTextShapingProvider(fontProgram, featureAware: false);
            var options = new PdfOptions()
                .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "OfficeIMO Chunk Width Font")
                .SetTextShapingProvider(provider);

            object result = InvokeWrapRichRunsWithOptions(
                new[] { PdfTextRun.Normal("AAAA") },
                15D,
                12D,
                PdfStandardFont.Helvetica,
                options);
            List<List<object>> lines = ExtractLines(result);

            Assert.Equal(2, lines.Count);
            Assert.All(lines, line => Assert.Equal("AA", ExtractText(Assert.Single(line))));
            int pairAdvance = (fontProgram.UnitsPerEm / 3) * 2;
            double expectedWidth = pairAdvance * 12D / fontProgram.UnitsPerEm;
            Assert.All(lines, line => Assert.Equal(expectedWidth, ExtractMeasuredWidth(Assert.Single(line)), precision: 6));
            List<string> multiScalarRequests = provider.Requests
                .FindAll(request => request.Text.Length > 1)
                .ConvertAll(request => request.Text);
            Assert.Equal(new[] { "AAAA", "AA", "AA" }, multiScalarRequests);
        }

        private sealed class ControlledAdvanceTextShapingProvider : IOfficeTextShapingProvider {
            private readonly PdfOpenTypeCffFontProgram _fontProgram;
            private readonly bool _featureAware;

            internal ControlledAdvanceTextShapingProvider(PdfOpenTypeCffFontProgram fontProgram, bool featureAware) {
                _fontProgram = fontProgram;
                _featureAware = featureAware;
            }

            internal List<OfficeTextShapingRequest> Requests { get; } = new();

            public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
                Requests.Add(request);
                var glyphs = new List<OfficeShapedGlyph>();
                int scalarCount = CountScalars(request.Text);
                bool explicitKernDisabled = request.FeatureSettings.TryGetValue("kern", out int kern) && kern == 0;
                int advance = _featureAware
                    ? (explicitKernDisabled ? request.UnitsPerEm : request.UnitsPerEm / 2)
                    : (scalarCount > 1 ? request.UnitsPerEm / 3 : request.UnitsPerEm * 3 / 5);

                for (int index = 0; index < request.Text.Length;) {
                    int sourceIndex = index;
                    int scalar = ReadScalar(request.Text, ref index);
                    if (!_fontProgram.TryGetGlyphId(scalar, out int glyphId)) {
                        return null;
                    }

                    glyphs.Add(new OfficeShapedGlyph(glyphId, char.ConvertFromUtf32(scalar), sourceIndex, advance));
                }

                return new OfficeTextShapingResult(glyphs);
            }

            private static int CountScalars(string text) {
                int count = 0;
                for (int index = 0; index < text.Length; count++) {
                    ReadScalar(text, ref index);
                }

                return count;
            }

            private static int ReadScalar(string text, ref int index) {
                char value = text[index++];
                if (char.IsHighSurrogate(value) && index < text.Length && char.IsLowSurrogate(text[index])) {
                    return char.ConvertToUtf32(value, text[index++]);
                }

                return value;
            }
        }
    }
}
