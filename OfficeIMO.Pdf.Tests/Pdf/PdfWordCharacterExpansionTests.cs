using System;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Letter = UglyToad.PdfPig.Content.Letter;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public sealed class PdfWordCharacterExpansionTests {
        [Theory]
        [InlineData("space-preferred", 3D, 0D, 0D, 0D, 1D, 2D)]
        [InlineData("character-preferred", 6.525D, 0.075D, 0.15D, 0.075D, 2D, 4D)]
        [InlineData("residual", 10.55D, 0.65D, 0.8D, 0.65D, 2.5D, 4.5D)]
        public void SourceSpacesThenSourceAveragesThenResidualOwnSavedGlyphOrigins(string caseName,
            double adjustment, double firstLetterDelta, double middleLetterDelta, double lastLetterDelta,
            double firstSpaceDelta, double secondSpaceDelta) {
            var runs = new[] {
                new PdfTextRun("AB", fontFamily: "Small Average", fontSize: global::OfficeIMO.OfficeLength.Points(12D)),
                new PdfTextRun(" ", fontFamily: "Small Average", fontSize: global::OfficeIMO.OfficeLength.Points(8D)),
                new PdfTextRun("CD", fontFamily: "Large Average", fontSize: global::OfficeIMO.OfficeLength.Points(16D)).WithHorizontalTextScaling(75D),
                new PdfTextRun(" ", fontFamily: "Small Average", fontSize: global::OfficeIMO.OfficeLength.Points(16D)),
                new PdfTextRun("EF", fontFamily: "Small Average", fontSize: global::OfficeIMO.OfficeLength.Points(12D))
            };
            Letter[] natural = Read(Render(runs, 200D, distributed: false), caseName + "-natural");
            double width = Span(natural) + adjustment;
            Letter[] actual = Read(Render(runs, width, distributed: true), caseName + "-actual");

            Assert.Equal("ABCDEF", string.Concat(actual.Select(letter => letter.Value)));
            Assert.Equal(firstLetterDelta, Gap(actual, 0, 1) - Gap(natural, 0, 1), 3);
            Assert.Equal(middleLetterDelta, Gap(actual, 2, 3) - Gap(natural, 2, 3), 3);
            Assert.Equal(lastLetterDelta, Gap(actual, 4, 5) - Gap(natural, 4, 5), 3);
            Assert.Equal(firstLetterDelta + firstSpaceDelta, Gap(actual, 1, 2) - Gap(natural, 1, 2), 3);
            Assert.Equal(middleLetterDelta + secondSpaceDelta, Gap(actual, 3, 4) - Gap(natural, 3, 4), 3);
            Assert.Equal(width, Span(actual), 3);
            AssertNaturalGlyphWidths(natural, actual);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void OptionalBreakKeepsWeightedWordGlyphPositions(bool embedded) {
            string? smallAverage = embedded ? "Small Average" : null;
            string? largeAverage = embedded ? "Large Average" : null;
            var controlRuns = new[] {
                new PdfTextRun("AB", fontFamily: smallAverage, fontSize: global::OfficeIMO.OfficeLength.Points(12D)),
                new PdfTextRun(" ", fontFamily: smallAverage, fontSize: global::OfficeIMO.OfficeLength.Points(8D)),
                new PdfTextRun("CD", fontFamily: largeAverage, fontSize: global::OfficeIMO.OfficeLength.Points(16D)).WithHorizontalTextScaling(75D),
                new PdfTextRun(" ", fontFamily: smallAverage, fontSize: global::OfficeIMO.OfficeLength.Points(16D)),
                new PdfTextRun("EF", fontFamily: smallAverage, fontSize: global::OfficeIMO.OfficeLength.Points(12D))
            };
            var optionalBreakRuns = (PdfTextRun[])controlRuns.Clone();
            optionalBreakRuns[2] = new PdfTextRun("C\u200BD", fontFamily: largeAverage, fontSize: global::OfficeIMO.OfficeLength.Points(16D))
                .WithHorizontalTextScaling(75D);
            string label = "optional-break-" + (embedded ? "embedded" : "standard");
            Letter[] natural = Read(Render(controlRuns, 200D, distributed: false), label + "-natural");
            double width = Span(natural) + 10.55D;
            Letter[] control = Read(Render(controlRuns, width, distributed: true), label + "-control");
            Letter[] actual = Read(Render(optionalBreakRuns, width, distributed: true), label + "-actual");

            Assert.Equal("ABCDEF", string.Concat(actual.Select(letter => letter.Value)));
            Assert.Equal(control.Length, actual.Length);
            for (int index = 0; index < control.Length; index++) {
                Assert.Equal(control[index].StartBaseLine.X, actual[index].StartBaseLine.X, 3);
                Assert.Equal(control[index].StartBaseLine.Y, actual[index].StartBaseLine.Y, 3);
                Assert.Equal(control[index].EndBaseLine.X, actual[index].EndBaseLine.X, 3);
            }
            Assert.Equal(width, Span(actual), 3);
            AssertNaturalGlyphWidths(control, actual);
        }

        [Fact]
        public void UnavailableSourceAverageLeavesItsGlyphBoundariesForResidualOnly() {
            var runs = new[] {
                new PdfTextRun("AB", fontFamily: "Small Average", fontSize: global::OfficeIMO.OfficeLength.Points(12D)),
                new PdfTextRun("CD", fontFamily: "Unavailable Average", fontSize: global::OfficeIMO.OfficeLength.Points(12D))
            };
            Letter[] natural = Read(Render(runs, 200D, distributed: false), "missing-natural");
            Letter[] actual = Read(Render(runs, Span(natural) + 0.15D, distributed: true), "missing-actual");

            Assert.Equal("ABCD", string.Concat(actual.Select(letter => letter.Value)));
            Assert.Equal(0.075D, Gap(actual, 0, 1) - Gap(natural, 0, 1), 3);
            Assert.Equal(0.075D, Gap(actual, 1, 2) - Gap(natural, 1, 2), 3);
            Assert.Equal(0D, Gap(actual, 2, 3) - Gap(natural, 2, 3), 3);
            Assert.Equal(Span(natural) + 0.15D, Span(actual), 3);
            AssertNaturalGlyphWidths(natural, actual);
        }

        [Fact]
        public void CffAndTrueTypeSourceAveragesKeepTheirSeparatePreferredShares() {
            string cffPath = Assert.IsType<string>(PdfComplianceTestFonts.FindBundledOpenTypeCffFont());
            byte[] cff = File.ReadAllBytes(cffPath);
            var runs = new[] {
                new PdfTextRun("AB", fontFamily: "Cff Average", fontSize: global::OfficeIMO.OfficeLength.Points(16D)),
                new PdfTextRun("CD", fontFamily: "Small Average", fontSize: global::OfficeIMO.OfficeLength.Points(12D))
            };
            // The bundled SourceSerif4 OS/2 xAvg is 558 with UPEM 1000;
            // the synthetic TrueType average is 400 with UPEM 1000.
            Letter[] natural = Read(Render(runs, 200D, distributed: false, cff), "cff-natural");
            // Keep the preferred-share budget above the page-size quantum so rounding
            // cannot ration the shares at the assertion's decimal midpoint.
            double width = (global::OfficeIMO.OfficeLength.Points(Span(natural) + 40D + 0.354D)
                + global::OfficeIMO.OfficeLength.Emus(2)).ToPoints() - 40D;
            Letter[] actual = Read(Render(runs, width, distributed: true, cff), "cff-actual");

            Assert.Equal("ABCD", string.Concat(actual.Select(letter => letter.Value)));
            Assert.Equal(0.1395D, Gap(actual, 0, 1) - Gap(natural, 0, 1), 3);
            Assert.Equal(0.1395D, Gap(actual, 1, 2) - Gap(natural, 1, 2), 3);
            Assert.Equal(0.075D, Gap(actual, 2, 3) - Gap(natural, 2, 3), 3);
            Assert.Equal(width, Span(actual), 3);
            AssertNaturalGlyphWidths(natural, actual);
        }

        private static byte[] Render(PdfTextRun[] runs, double width, bool distributed, byte[]? cff = null) {
            var options = new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(width + 40D), PageHeight = global::OfficeIMO.OfficeLength.Points(220D),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(20D), MarginRight = global::OfficeIMO.OfficeLength.Points(20D), MarginTop = global::OfficeIMO.OfficeLength.Points(20D), MarginBottom = global::OfficeIMO.OfficeLength.Points(20D),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12D), CompressContentStreams = false
            };
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Small Average",
                ManagedTextShapingTestAssets.CreateWordCharacterAverageFont(400)));
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Large Average",
                ManagedTextShapingTestAssets.CreateWordCharacterAverageFont(800)));
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Unavailable Average",
                ManagedTextShapingTestAssets.CreateWordCharacterAverageFont(null)));
            if (cff != null) options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Cff Average", cff));
            return PdfDocument.Create(options).Paragraph(paragraph => paragraph.Runs(runs),
                distributed ? PdfAlign.Justify : PdfAlign.Left,
                style: new PdfParagraphStyle {
                    UseWordJustificationSpacing = true, JustificationMode = PdfJustificationMode.Characters,
                    JustifyLastLine = true, TextWhitespaceMode = PdfTextWhitespaceMode.Preserve
                }).ToBytes();
        }

        private static double Span(Letter[] letters) => letters.Last().EndBaseLine.X - letters.First().StartBaseLine.X;
        private static double Gap(Letter[] letters, int left, int right) => letters[right].StartBaseLine.X - letters[left].EndBaseLine.X;
        private static void AssertNaturalGlyphWidths(Letter[] natural, Letter[] actual) {
            Assert.Equal(natural.Length, actual.Length);
            for (int index = 0; index < natural.Length; index++) Assert.Equal(natural[index].Width, actual[index].Width, 3);
        }
        private static Letter[] Read(byte[] bytes, string label) {
            string? output = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_ENGINE_PROOF_OUTPUT");
            if (!string.IsNullOrWhiteSpace(output)) {
                Directory.CreateDirectory(output!);
                File.WriteAllBytes(Path.Combine(output!, "character-expansion-" + label + "-" + Guid.NewGuid().ToString("N") + ".pdf"), bytes);
            }
            using var pdf = PdfPigDocument.Open(bytes);
            Assert.Equal(1, pdf.NumberOfPages);
            return pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .OrderBy(letter => letter.StartBaseLine.X).ToArray();
        }
    }
}

namespace OfficeIMO.TestAssets {
    internal static partial class ManagedTextShapingTestAssets {
        internal static byte[] CreateWordCharacterAverageFont(short? average) {
            byte[]? os2 = null;
            if (average.HasValue) {
                os2 = new byte[78];
                WriteUInt16(os2, 2, unchecked((ushort)average.Value));
                WriteUInt16(os2, 4, 400);
                WriteUInt16(os2, 74, 800);
                WriteUInt16(os2, 76, 200);
            }
            return CreateFontFromCmap(CreateDistinctFormat12Cmap(new[] { (int)' ', 'A', 'B', 'C', 'D', 'E', 'F' }),
                glyphCount: 8, os2: os2);
        }
    }
}
