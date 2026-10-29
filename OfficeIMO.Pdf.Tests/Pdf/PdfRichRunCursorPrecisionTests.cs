using System;
using System.Linq;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf {
    public sealed class PdfRichRunCursorPrecisionTests {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void AdjacentPaintRunsStartAtTheirMeasuredPositions(bool useFontUnits) {
            using var pdf = PdfPigDocument.Open(CreateFixture(useFontUnits));
            var page = Assert.Single(pdf.GetPages());
            var letters = page.Letters.ToArray();
            Assert.Equal(new string('A', 250) + "B", string.Concat(letters.Select(letter => letter.Value)));
            double runAdvance = useFontUnits ? 50D * 1229D / 2048D * 20D : 50D * .6D * 20D;
            for (int index = 1; index <= 5; index++) {
                Assert.Equal(20D + index * runAdvance, letters[index * 50].StartBaseLine.X, 3);
            }
            var underline = Assert.Single(page.Paths.Where(path => path.IsStroked)
                .Select(path => path.GetBoundingRectangle()), bounds => bounds.HasValue);
            Assert.Equal(letters[250].StartBaseLine.X, underline!.Value.Left, 3);
        }

        internal static byte[] CreateFixture(bool useFontUnits) {
            byte[] font = ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(2048, 1229, 'A', 'B');
            var options = new PdfCreateOptions {
                PageWidth = 4000D, PageHeight = 200D, DefaultFontSize = 20D,
                MarginLeft = 20D, MarginRight = 20D, MarginTop = 20D, MarginBottom = 20D,
                UseFontUnitTextLayout = useFontUnits
            };
            options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName, font));
            PdfTextRun[] runs = Enumerable.Range(0, 5).Select(index => new PdfTextRun(new string('A', 50),
                fontFamily: ManagedTextShapingTestAssets.FamilyName,
                color: index % 2 == 0 ? global::OfficeIMO.OfficeColor.Black : global::OfficeIMO.OfficeColor.FromRgb(0, 0, 255))).Concat(new[] {
                    new PdfTextRun("B", fontFamily: ManagedTextShapingTestAssets.FamilyName, underline: true,
                        backgroundColor: global::OfficeIMO.OfficeColor.FromRgb(255, 255, 0))
                }).ToArray();
            return PdfDocument.Create(options).Paragraph(p => p.Runs(runs)).ToBytes();
        }
    }
}
