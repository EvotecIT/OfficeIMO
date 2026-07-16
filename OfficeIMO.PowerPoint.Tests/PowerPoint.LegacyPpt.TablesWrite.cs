using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptTableWriteTests {
        [Fact]
        public void NativeWriter_ExplicitlyConvertsStyledTableToStaticPngPicture() {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointTable table = source
                    .AddSlide(P.SlideLayoutValues.Blank)
                    .AddTable(3, 3, PowerPointTableStylePreset.Default,
                        PowerPointUnits.FromPoints(24D),
                        PowerPointUnits.FromPoints(36D),
                        PowerPointUnits.FromPoints(300D),
                        PowerPointUnits.FromPoints(144D));
                table.GetCell(0, 0).Text = "Region";
                table.GetCell(0, 1).Text = "Revenue";
                table.GetCell(0, 2).Text = "Plan";
                table.GetCell(1, 0).Text = "North";
                table.GetCell(1, 1).Text = "120";
                table.GetCell(1, 2).Text = "110";
                table.GetCell(2, 0).Text = "Total";
                table.GetCell(2, 1).Text = "230";
                table.GetCell(2, 1).FillColor = "DBEAFE";
                table.MergeCells(2, 1, 2, 2);
                table.GetCell(2, 1).HorizontalAlignment =
                    A.TextAlignmentTypeValues.Center;

                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.False(preflight.CanWrite);
                LegacyPptWriteFinding finding = Assert.Single(
                    preflight.Findings, candidate =>
                        candidate.Code == "PPT-WRITE-TABLE-CONVERTED");
                Assert.Equal(LegacyPptFeature.Tables, finding.Feature);
                Assert.Throws<NotSupportedException>(() => source.ToBytes(
                    PowerPointFileFormat.Ppt));
                bytes = source.ToBytes(PowerPointFileFormat.Ppt,
                    new PowerPointSaveOptions {
                        LossPolicy = PowerPointConversionLossPolicy.Allow
                    });
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            OfficeArtBlipStoreEntry image = Assert.Single(
                legacy.BlipStoreEntries);
            Assert.Equal("image/png", image.ContentType);
            Assert.True(OfficePngReader.TryDecode(image.ImageBytes,
                out OfficeRasterImage? raster));
            Assert.NotNull(raster);
            Assert.True(raster!.Width >= 300);
            Assert.True(raster.Height >= 144);
            LegacyPptShape picture = Assert.Single(
                Assert.Single(legacy.Slides).Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Picture);
            Assert.Equal(new LegacyPptBounds(192, 288, 2400, 1152),
                picture.Bounds);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(
                input);
            Assert.Empty(projected.Slides[0].Tables);
            Assert.Single(projected.Slides[0].Pictures);
            Assert.Empty(projected.ValidateDocument());
            Assert.True(projected.AnalyzeLegacyPptWrite().CanWrite);
            Assert.Equal(bytes, projected.ToBytes(PowerPointFileFormat.Ppt));
        }
    }
}
