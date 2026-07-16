using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptChartWriteTests {
        [Fact]
        public void NativeWriter_ExplicitlyConvertsChartToStaticPngPicture() {
            var data = new OfficeChartData(new[] { "Q1", "Q2", "Q3" },
                new[] {
                    new OfficeChartSeries("Revenue",
                        new[] { 12D, 19D, 27D }, xValues: null,
                        color: OfficeColor.CornflowerBlue),
                    new OfficeChartSeries("Plan",
                        new[] { 15D, 18D, 24D }, xValues: null,
                        color: OfficeColor.Orange)
                });
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                source.AddSlide(P.SlideLayoutValues.Blank)
                    .AddChartPoints(OfficeChartKind.ColumnClustered, data,
                        30D, 20D, 240D, 150D)
                    .SetTitle("Quarterly revenue");

                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.False(preflight.CanWrite);
                LegacyPptWriteFinding finding = Assert.Single(
                    preflight.Findings, candidate =>
                        candidate.Code == "PPT-WRITE-CHART-CONVERTED");
                Assert.Equal(LegacyPptFeature.Charts, finding.Feature);
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
            Assert.True(raster!.Width >= 240);
            Assert.True(raster.Height >= 150);
            LegacyPptShape picture = Assert.Single(
                Assert.Single(legacy.Slides).Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Picture);
            Assert.Equal(new LegacyPptBounds(240, 160, 1920, 1200),
                picture.Bounds);
            Assert.Equal(1, picture.PictureStoreIndex);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(
                input);
            Assert.Empty(projected.Slides[0].Charts);
            PowerPointPicture projectedPicture = Assert.Single(
                projected.Slides[0].Pictures);
            Assert.Equal(image.ImageBytes, projectedPicture.GetImageBytes());
            Assert.Empty(projected.ValidateDocument());
            Assert.True(projected.AnalyzeLegacyPptWrite().CanWrite);
            Assert.Equal(bytes, projected.ToBytes(PowerPointFileFormat.Ppt));
        }
    }
}
