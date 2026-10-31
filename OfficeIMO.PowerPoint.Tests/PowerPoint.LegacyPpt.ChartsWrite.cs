using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptChartWriteTests {
        [Theory]
        [InlineData(PowerPointClassicAnimationBuildType.ChartBySeries)]
        [InlineData(PowerPointClassicAnimationBuildType.ChartByCategory)]
        [InlineData(PowerPointClassicAnimationBuildType.ChartByElementInSeries)]
        [InlineData(PowerPointClassicAnimationBuildType.ChartByElementInCategory)]
        public void ClassicChartAnimationBuildsSurviveStaticLegacyChartProjection(
            PowerPointClassicAnimationBuildType buildType) {
            var data = new OfficeChartData(new[] { "Q1", "Q2", "Q3" },
                new[] {
                    new OfficeSeries("Revenue", new[] { 12D, 19D, 27D }),
                    new OfficeSeries("Plan", new[] { 15D, 18D, 24D })
                });
            byte[] bytes;
            using (PowerPointDocument source = PowerPointDocument.Create()) {
                PowerPointSlide slide = source.AddSlide(PowerPointSlideLayoutType.Blank);
                PowerPointChart chart = slide.AddChart(OfficeChartType.ColumnClustered, data, new OfficeIMO.PowerPoint.PowerPointChartOptions { Left = global::OfficeIMO.OfficeLength.Points(30D), Top = global::OfficeIMO.OfficeLength.Points(20D), Width = global::OfficeIMO.OfficeLength.Points(240D), Height = global::OfficeIMO.OfficeLength.Points(150D) });
                slide.AddClassicAnimation(chart, PowerPointClassicAnimationEffect.Fade,
                    new PowerPointClassicAnimationOptions { BuildType = buildType });
                bytes = source.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt,
                        LossPolicy = OfficeConversionLossPolicy.Allow
                    });
            }

            LegacyPptShape legacyShape = Assert.Single(
                Assert.Single(LegacyPptDocument.Load(bytes).Slides).Shapes);
            Assert.Equal(LegacyPptShapeKind.Picture, legacyShape.Kind);
            Assert.Equal((LegacyPptAnimationBuildType)buildType,
                legacyShape.Animation!.BuildType);

            using PowerPointDocument projected = PowerPointDocument.Load(
                new MemoryStream(bytes, writable: false));
            PowerPointClassicAnimation animation = Assert.Single(
                projected.Slides[0].ClassicAnimations);
            Assert.Equal(buildType, animation.BuildType);
            Assert.IsType<PowerPointPicture>(Assert.Single(projected.Slides[0].Shapes));
        }

        [Fact]
        public void NativeWriter_ExplicitlyConvertsChartToStaticPngPicture() {
            var data = new OfficeChartData(new[] { "Q1", "Q2", "Q3" },
                new[] {
                    new OfficeSeries("Revenue",
                        new[] { 12D, 19D, 27D }, xValues: null,
                        color: OfficeColor.CornflowerBlue),
                    new OfficeSeries("Plan",
                        new[] { 15D, 18D, 24D }, xValues: null,
                        color: OfficeColor.Orange)
                });
            byte[] bytes;
            using (PowerPointDocument source = PowerPointDocument.Create()) {
                source.AddSlide(PowerPointSlideLayoutType.Blank)
                    .AddChart(OfficeChartType.ColumnClustered, data, new OfficeIMO.PowerPoint.PowerPointChartOptions { Left = global::OfficeIMO.OfficeLength.Points(30D), Top = global::OfficeIMO.OfficeLength.Points(20D), Width = global::OfficeIMO.OfficeLength.Points(240D), Height = global::OfficeIMO.OfficeLength.Points(150D) })
                    .SetTitle("Quarterly revenue");

                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.False(preflight.CanWrite);
                LegacyPptWriteFinding finding = Assert.Single(
                    preflight.Findings, candidate =>
                        candidate.Code == "ppt-write-chart-converted");
                Assert.Equal(LegacyPptFeature.Charts, finding.Feature);
                Assert.Throws<NotSupportedException>(() => source.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt }));

                bytes = source.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt,
                        LossPolicy = OfficeConversionLossPolicy.Allow
                    });
            }

            LegacyPptDocument legacy = LegacyPptDocument.Load(bytes);
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
            using PowerPointDocument projected = PowerPointDocument.Load(
                input);
            Assert.Empty(projected.Slides[0].Charts);
            PowerPointPicture projectedPicture = Assert.Single(
                projected.Slides[0].Pictures);
            Assert.Equal(image.ImageBytes, projectedPicture.GetImageBytes());
            Assert.Empty(projected.ValidateDocument());
            Assert.True(projected.AnalyzeLegacyPptWrite().CanWrite);
            Assert.Equal(bytes, projected.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt }));
        }
    }
}
