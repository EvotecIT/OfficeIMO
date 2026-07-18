using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptSmartArtWriteTests {
        [Theory]
        [InlineData(PowerPointSmartArtType.BasicProcess,
            OfficeDiagramKind.Process)]
        [InlineData(PowerPointSmartArtType.BasicHierarchy,
            OfficeDiagramKind.Hierarchy)]
        [InlineData(PowerPointSmartArtType.BasicCycle,
            OfficeDiagramKind.Cycle)]
        public void NativeWriter_ExplicitlyConvertsSmartArtToStaticPngPicture(
            PowerPointSmartArtType sourceKind, OfficeDiagramKind expectedKind) {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSmartArt smartArt = source
                    .AddSlide(P.SlideLayoutValues.Blank)
                    .AddSmartArt(sourceKind,
                        new[] { "Discover", "Build", "Validate", "Ship" },
                        PowerPointUnits.FromPoints(20D),
                        PowerPointUnits.FromPoints(30D),
                        PowerPointUnits.FromPoints(260D),
                        PowerPointUnits.FromPoints(160D));
                Assert.True(smartArt.TryGetOfficeDiagramSnapshot(
                    out OfficeDiagramSnapshot snapshot));
                Assert.Equal(expectedKind, snapshot.Kind);
                Assert.Equal(new[] { "Discover", "Build", "Validate", "Ship" },
                    snapshot.Nodes);

                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.False(preflight.CanWrite);
                LegacyPptWriteFinding finding = Assert.Single(
                    preflight.Findings, candidate =>
                        candidate.Code == "PPT-WRITE-SMARTART-CONVERTED");
                Assert.Equal(LegacyPptFeature.SmartArt, finding.Feature);
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
            LegacyPptShape picture = Assert.Single(
                Assert.Single(legacy.Slides).Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Picture);
            Assert.Equal(new LegacyPptBounds(160, 240, 2080, 1280),
                picture.Bounds);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(
                input);
            Assert.Empty(projected.Slides[0].SmartArts);
            Assert.Single(projected.Slides[0].Pictures);
            Assert.Empty(projected.ValidateDocument());
            Assert.True(projected.AnalyzeLegacyPptWrite().CanWrite);
            Assert.Equal(bytes, projected.ToBytes(PowerPointFileFormat.Ppt));
        }
    }
}
