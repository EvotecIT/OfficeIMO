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
        [InlineData(PowerPointSmartArtType.BasicList,
            OfficeDiagramKind.List)]
        [InlineData(PowerPointSmartArtType.BasicMatrix,
            OfficeDiagramKind.Matrix)]
        [InlineData(PowerPointSmartArtType.BasicPyramid,
            OfficeDiagramKind.Pyramid)]
        [InlineData(PowerPointSmartArtType.BasicRelationship,
            OfficeDiagramKind.Relationship)]
        public void NativeWriter_ExplicitlyConvertsSmartArtToStaticPngPicture(
            PowerPointSmartArtType sourceKind, OfficeDiagramKind expectedKind) {
            byte[] bytes;
            using (PowerPointDocument source = PowerPointDocument.Create()) {
                PowerPointSmartArt smartArt = source
                    .AddSlide(PowerPointSlideLayoutType.Blank)
                    .AddSmartArt(sourceKind,
                        new[] { "Discover", "Build", "Validate", "Ship" },
                        OfficeIMO.OfficeLength.Points(20D),
                        OfficeIMO.OfficeLength.Points(30D),
                        OfficeIMO.OfficeLength.Points(260D),
                        OfficeIMO.OfficeLength.Points(160D));
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
                        candidate.Code == "ppt-write-smartart-converted");
                Assert.Equal(LegacyPptFeature.SmartArt, finding.Feature);
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
            LegacyPptShape picture = Assert.Single(
                Assert.Single(legacy.Slides).Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Picture);
            Assert.Equal(new LegacyPptBounds(160, 240, 2080, 1280),
                picture.Bounds);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointDocument projected = PowerPointDocument.Load(
                input);
            Assert.Empty(projected.Slides[0].SmartArts);
            Assert.Single(projected.Slides[0].Pictures);
            Assert.Empty(projected.ValidateDocument());
            Assert.True(projected.AnalyzeLegacyPptWrite().CanWrite);
            Assert.Equal(bytes, projected.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt }));
        }
    }
}
