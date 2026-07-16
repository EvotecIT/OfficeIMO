using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptPictureWriteTests {
        [Fact]
        public void NativeWriter_AuthorsDeduplicatesAndProjectsPngPictures() {
            byte[] image = OfficePngWriter.Encode(new OfficeRasterImage(
                4, 3, OfficeColor.FromRgb(37, 99, 235)));
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(P.SlideLayoutValues.Blank);
                using (var first = new MemoryStream(image, writable: false)) {
                    PowerPointPicture picture = slide.AddPicture(first,
                        ImagePartType.Png, 158750, 317500, 635000, 476250);
                    picture.Crop(10D, 20D, 5D, 15D);
                }
                using (var second = new MemoryStream(image, writable: false)) {
                    slide.AddPicture(second, ImagePartType.Png,
                        952500, 317500, 635000, 476250);
                }

                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            OfficeArtBlipStoreEntry entry = Assert.Single(
                legacy.BlipStoreEntries);
            Assert.Equal(OfficeArtBlipStorage.Delayed, entry.Storage);
            Assert.Equal(OfficeArtBlipType.Png, entry.RecordInstanceBlipType);
            Assert.Equal(2U, entry.ReferenceCount);
            Assert.Equal(0U, entry.DelayedStreamOffset);
            Assert.Equal("image/png", entry.ContentType);
            Assert.Equal(image, entry.ImageBytes);
            byte[] picturesStream = legacy.Package.CopyCompoundStreams()[
                "Pictures"];
            Assert.Equal(entry.SizeBytes, checked((uint)picturesStream.Length));
            Assert.Equal(0xF01E, picturesStream[2]
                | picturesStream[3] << 8);
            LegacyPptShape[] pictures = Assert.Single(legacy.Slides).Shapes
                .Where(shape => shape.Kind == LegacyPptShapeKind.Picture)
                .OrderBy(shape => shape.Bounds.Left)
                .ToArray();
            Assert.Equal(2, pictures.Length);
            Assert.All(pictures, picture => Assert.Equal(1,
                picture.PictureStoreIndex));
            Assert.Equal(new LegacyPptBounds(100, 200, 400, 300),
                pictures[0].Bounds);
            Assert.Equal(6554, pictures[0].PictureProperties.CropFromLeftRaw);
            Assert.Equal(13107, pictures[0].PictureProperties.CropFromTopRaw);
            Assert.Equal(3277, pictures[0].PictureProperties.CropFromRightRaw);
            Assert.Equal(9830, pictures[0].PictureProperties.CropFromBottomRaw);
            Assert.DoesNotContain(legacy.Diagnostics, diagnostic =>
                diagnostic.Code.StartsWith("PPT-PICTURE-",
                    StringComparison.Ordinal));

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(
                input);
            PowerPointPicture[] projectedPictures = projected.Slides[0].Pictures
                .OrderBy(picture => picture.Left)
                .ToArray();
            Assert.Equal(2, projectedPictures.Length);
            Assert.All(projectedPictures, picture => Assert.Equal(image,
                picture.GetImageBytes()));
            Assert.Equal(0.1D, projectedPictures[0].CropLeftRatio, 4);
            Assert.Equal(0.2D, projectedPictures[0].CropTopRatio, 4);
            Assert.Equal(0.05D, projectedPictures[0].CropRightRatio, 4);
            Assert.Equal(0.15D, projectedPictures[0].CropBottomRatio, 4);
            Assert.Empty(projected.ValidateDocument());
            Assert.True(projected.AnalyzeLegacyPptWrite().CanWrite);
            Assert.Equal(bytes, projected.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void NativeWriter_BlocksUnsupportedRasterFormatAndPictureEffects() {
            using PowerPointPresentation gifPresentation = PowerPointPresentation
                .Create();
            using (var gif = new MemoryStream(new byte[] {
                       0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00,
                       0x01, 0x00
                   }, writable: false)) {
                gifPresentation.AddSlide(P.SlideLayoutValues.Blank)
                    .AddPicture(gif, ImagePartType.Gif);
            }
            LegacyPptWritePreflightReport gifPreflight = gifPresentation
                .AnalyzeLegacyPptWrite();
            Assert.False(gifPreflight.CanWrite);
            Assert.Contains(gifPreflight.Findings, finding =>
                finding.Code == "PPT-WRITE-PICTURE"
                && finding.Feature == LegacyPptFeature.RasterPictures);

            byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(
                2, 2, OfficeColor.CornflowerBlue));
            using PowerPointPresentation effectPresentation = PowerPointPresentation
                .Create();
            using (var stream = new MemoryStream(png, writable: false)) {
                PowerPointPicture picture = effectPresentation
                    .AddSlide(P.SlideLayoutValues.Blank)
                    .AddPicture(stream, ImagePartType.Png);
                P.Picture element = Assert.IsType<P.Picture>(picture.Element);
                element.BlipFill!.Blip!.Append(new A.Grayscale());
            }
            LegacyPptWritePreflightReport effectPreflight = effectPresentation
                .AnalyzeLegacyPptWrite();
            Assert.False(effectPreflight.CanWrite);
            Assert.Contains(effectPreflight.Findings, finding =>
                finding.Code == "PPT-WRITE-PICTURE"
                && finding.Description.Contains("effects",
                    StringComparison.OrdinalIgnoreCase));
        }
    }
}
