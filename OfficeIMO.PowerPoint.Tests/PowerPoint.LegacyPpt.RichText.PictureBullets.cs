using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        private static string PictureBulletFixturePath => Path.Combine(
            AppContext.BaseDirectory, "Documents", "LegacyPptCorpus",
            "apache-poi-testdata", "bug61881.ppt");


        [Fact]
        public void FreshPictureBullet_IsWrittenAsPpt9CollectionAndReference() {
            byte[] image = OfficePngWriter.Encode(new OfficeRasterImage(
                4, 4, OfficeColor.FromRgb(20, 120, 220)));
            byte[] binary;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                PowerPointTextBox textBox = slide.AddTextBox(
                    "Picture bullet");
                ImagePart imagePart = slide.SlidePart.AddImagePart(
                    DocumentFormat.OpenXml.Packaging.ImagePartType.Png);
                using (var stream = new MemoryStream(image,
                           writable: false)) {
                    imagePart.FeedData(stream);
                }
                string relationshipId = slide.SlidePart.GetIdOfPart(
                    imagePart);
                A.Paragraph paragraph = Assert.Single(
                    Assert.IsType<P.Shape>(textBox.Element).TextBody!
                        .Elements<A.Paragraph>());
                paragraph.ParagraphProperties = new A.ParagraphProperties(
                    new A.PictureBullet(
                        new A.Blip { Embed = relationshipId }));

                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                binary = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(
                binary);
            Assert.Empty(legacy.BlipStoreEntries);
            LegacyPptPictureBullet pictureBullet = Assert.Single(
                legacy.PictureBullets);
            Assert.Equal((ushort)0, pictureBullet.Index);
            Assert.Equal("image/png", pictureBullet.ContentType);
            Assert.Equal(image, pictureBullet.ImageBytes);
            LegacyPptParagraphRun run = Assert.Single(Assert.Single(
                    Assert.Single(legacy.Slides).Shapes).TextBody
                .ParagraphRuns);
            Assert.Equal((ushort)0, run.BulletPictureReference);
            Assert.Same(pictureBullet, run.PictureBullet);
            Assert.True(run.HasBullet);

            InvalidDataException budgetException = Assert.Throws<
                InvalidDataException>(() => LegacyPptPresentation.Load(
                    binary, new LegacyPptImportOptions {
                        MaxDecodedStorageBytes = image.Length - 1
                    }));
            Assert.Contains("aggregate decoded embedded-storage",
                budgetException.Message,
                StringComparison.OrdinalIgnoreCase);

            using var binaryStream = new MemoryStream(binary,
                writable: false);
            using PowerPointPresentation projected =
                PowerPointPresentation.Load(binaryStream);
            Assert.Single(projected.Slides[0].SlidePart.Slide!
                .Descendants<A.PictureBullet>());
        }

        [Fact]
        public void ApachePoiPictureBullet_IsDecodedAndProjectedNatively() {
            LegacyPptPresentation source = LegacyPptPresentation.Load(
                PictureBulletFixturePath);
            LegacyPptPictureBullet pictureBullet = Assert.Single(
                source.PictureBullets);
            Assert.Equal((ushort)0, pictureBullet.Index);
            Assert.Equal((byte)0x06, pictureBullet.PreferredBlipType);
            Assert.Equal("image/png", pictureBullet.ContentType);
            Assert.Equal(225, pictureBullet.ImageBytes.Length);
            Assert.False(pictureBullet.IsPayloadTruncated);
            Assert.True(pictureBullet.HasImportableImage);
            Assert.Equal(1, source.CreateImportReport()
                .PictureBulletCount);
            LegacyPptParagraphRun paragraph = Assert.Single(source.Slides
                .SelectMany(slide => slide.Shapes)
                .SelectMany(shape => shape.TextBody.ParagraphRuns),
                run => run.BulletPictureReference == 0);
            Assert.Same(pictureBullet, paragraph.PictureBullet);
            Assert.False(paragraph.HasUnprojectedFormatting);

            using PowerPointPresentation projected =
                PowerPointPresentation.LoadLegacyPpt(PictureBulletFixturePath);
            var projectedBullets = projected.Slides
                .SelectMany(slide => slide.SlidePart.Slide!
                    .Descendants<A.PictureBullet>()
                    .Select(bullet => (slide.SlidePart, Bullet: bullet)))
                .ToArray();
            Assert.Equal(2, projectedBullets.Length);
            var projectedBullet = projectedBullets[0];
            SlidePart slidePart = projectedBullet.SlidePart;
            A.PictureBullet nativeBullet = projectedBullet.Bullet;
            A.Blip blip = Assert.Single(nativeBullet.Elements<A.Blip>());
            ImagePart imagePart = Assert.IsType<ImagePart>(
                slidePart.GetPartById(blip.Embed!.Value!));
            Assert.Equal("image/png", imagePart.ContentType);
            using Stream stream = imagePart.GetStream(FileMode.Open,
                FileAccess.Read);
            using var bytes = new MemoryStream();
            stream.CopyTo(bytes);
            Assert.Equal(pictureBullet.ImageBytes, bytes.ToArray());
            Assert.Single(projectedBullets.Select(item => item.Bullet
                .GetFirstChild<A.Blip>()!.Embed!.Value).Distinct());
        }

        [Fact]
        public void ImportedPictureBullet_AddChangeAndRemove_AppendPreservingRecords() {
            byte[] sourceBytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                source.AddSlide(P.SlideLayoutValues.Blank)
                    .AddTextBox("Editable picture bullet");
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);
            byte[] addedImage = OfficePngWriter.Encode(
                new OfficeRasterImage(4, 4,
                    OfficeColor.FromRgb(30, 150, 90)));
            byte[] addedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported =
                   PowerPointPresentation.Load(input)) {
                SlidePart slidePart = imported.Slides[0].SlidePart;
                ImagePart imagePart = slidePart.AddImagePart(
                    DocumentFormat.OpenXml.Packaging.ImagePartType.Png);
                using (var imageStream = new MemoryStream(addedImage,
                           writable: false)) {
                    imagePart.FeedData(imageStream);
                }
                A.Paragraph paragraph = Assert.Single(slidePart.Slide!
                    .Descendants<A.Paragraph>());
                paragraph.ParagraphProperties = new A.ParagraphProperties(
                    new A.PictureBullet(new A.Blip {
                        Embed = slidePart.GetIdOfPart(imagePart)
                    }));

                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                addedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation added = LegacyPptPresentation.Load(
                addedBytes);
            Assert.Equal(addedImage,
                Assert.Single(added.PictureBullets).ImageBytes);
            Assert.True(Assert.Single(Assert.Single(added.Slides).Shapes)
                .TextBody.ParagraphRuns.Single().HasBullet);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                added.Package.UserEdits.Count);
            Assert.True(added.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            byte[] changedImage = OfficePngWriter.Encode(
                new OfficeRasterImage(5, 3,
                    OfficeColor.FromRgb(180, 60, 40)));
            byte[] changedBytes;
            using (var input = new MemoryStream(addedBytes,
                       writable: false))
            using (PowerPointPresentation imported =
                   PowerPointPresentation.Load(input)) {
                SlidePart slidePart = imported.Slides[0].SlidePart;
                A.Blip blip = Assert.Single(slidePart.Slide!
                    .Descendants<A.PictureBullet>()).GetFirstChild<A.Blip>()!;
                ImagePart imagePart = Assert.IsType<ImagePart>(
                    slidePart.GetPartById(blip.Embed!.Value!));
                using (var imageStream = new MemoryStream(changedImage,
                           writable: false)) {
                    imagePart.FeedData(imageStream);
                }

                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                changedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation changed = LegacyPptPresentation.Load(
                changedBytes);
            Assert.Equal(changedImage,
                Assert.Single(changed.PictureBullets).ImageBytes);
            Assert.Equal(added.Package.UserEdits.Count + 1,
                changed.Package.UserEdits.Count);
            Assert.True(changed.Package.DocumentStream.AsSpan(0,
                    added.Package.DocumentStream.Length)
                .SequenceEqual(added.Package.DocumentStream));

            byte[] removedBytes;
            using (var input = new MemoryStream(changedBytes,
                       writable: false))
            using (PowerPointPresentation imported =
                   PowerPointPresentation.Load(input)) {
                A.PictureBullet bullet = Assert.Single(imported.Slides[0]
                    .SlidePart.Slide!.Descendants<A.PictureBullet>());
                A.ParagraphProperties properties = Assert.IsType<
                    A.ParagraphProperties>(bullet.Parent);
                bullet.Remove();
                properties.Append(new A.NoBullet());

                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                removedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation removed = LegacyPptPresentation.Load(
                removedBytes);
            Assert.Empty(removed.PictureBullets);
            Assert.False(Assert.Single(Assert.Single(removed.Slides).Shapes)
                .TextBody.ParagraphRuns.Single().HasBullet);
            Assert.Equal(changed.Package.UserEdits.Count + 1,
                removed.Package.UserEdits.Count);
            Assert.True(removed.Package.DocumentStream.AsSpan(0,
                    changed.Package.DocumentStream.Length)
                .SequenceEqual(changed.Package.DocumentStream));
        }

    }
}
