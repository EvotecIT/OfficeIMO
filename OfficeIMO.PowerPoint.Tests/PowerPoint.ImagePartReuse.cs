using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Svg = DocumentFormat.OpenXml.Office2019.Drawing.SVG;

namespace OfficeIMO.Tests {
    public class PowerPointImagePartReuseTests {
        private static byte[] Asset(string name) =>
            File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Images", name));

        [Fact]
        public void IdenticalPicturesShareOneImagePart() {
            byte[] logo = Asset("EvotecLogo.png");
            byte[] photo = Asset("Kulek.jpg");
            byte[] saved;
            using (PowerPointDocument presentation = PowerPointDocument.Create()) {
                for (int index = 0; index < 3; index++) {
                    using var stream = new MemoryStream(logo, writable: false);
                    presentation.AddSlide().AddImage(stream, OfficeImageFormat.Png);
                }
                using (var stream = new MemoryStream(photo, writable: false)) {
                    presentation.Slides[1].AddImage(stream, OfficeImageFormat.Jpeg);
                }
                saved = presentation.ToBytes();
            }

            using PresentationDocument package = Open(saved);
            Assert.Equal(new[] { logo, logo, photo, logo }, PictureBytes(package));
            Assert.Equal(2, package.GetAllParts().OfType<ImagePart>().Count());
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }

        [Fact]
        public void PictureAddedAfterItsSharedPartWasDeletedGetsANewPart() {
            byte[] logo = Asset("EvotecLogo.png");
            byte[] saved;
            using (PowerPointDocument presentation = PowerPointDocument.Create()) {
                AddLogo(presentation.AddSlide(), logo);
                AddLogo(presentation.AddSlide(), logo);
                presentation.RemoveSlide(0);
                // The remaining slide still references the part, so this picture reuses it.
                AddLogo(presentation.AddSlide(), logo);
                presentation.RemoveSlide(0);
                presentation.RemoveSlide(0);
                // Removing every slide deleted the part; the next picture must not point at it.
                AddLogo(presentation.AddSlide(), logo);
                saved = presentation.ToBytes();
            }

            using PresentationDocument package = Open(saved);
            Assert.Equal(new[] { logo }, PictureBytes(package));
            Assert.Single(package.GetAllParts().OfType<ImagePart>());
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }

        [Fact]
        public void RepeatedSvgPictureReusesItsFallbackAndSource() {
            byte[] svg = Asset("Sample.svg");
            byte[] saved;
            using (PowerPointDocument presentation = PowerPointDocument.Create()) {
                for (int index = 0; index < 2; index++) {
                    using var stream = new MemoryStream(svg, writable: false);
                    presentation.AddSlide().AddImage(stream, OfficeImageFormat.Svg);
                }
                saved = presentation.ToBytes();
            }

            using PresentationDocument package = Open(saved);
            ImagePart[] parts = package.GetAllParts().OfType<ImagePart>().ToArray();
            Assert.Equal(2, parts.Length);
            Assert.Single(parts, part => part.ContentType == "image/png");
            Assert.Single(parts, part => part.ContentType == "image/svg+xml");
            foreach (SlidePart slide in SlideParts(package)) {
                A.Blip blip = slide.Slide!.Descendants<A.Blip>().Single();
                Assert.Equal("image/png", slide.GetPartById(blip.Embed!.Value!).ContentType);
                string svgId = blip.Descendants<Svg.SVGBlip>().Single().Embed!.Value!;
                Assert.Equal(svg, ReadPart(slide.GetPartById(svgId)));
            }
        }

        private static void AddLogo(PowerPointSlide slide, byte[] logo) {
            using var stream = new MemoryStream(logo, writable: false);
            slide.AddImage(stream, OfficeImageFormat.Png);
        }

        private static PresentationDocument Open(byte[] bytes) =>
            PresentationDocument.Open(new MemoryStream(bytes, writable: false), false);

        private static IEnumerable<SlidePart> SlideParts(PresentationDocument package) {
            PresentationPart presentation = package.PresentationPart!;
            return presentation.Presentation!.SlideIdList!.Elements<P.SlideId>()
                .Select(id => (SlidePart)presentation.GetPartById(id.RelationshipId!.Value!));
        }

        private static List<byte[]> PictureBytes(PresentationDocument package) =>
            SlideParts(package)
                .SelectMany(slide => slide.Slide!.Descendants<A.Blip>()
                    .Select(blip => ReadPart(slide.GetPartById(blip.Embed!.Value!))))
                .ToList();

        private static byte[] ReadPart(OpenXmlPart part) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            return copy.ToArray();
        }
    }
}
