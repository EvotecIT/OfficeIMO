using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.PowerPoint.Tests {
    public class PowerPointLazyShapeLoadingTests {
        [Fact]
        public void LoadedSlidesMaterializeShapesOnlyWhenAccessedOrEdited() {
            byte[] source;
            using (var created = PowerPointPresentation.Create(new MemoryStream())) {
                created.AddSlide().AddTextBoxPoints("First", 20, 20, 120, 30);
                created.AddSlide().AddTextBoxPoints("Second", 20, 20, 120, 30);
                PowerPointSlide third = created.AddSlide();
                third.AddTextBoxPoints("Third", 20, 20, 120, 30);
                third.Hide();
                source = created.ToBytes();
            }

            using var presentation = PowerPointPresentation.Load(
                new MemoryStream(source, writable: false));
            Assert.All(presentation.Slides,
                slide => Assert.False(slide.SlidePart.IsRootElementLoaded));

            Assert.Single(presentation.Slides[0].Shapes);
            Assert.True(presentation.Slides[0].SlidePart.IsRootElementLoaded);
            Assert.False(presentation.Slides[1].SlidePart.IsRootElementLoaded);
            Assert.False(presentation.Slides[2].SlidePart.IsRootElementLoaded);

            presentation.Slides[1].AddTextBoxPoints("Reviewed", 20, 60, 120, 30);
            Assert.Equal(2, presentation.Slides[1].Shapes.Count);
            Assert.False(presentation.Slides[2].SlidePart.IsRootElementLoaded);

            using var output = new MemoryStream();
            presentation.Save(output);
            Assert.Equal("1", presentation.ApplicationProperties.HiddenSlides);
            Assert.False(presentation.Slides[2].SlidePart.IsRootElementLoaded);
            output.Position = 0;
            using var reopened = PowerPointPresentation.Load(output);
            Assert.Equal(new[] { 1, 2, 1 },
                reopened.Slides.Select(slide => slide.Shapes.Count));
            Assert.True(reopened.Slides[2].Hidden);
            Assert.Empty(new OpenXmlValidator().Validate(reopened.OpenXmlDocument));
        }

        [Fact]
        public void SaveStillNormalizesLegacySlideVisibilityMarkup() {
            byte[] source;
            using (var created = PowerPointPresentation.Create(new MemoryStream())) {
                created.AddSlide().AddTextBoxPoints("Legacy hidden", 20, 20, 160, 30);
                source = created.ToBytes();
            }

            using var legacy = new MemoryStream(source.Length + 8192);
            legacy.Write(source, 0, source.Length);
            legacy.Position = 0;
            using (var package = PresentationDocument.Open(legacy, true,
                       new OpenSettings { AutoSave = false })) {
                PresentationPart presentationPart = package.PresentationPart!;
                SlideId slideId = presentationPart.Presentation.SlideIdList!
                    .Elements<SlideId>().Single();
                string relationshipId = slideId.RelationshipId!.Value!;
                SlidePart slidePart = (SlidePart)presentationPart.GetPartById(relationshipId);
                slidePart.Slide!.Show = null;
                slidePart.Slide.Save();
                slideId.SetAttribute(new OpenXmlAttribute("show", string.Empty, "0"));
                presentationPart.Presentation.Save();
            }

            legacy.Position = 0;
            using var presentation = PowerPointPresentation.Load(legacy);
            Assert.False(presentation.Slides[0].SlidePart.IsRootElementLoaded);
            using var output = new MemoryStream();
            presentation.Save(output);
            Assert.True(presentation.Slides[0].SlidePart.IsRootElementLoaded);

            output.Position = 0;
            using var reopened = PresentationDocument.Open(output, false);
            PresentationPart reopenedPart = reopened.PresentationPart!;
            SlideId reopenedId = reopenedPart.Presentation.SlideIdList!
                .Elements<SlideId>().Single();
            Assert.DoesNotContain(reopenedId.GetAttributes(), attribute =>
                attribute.LocalName == "show" && string.IsNullOrEmpty(attribute.NamespaceUri));
            Assert.False(reopenedPart.SlideParts.Single().Slide!.Show!.Value);
        }
    }
}
