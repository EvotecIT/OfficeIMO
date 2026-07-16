using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptBackgroundPreservationTests {
        [Fact]
        public void ImportedSlideBackgroundEdit_AppendsPreservingIncrementalRecord() {
            byte[] sourceBytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = created.AddSlide(
                    P.SlideLayoutValues.Blank);
                slide.BackgroundColor = "112233";
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            imported.Slides[0].BackgroundColor = "445566";

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            LegacyPptBackground background = Assert.IsType<LegacyPptBackground>(
                Assert.Single(saved.Slides).Background);

            Assert.Equal(LegacyPptBackgroundKind.Solid, background.Kind);
            Assert.Equal("445566", background.ForegroundColor);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);

            using var reopenedInput = new MemoryStream(savedBytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            PowerPointSlideBackground projected = reopened.Slides[0]
                .GetBackground();
            Assert.Equal(PowerPointSlideBackgroundKind.SolidColor,
                projected.Kind);
            Assert.Equal("445566", projected.Color);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedLayoutBackgroundEdit_MaterializesIntoAffectedSlides() {
            byte[] sourceBytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                created.AddSlide(P.SlideLayoutValues.Blank);
                created.AddSlide(P.SlideLayoutValues.Blank);
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            SlideLayoutPart layoutPart = imported.Slides[0].SlidePart
                .SlideLayoutPart!;
            Assert.All(imported.Slides, slide => Assert.Same(layoutPart,
                slide.SlidePart.SlideLayoutPart));
            Assert.All(imported.LegacyPptProjectionMap!.Slides,
                slide => Assert.False(slide.HasExplicitBackground));
            Assert.True(imported.LegacyPptProjectionMap
                .IsEditableProjectedLayoutBackgroundPart(
                    layoutPart.Uri.ToString()));
            layoutPart.SlideLayout!.CommonSlideData!.Background =
                new P.Background(new P.BackgroundProperties(
                    new A.SolidFill(
                        new A.RgbColorModelHex { Val = "315274" })));

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);

            Assert.Equal(2, saved.Slides.Count);
            Assert.All(saved.Slides, slide => {
                Assert.False(slide.FollowsMasterBackground);
                Assert.Equal("315274", Assert.IsType<LegacyPptBackground>(
                    slide.Background).ForegroundColor);
            });
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            using var reopenedInput = new MemoryStream(savedBytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            Assert.All(reopened.Slides, slide => {
                PowerPointSlideBackground background = slide.GetBackground();
                Assert.Equal(PowerPointSlideBackgroundKind.SolidColor,
                    background.Kind);
                Assert.Equal("315274", background.Color);
            });
            Assert.Empty(reopened.ValidateDocument());
        }
    }
}
