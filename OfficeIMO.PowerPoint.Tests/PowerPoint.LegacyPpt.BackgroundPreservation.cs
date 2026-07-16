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

        [Fact]
        public void NativeWriter_RoundTripsPathGradientsAndLinearOpacityRamp() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide center = presentation.AddSlide(
                P.SlideLayoutValues.Blank);
            center.SlidePart.Slide!.CommonSlideData!.Background =
                CreatePathGradientBackground(A.PathShadeValues.Circle,
                    ("112233", 0, 25000),
                    ("778899", 50000, 50000),
                    ("DDEEFF", 100000, 75000));
            PowerPointSlide shape = presentation.AddSlide(
                P.SlideLayoutValues.Blank);
            shape.SlidePart.Slide!.CommonSlideData!.Background =
                CreatePathGradientBackground(A.PathShadeValues.Shape,
                    ("102030", 0, 100000),
                    ("A0B0C0", 100000, 100000));

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);

            LegacyPptBackground centerBackground = Assert.IsType<
                LegacyPptBackground>(binary.Slides[0].Background);
            Assert.Equal(LegacyPptBackgroundKind.CenterGradient,
                centerBackground.Kind);
            Assert.Equal(3, centerBackground.GradientStops.Count);
            Assert.InRange(centerBackground.ForegroundOpacity!.Value,
                0.250D, 0.252D);
            Assert.InRange(centerBackground.BackgroundOpacity!.Value,
                0.748D, 0.750D);
            Assert.Equal(LegacyPptBackgroundKind.ShapeGradient,
                Assert.IsType<LegacyPptBackground>(binary.Slides[1].Background)
                    .Kind);

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(stream);
            A.GradientFill projectedCenter = Assert.IsType<A.GradientFill>(
                reopened.Slides[0].SlidePart.Slide!.CommonSlideData!.Background!
                    .BackgroundProperties!.GetFirstChild<A.GradientFill>());
            Assert.Equal(A.PathShadeValues.Circle, projectedCenter
                .GetFirstChild<A.PathGradientFill>()!.Path!.Value);
            int[] projectedAlpha = projectedCenter
                .GetFirstChild<A.GradientStopList>()!
                .Elements<A.GradientStop>()
                .Select(stop => stop.GetFirstChild<A.RgbColorModelHex>()!
                    .GetFirstChild<A.Alpha>()!.Val!.Value)
                .ToArray();
            Assert.Collection(projectedAlpha,
                alpha => Assert.InRange(alpha, 25090, 25110),
                alpha => Assert.InRange(alpha, 49990, 50010),
                alpha => Assert.InRange(alpha, 74890, 74910));
            A.GradientFill projectedShape = Assert.IsType<A.GradientFill>(
                reopened.Slides[1].SlidePart.Slide!.CommonSlideData!.Background!
                    .BackgroundProperties!.GetFirstChild<A.GradientFill>());
            Assert.Equal(A.PathShadeValues.Shape, projectedShape
                .GetFirstChild<A.PathGradientFill>()!.Path!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_BlocksNonLinearGradientOpacity() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide(
                P.SlideLayoutValues.Blank);
            slide.SlidePart.Slide!.CommonSlideData!.Background =
                CreatePathGradientBackground(A.PathShadeValues.Circle,
                    ("112233", 0, 20000),
                    ("778899", 50000, 90000),
                    ("DDEEFF", 100000, 40000));

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-BACKGROUND"
                && finding.Description.Contains("non-linear stop opacity",
                    StringComparison.Ordinal));
        }

        [Fact]
        public void ImportedSlidePathGradientEdit_AppendsPreservingRecord() {
            byte[] sourceBytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                created.AddSlide(P.SlideLayoutValues.Blank).BackgroundColor =
                    "112233";
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original =
                LegacyPptPresentation.Load(sourceBytes);

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            imported.Slides[0].SlidePart.Slide!.CommonSlideData!.Background =
                CreatePathGradientBackground(A.PathShadeValues.Circle,
                    ("224466", 0, 100000),
                    ("AACCEE", 100000, 100000));

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved =
                LegacyPptPresentation.Load(savedBytes);

            Assert.Equal(LegacyPptBackgroundKind.CenterGradient,
                Assert.IsType<LegacyPptBackground>(
                    Assert.Single(saved.Slides).Background).Kind);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
        }

        private static P.Background CreatePathGradientBackground(
            A.PathShadeValues path,
            params (string Color, int Position, int Alpha)[] stops) {
            var stopList = new A.GradientStopList();
            foreach ((string color, int position, int alpha) in stops) {
                stopList.Append(new A.GradientStop(
                    new A.RgbColorModelHex(
                        new A.Alpha { Val = alpha }) { Val = color }) {
                    Position = position
                });
            }
            return new P.Background(new P.BackgroundProperties(
                new A.GradientFill(stopList,
                    new A.PathGradientFill { Path = path }) {
                    RotateWithShape = true
                }));
        }
    }
}
