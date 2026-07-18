using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.Tests.Pdf;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptBackgroundPreservationTests {
        [Fact]
        public void NativeWriter_ResolvesBackgroundStyleFromSlideThemeOverride() {
            byte[] binary;
            using (PowerPointPresentation presentation =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = presentation.AddSlide(
                    P.SlideLayoutValues.Blank);
                A.ThemeElements masterElements = slide.SlidePart
                    .SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme!
                    .ThemeElements!;
                A.FormatScheme format = (A.FormatScheme)masterElements
                    .FormatScheme!.CloneNode(true);
                A.BackgroundFillStyleList fills = format
                    .GetFirstChild<A.BackgroundFillStyleList>()!;
                fills.RemoveAllChildren();
                fills.Append(new A.SolidFill(
                    new A.RgbColorModelHex { Val = "12AB34" }));
                ThemeOverridePart overridePart = slide.SlidePart
                    .AddNewPart<ThemeOverridePart>();
                overridePart.ThemeOverride = new A.ThemeOverride(
                    masterElements.ColorScheme!.CloneNode(true),
                    masterElements.FontScheme!.CloneNode(true), format);
                slide.SlidePart.Slide!.CommonSlideData!.Background =
                    new P.Background(new P.BackgroundStyleReference(
                        new A.SchemeColor {
                            Val = A.SchemeColorValues.PhColor
                        }) { Index = 1001U });

                LegacyPptWritePreflightReport preflight = presentation
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                binary = presentation.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptBackground background = Assert.IsType<
                LegacyPptBackground>(Assert.Single(
                    LegacyPptPresentation.Load(binary).Slides).Background);
            Assert.Equal(LegacyPptBackgroundKind.Solid, background.Kind);
            Assert.Equal("12AB34", background.ForegroundColor);
        }

        [Fact]
        public void NativeWriter_ReportsOutOfRangeBackgroundStyleIndex() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide(
                P.SlideLayoutValues.Blank);
            slide.SlidePart.Slide!.CommonSlideData!.Background =
                new P.Background(new P.BackgroundStyleReference {
                    Index = uint.MaxValue
                });

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-BACKGROUND"
                && finding.Description.Contains("cannot be resolved",
                    StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void NativeWriter_WritesDeduplicatedPictureBackgrounds() {
            byte[] imageBytes = PdfPngTestImages.CreateRgbPng(37, 99, 235);
            byte[] binary;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                for (int index = 0; index < 2; index++) {
                    PowerPointSlide slide = created.AddSlide(
                        P.SlideLayoutValues.Blank);
                    using var image = new MemoryStream(imageBytes,
                        writable: false);
                    slide.SetBackgroundImage(image,
                        OfficeIMO.PowerPoint.ImagePartType.Png);
                }

                LegacyPptWritePreflightReport preflight = created
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                binary = created.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation neutral = LegacyPptPresentation.Load(binary);
            Assert.Equal(2, neutral.Slides.Count);
            Assert.All(neutral.Slides, slide => {
                LegacyPptBackground background = Assert.IsType<
                    LegacyPptBackground>(slide.Background);
                Assert.Equal(LegacyPptBackgroundKind.Picture,
                    background.Kind);
                Assert.Equal(1, background.PictureStoreIndex);
                Assert.Equal(imageBytes, background.Picture!.ImageBytes);
            });
            Assert.Equal(2U, Assert.Single(neutral.BlipStoreEntries)
                .ReferenceCount);

            using var input = new MemoryStream(binary, writable: false);
            using PowerPointPresentation projected =
                PowerPointPresentation.Load(input);
            Assert.All(projected.Slides, slide => {
                PowerPointSlideBackground background = slide.GetBackground();
                Assert.Equal(PowerPointSlideBackgroundKind.Image,
                    background.Kind);
                Assert.Equal("image/png", background.ImageContentType);
                Assert.Equal(imageBytes, background.ImageBytes);
            });
            Assert.Empty(projected.ValidateDocument());
            Assert.Equal(binary,
                projected.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void NativeWriter_BlocksNonRepresentablePictureBackgroundPlacementAndEffects() {
            byte[] imageBytes = PdfPngTestImages.CreateRgbPng(37, 99, 235);
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide tiled = AddPictureBackgroundSlide(presentation,
                imageBytes);
            A.BlipFill tiledFill = tiled.SlidePart.Slide!.CommonSlideData!
                .Background!.BackgroundProperties!
                .GetFirstChild<A.BlipFill>()!;
            tiledFill.RemoveAllChildren<A.Stretch>();
            tiledFill.Append(new A.Tile());

            PowerPointSlide cropped = AddPictureBackgroundSlide(presentation,
                imageBytes);
            cropped.SlidePart.Slide!.CommonSlideData!.Background!
                .BackgroundProperties!.GetFirstChild<A.BlipFill>()!
                .SourceRectangle = new A.SourceRectangle { Left = 10000 };

            PowerPointSlide effected = AddPictureBackgroundSlide(presentation,
                imageBytes);
            effected.SlidePart.Slide!.CommonSlideData!.Background!
                .BackgroundProperties!.GetFirstChild<A.BlipFill>()!.Blip!
                .Append(new A.Grayscale());

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-BACKGROUND"
                && finding.Description.Contains("tiled placement",
                    StringComparison.Ordinal));
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-BACKGROUND"
                && finding.Description.Contains("source cropping",
                    StringComparison.Ordinal));
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-BACKGROUND"
                && finding.Description.Contains("image effects",
                    StringComparison.Ordinal));
        }

        private static PowerPointSlide AddPictureBackgroundSlide(
            PowerPointPresentation presentation, byte[] imageBytes) {
            PowerPointSlide slide = presentation.AddSlide(
                P.SlideLayoutValues.Blank);
            using var image = new MemoryStream(imageBytes, writable: false);
            slide.SetBackgroundImage(image,
                OfficeIMO.PowerPoint.ImagePartType.Png);
            return slide;
        }

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
        public void ImportedSlidePictureBackgroundEdit_AppendsPreservingBlipStoreEdit() {
            byte[] imageBytes = PdfPngTestImages.CreateRgbPng(233, 113, 50);
            byte[] sourceBytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                created.AddSlide(P.SlideLayoutValues.Blank).BackgroundColor =
                    "112233";
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);

            using var input = new MemoryStream(sourceBytes, writable: false);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            using (var image = new MemoryStream(imageBytes,
                       writable: false)) {
                imported.Slides[0].SetBackgroundImage(image,
                    OfficeIMO.PowerPoint.ImagePartType.Png);
            }

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                savedBytes);
            LegacyPptBackground background = Assert.IsType<
                LegacyPptBackground>(Assert.Single(saved.Slides).Background);
            Assert.Equal(LegacyPptBackgroundKind.Picture, background.Kind);
            Assert.Equal(1, background.PictureStoreIndex);
            Assert.Equal(imageBytes, background.Picture!.ImageBytes);
            Assert.Equal(1U, Assert.Single(saved.BlipStoreEntries)
                .ReferenceCount);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);

            using var reopenedInput = new MemoryStream(savedBytes,
                writable: false);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            Assert.Equal(imageBytes,
                Assert.Single(reopened.Slides).GetBackground().ImageBytes);
            Assert.Empty(reopened.ValidateDocument());
            Assert.Equal(savedBytes,
                reopened.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ImportedPictureBackgroundReplacementAndRemoval_BalanceBlipReferences() {
            byte[] originalImage = PdfPngTestImages.CreateRgbPng(20, 80, 140);
            byte[] replacementImage = PdfPngTestImages.CreateRgbPng(140, 80, 20);
            byte[] sourceBytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                AddPictureBackgroundSlide(created, originalImage);
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);
            byte[] originalPictures = Assert.IsType<byte[]>(
                original.Package.PicturesStream);

            using (var replacementInput = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation replacement =
                   PowerPointPresentation.Load(replacementInput)) {
                using var image = new MemoryStream(replacementImage,
                    writable: false);
                replacement.Slides[0].SetBackgroundImage(image,
                    OfficeIMO.PowerPoint.ImagePartType.Png);
                LegacyPptWritePreflightReport preflight = replacement
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));

                byte[] replacedBytes = replacement.ToBytes(
                    PowerPointFileFormat.Ppt);
                LegacyPptPresentation replaced = LegacyPptPresentation.Load(
                    replacedBytes);
                Assert.True(replaced.Package.DocumentStream.AsSpan(0,
                        original.Package.DocumentStream.Length)
                    .SequenceEqual(original.Package.DocumentStream));
                Assert.True(replaced.Package.PicturesStream!.AsSpan(0,
                        originalPictures.Length)
                    .SequenceEqual(originalPictures));
                Assert.Collection(replaced.BlipStoreEntries,
                    entry => {
                        Assert.Equal(0U, entry.ReferenceCount);
                        Assert.Equal(originalImage, entry.ImageBytes);
                    },
                    entry => {
                        Assert.Equal(1U, entry.ReferenceCount);
                        Assert.Equal(replacementImage, entry.ImageBytes);
                    });
                LegacyPptBackground background = Assert.IsType<
                    LegacyPptBackground>(Assert.Single(replaced.Slides)
                    .Background);
                Assert.Equal(2, background.PictureStoreIndex);
                Assert.Equal(replacementImage,
                    background.Picture!.ImageBytes);
            }

            using var removalInput = new MemoryStream(sourceBytes,
                writable: false);
            using PowerPointPresentation removal = PowerPointPresentation.Load(
                removalInput);
            removal.Slides[0].BackgroundColor = "445566";
            LegacyPptWritePreflightReport removalPreflight = removal
                .AnalyzeLegacyPptWrite();
            Assert.True(removalPreflight.CanWrite,
                string.Join(Environment.NewLine,
                    removalPreflight.Findings));
            LegacyPptPresentation removed = LegacyPptPresentation.Load(
                removal.ToBytes(PowerPointFileFormat.Ppt));
            Assert.Equal(originalPictures, removed.Package.PicturesStream);
            Assert.Equal(0U, Assert.Single(removed.BlipStoreEntries)
                .ReferenceCount);
            LegacyPptBackground solid = Assert.IsType<LegacyPptBackground>(
                Assert.Single(removed.Slides).Background);
            Assert.Equal(LegacyPptBackgroundKind.Solid, solid.Kind);
            Assert.Equal("445566", solid.ForegroundColor);
        }

        [Fact]
        public void ImportedInPlaceBackgroundImagePartReplacement_IsNotSilentlyDiscarded() {
            byte[] originalImage = PdfPngTestImages.CreateRgbPng(15, 45, 75);
            byte[] replacementImage = PdfPngTestImages.CreateRgbPng(75, 45, 15);
            byte[] sourceBytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                AddPictureBackgroundSlide(created, originalImage);
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }

            using var input = new MemoryStream(sourceBytes, writable: false);
            using PowerPointPresentation imported = PowerPointPresentation.Load(
                input);
            SlidePart slidePart = imported.Slides[0].SlidePart;
            string relationshipId = slidePart.Slide!.CommonSlideData!
                .Background!.BackgroundProperties!
                .GetFirstChild<A.BlipFill>()!.Blip!.Embed!.Value!;
            ImagePart imagePart = Assert.IsType<ImagePart>(
                slidePart.GetPartById(relationshipId));
            using (var replacement = new MemoryStream(replacementImage,
                       writable: false)) {
                imagePart.FeedData(replacement);
            }

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                imported.ToBytes(PowerPointFileFormat.Ppt));

            Assert.Collection(saved.BlipStoreEntries,
                entry => Assert.Equal(0U, entry.ReferenceCount),
                entry => {
                    Assert.Equal(1U, entry.ReferenceCount);
                    Assert.Equal(replacementImage, entry.ImageBytes);
                });
            Assert.Equal(replacementImage, Assert.IsType<LegacyPptBackground>(
                Assert.Single(saved.Slides).Background).Picture!.ImageBytes);
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
        public void ImportedLayoutPictureBackgroundEdit_MaterializesAndDeduplicatesBlip() {
            byte[] imageBytes = PdfPngTestImages.CreateRgbPng(121, 51, 201);
            byte[] sourceBytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                created.AddSlide(P.SlideLayoutValues.Blank);
                created.AddSlide(P.SlideLayoutValues.Blank);
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);

            using var input = new MemoryStream(sourceBytes, writable: false);
            using PowerPointPresentation imported = PowerPointPresentation.Load(
                input);
            SlideLayoutPart layoutPart = imported.Slides[0].SlidePart
                .SlideLayoutPart!;
            ImagePart imagePart = layoutPart.AddNewPart<ImagePart>("image/png");
            using (var image = new MemoryStream(imageBytes, writable: false)) {
                imagePart.FeedData(image);
            }
            layoutPart.SlideLayout!.CommonSlideData!.Background =
                new P.Background(new P.BackgroundProperties(new A.BlipFill(
                    new A.Blip {
                        Embed = layoutPart.GetIdOfPart(imagePart)
                    },
                    new A.Stretch(new A.FillRectangle()))));

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);

            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
            Assert.Equal(2U, Assert.Single(saved.BlipStoreEntries)
                .ReferenceCount);
            Assert.All(saved.Slides, slide => {
                Assert.False(slide.FollowsMasterBackground);
                LegacyPptBackground background = Assert.IsType<
                    LegacyPptBackground>(slide.Background);
                Assert.Equal(LegacyPptBackgroundKind.Picture,
                    background.Kind);
                Assert.Equal(imageBytes, background.Picture!.ImageBytes);
            });
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
