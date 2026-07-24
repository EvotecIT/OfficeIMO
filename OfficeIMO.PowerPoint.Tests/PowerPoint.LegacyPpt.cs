using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.Drawing.Binary;
using System.Threading.Tasks;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        private static string FixturePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "BasicPowerPoint.ppt");

        private static string PictureFixturePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "PicturePowerPoint.ppt");

        private static string CroppedPictureFixturePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "CroppedPicturePowerPoint.ppt");

        private static string PictureEffectsFixturePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "PictureEffectsPowerPoint.ppt");

        [Fact]
        public void NeutralReader_DecodesRealBinaryPresentation() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(FixturePath);

            LegacyPptSlide slide = Assert.Single(legacy.Slides);
            Assert.Equal(7680, legacy.SlideWidth);
            Assert.Equal(4320, legacy.SlideHeight);
            Assert.Equal(11, legacy.Masters.Count);
            Assert.All(legacy.Masters, master => Assert.True(master.IsMainMaster));
            Assert.Equal(legacy.Masters[0].MasterId, slide.MasterId);
            Assert.Contains(legacy.Masters[0].Shapes,
                shape => shape.PlaceholderKind == LegacyPptPlaceholderKind.MasterTitle);
            LegacyPptColorScheme masterScheme = Assert.IsType<LegacyPptColorScheme>(
                legacy.Masters[0].ColorScheme);
            Assert.Equal("FFFFFF", masterScheme.Background);
            Assert.Equal("000000", masterScheme.Text);
            Assert.Equal("3333CC", masterScheme.Accent1);
            Assert.True(slide.FollowsMasterObjects);
            Assert.True(slide.FollowsMasterColorScheme);
            Assert.True(slide.FollowsMasterBackground);
            Assert.Equal(LegacyPptSlideLayoutType.TitleSlide, slide.Layout);
            Assert.Equal(new[] {
                LegacyPptPlaceholderKind.Title,
                LegacyPptPlaceholderKind.Subtitle,
                LegacyPptPlaceholderKind.None,
                LegacyPptPlaceholderKind.None,
                LegacyPptPlaceholderKind.None,
                LegacyPptPlaceholderKind.None,
                LegacyPptPlaceholderKind.None,
                LegacyPptPlaceholderKind.None
            }, slide.LayoutPlaceholderTypes);
            Assert.NotNull(slide.ColorScheme);
            Assert.Equal(3, slide.Shapes.Count(shape => shape.Kind == LegacyPptShapeKind.TextBox));
            LegacyPptShape title = Assert.Single(slide.Shapes,
                shape => shape.Text == "OfficeIMO PowerPoint Basics");
            LegacyPptPlaceholder titlePlaceholder = Assert.IsType<LegacyPptPlaceholder>(
                title.Placeholder);
            Assert.Equal(0, titlePlaceholder.Position);
            Assert.Equal(LegacyPptPlaceholderSize.Full, titlePlaceholder.Size);
            Assert.False(title.Style.FillEnabled);
            Assert.False(title.Style.LineEnabled);
            Assert.Equal("3465A4", title.LineColor);
            Assert.Contains(title.Style.Properties, property => property.PropertyName == "lineColor");
            Assert.Contains(slide.Shapes, shape => shape.PlaceholderKind == LegacyPptPlaceholderKind.Title);
            Assert.Equal(8, legacy.Masters[0].TextMasterStyles.Count);
            Assert.All(legacy.Masters[0].TextMasterStyles, style => Assert.False(style.IsTruncated));
            LegacyPptTextMasterStyle bodyStyle = Assert.Single(legacy.Masters[0].TextMasterStyles,
                style => style.TextType == LegacyPptTextType.Body);
            Assert.Equal(5, bodyStyle.Levels.Count);
            Assert.Equal((short)216, bodyStyle.Levels[0].ParagraphProperties.LeftMargin);
            Assert.Equal((short)18, bodyStyle.Levels[0].CharacterProperties.FontSizePoints);
            Assert.NotNull(bodyStyle.Levels[0].CharacterProperties.Typeface);
            LegacyPptTextRuler titleRuler = Assert.IsType<LegacyPptTextRuler>(title.TextBody.Ruler);
            Assert.Equal(LegacyPptTextType.Title, title.TextBody.TextType);
            Assert.Equal(12, titleRuler.TabStops.Count);
            Assert.Equal((short)576, titleRuler.TabStops[0].Position);
            Assert.Equal(LegacyPptTabAlignment.Left, titleRuler.TabStops[0].Alignment);
            Assert.False(title.TextBody.IsRulerTruncated);
            LegacyPptImportReport report = legacy.CreateImportReport();
            Assert.Equal(3, report.TextRulerCount);
            Assert.Equal(88, report.MasterTextStyleCount);
            Assert.Equal(440, report.MasterTextStyleLevelCount);
            Assert.Equal(1, report.DistinctSlideLayoutCount);
            Assert.True(report.PlaceholderShapeCount >= 4);
            Assert.True(report.HasConversionLoss);
            Assert.DoesNotContain(legacy.Diagnostics,
                diagnostic => diagnostic.Code == "PPT-TEXT-MASTER-STYLE-PRESERVE-ONLY"
                    || diagnostic.Code == "PPT-TEXT-MASTER-STYLE-TRUNCATED"
                    || diagnostic.Code == "PPT-TEXT-MASTER-STYLE-PARTIAL");
            Assert.NotEmpty(legacy.Package.UserEdits);
            Assert.NotEmpty(legacy.Package.PersistObjects);
            Assert.True(legacy.CreateImportReport().CompoundStreamCount >= 2);
        }

        [Fact]
        public void NeutralReader_DecodesDelayedPngPictureStoreAndFrameReference() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(PictureFixturePath);

            OfficeArtBlipStoreEntry entry = Assert.Single(legacy.BlipStoreEntries);
            Assert.Equal(OfficeArtBlipStorage.Delayed, entry.Storage);
            Assert.Equal(OfficeArtBlipType.Png, entry.RecordInstanceBlipType);
            Assert.Equal("OfficeArtBlipPNG", entry.BlipRecordTypeName);
            Assert.Equal("image/png", entry.ContentType);
            Assert.Equal(21283U, entry.BlipPayloadLength);
            Assert.Equal(21283, entry.BlipPayloadAvailableLength);
            Assert.Equal("E715CBA6CDBDDB873B50FB2B8ACBEC0CF72D4DD74468440B2B95B921FD99DD51",
                entry.BlipPayloadSha256);
            Assert.Equal(21266, entry.ImageBytes.Length);
            Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
                entry.ImageBytes.Take(8));

            LegacyPptShape picture = Assert.Single(Assert.Single(legacy.Slides).Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Picture);
            Assert.Equal(75, picture.OfficeArtShapeType);
            Assert.Equal(1, picture.PictureStoreIndex);
            Assert.Same(entry, picture.Picture);
            Assert.Equal(new LegacyPptBounds(576, 576, 1728, 403), picture.Bounds);
            LegacyPptImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.PictureShapeCount);
            Assert.Equal(1, report.BlipStoreEntryCount);
            Assert.Equal(1, report.ImportableBlipCount);
            Assert.DoesNotContain(legacy.Diagnostics, diagnostic =>
                diagnostic.Code == "PPT-PICTURE-BLIP-MISSING"
                || diagnostic.Code == "PPT-PICTURE-BLIP-TRUNCATED"
                || diagnostic.Code == "PPT-PICTURE-FORMAT-UNSUPPORTED");
        }

        [Fact]
        public void NormalLoad_ProjectsBinaryPictureIntoValidEditablePptxModel() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(PictureFixturePath);
            byte[] expected = Assert.Single(legacy.BlipStoreEntries).ImageBytes;

            using PowerPointPresentation presentation = PowerPointPresentation.Load(PictureFixturePath);

            PowerPointSlide slide = Assert.Single(presentation.Slides);
            PowerPointPicture picture = Assert.Single(slide.Pictures);
            Assert.Equal("image/png", picture.ContentType);
            Assert.Equal(expected, picture.GetImageBytes());
            Assert.Single(slide.SlidePart.ImageParts);
            Assert.Empty(presentation.ValidateDocument());

            using MemoryStream pptx = presentation.ToStream();
            using PowerPointPresentation reopened = PowerPointPresentation.Load(pptx);
            Assert.Equal(expected, Assert.Single(reopened.Slides[0].Pictures).GetImageBytes());
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NeutralReader_DecodesSignedPictureCropFractions() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(CroppedPictureFixturePath);
            LegacyPptShape[] pictures = Assert.Single(legacy.Slides).Shapes
                .Where(shape => shape.Kind == LegacyPptShapeKind.Picture)
                .OrderBy(shape => shape.Bounds.Left)
                .ToArray();

            Assert.Equal(2, pictures.Length);
            OfficeArtPictureProperties positive = pictures[0].PictureProperties;
            Assert.Equal(16379, positive.CropFromTopRaw);
            Assert.Equal(8189, positive.CropFromBottomRaw);
            Assert.Equal(8192, positive.CropFromLeftRaw);
            Assert.Equal(4093, positive.CropFromRightRaw);
            OfficeArtPictureProperties negative = pictures[1].PictureProperties;
            Assert.Equal(-4094, negative.CropFromTopRaw);
            Assert.Equal(-8192, negative.CropFromLeftRaw);
            Assert.True(negative.HasCrop);
        }

        [Fact]
        public void NeutralReader_DecodesPictureEffectProperties() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(PictureEffectsFixturePath);
            LegacyPptShape[] pictures = Assert.Single(legacy.Slides).Shapes
                .Where(shape => shape.Kind == LegacyPptShapeKind.Picture)
                .OrderBy(shape => shape.Bounds.Top)
                .ThenBy(shape => shape.Bounds.Left)
                .ToArray();

            Assert.Equal(6, pictures.Length);
            Assert.Equal(8175, pictures[0].PictureProperties.BrightnessRaw);
            Assert.Equal(45875, pictures[1].PictureProperties.ContrastRaw);
            Assert.Equal(109226, pictures[2].PictureProperties.ContrastRaw);
            Assert.True(pictures[3].PictureProperties.Grayscale);
            Assert.Null(pictures[3].PictureProperties.BiLevel);
            Assert.True(pictures[4].PictureProperties.Grayscale);
            Assert.True(pictures[4].PictureProperties.BiLevel);
            Assert.False(pictures[5].PictureProperties.HasPictureEffect);
        }

        [Fact]
        public void NormalLoad_ProjectsNativePictureEffectsAndPreservesBinaryExactly() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(
                PictureEffectsFixturePath);
            P.Picture[] pictures = Assert.Single(presentation.Slides).Pictures
                .OrderBy(picture => picture.Top)
                .ThenBy(picture => picture.Left)
                .Select(picture => Assert.IsType<P.Picture>(picture.Element))
                .ToArray();

            Assert.Equal(24948, pictures[0].BlipFill!.Blip!
                .GetFirstChild<A.LuminanceEffect>()!.Brightness!.Value);
            Assert.Equal(-30000, pictures[1].BlipFill!.Blip!
                .GetFirstChild<A.LuminanceEffect>()!.Contrast!.Value);
            Assert.Equal(40000, pictures[2].BlipFill!.Blip!
                .GetFirstChild<A.LuminanceEffect>()!.Contrast!.Value);
            Assert.NotNull(pictures[3].BlipFill!.Blip!.GetFirstChild<A.Grayscale>());
            Assert.Null(pictures[3].BlipFill.Blip.GetFirstChild<A.BiLevel>());
            Assert.Equal(50000, pictures[4].BlipFill!.Blip!
                .GetFirstChild<A.BiLevel>()!.Threshold!.Value);
            Assert.Null(pictures[4].BlipFill.Blip.GetFirstChild<A.Grayscale>());
            Assert.Empty(presentation.ValidateDocument());
            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            Assert.Equal(File.ReadAllBytes(PictureEffectsFixturePath),
                presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ImportedPictureEffectEdit_UsesIncrementalFoptRewrite() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                PictureEffectsFixturePath);
            byte[] picturesStream = original.Package.CopyCompoundStreams()[
                "Pictures"];
            using PowerPointPresentation presentation = PowerPointPresentation.Load(
                PictureEffectsFixturePath);
            PowerPointPicture picture = presentation.Slides[0].Pictures
                .OrderBy(item => item.Top)
                .ThenBy(item => item.Left)
                .First();
            picture.LuminanceBrightness = 10;

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptShape savedPicture = Assert.Single(saved.Slides).Shapes
                .Where(shape => shape.Kind == LegacyPptShapeKind.Picture)
                .OrderBy(shape => shape.Bounds.Top)
                .ThenBy(shape => shape.Bounds.Left)
                .First();
            Assert.Equal(3277, savedPicture.PictureProperties.BrightnessRaw);
            Assert.Equal(picturesStream,
                saved.Package.CopyCompoundStreams()["Pictures"]);
        }

        [Fact]
        public void NormalLoad_ProjectsSignedPictureCropAndPreservesBinaryExactly() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(
                CroppedPictureFixturePath);
            PowerPointPicture[] pictures = Assert.Single(presentation.Slides).Pictures
                .OrderBy(picture => picture.Left)
                .ToArray();

            Assert.Equal(2, pictures.Length);
            Assert.Equal(0.125D, pictures[0].CropLeftRatio, 5);
            Assert.Equal(0.24992D, pictures[0].CropTopRatio, 5);
            Assert.Equal(0.06245D, pictures[0].CropRightRatio, 5);
            Assert.Equal(0.12495D, pictures[0].CropBottomRatio, 5);
            P.Picture negative = Assert.IsType<P.Picture>(pictures[1].Element);
            Assert.Equal(-12500, negative.BlipFill!.SourceRectangle!.Left!.Value);
            Assert.Equal(-6247, negative.BlipFill.SourceRectangle.Top!.Value);
            Assert.Empty(presentation.ValidateDocument());
            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            Assert.Equal(File.ReadAllBytes(CroppedPictureFixturePath),
                presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ImportedPictureCropEdit_UsesIncrementalFoptRewrite() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                CroppedPictureFixturePath);
            byte[] picturesStream = original.Package.CopyCompoundStreams()[
                "Pictures"];
            using PowerPointPresentation presentation = PowerPointPresentation.Load(
                CroppedPictureFixturePath);
            PowerPointPicture picture = presentation.Slides[0].Pictures.OrderBy(item => item.Left).First();

            picture.Crop(10D, 20D, 5D, 10D);

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptShape savedPicture = Assert.Single(saved.Slides).Shapes
                .Where(shape => shape.Kind == LegacyPptShapeKind.Picture)
                .OrderBy(shape => shape.Bounds.Left)
                .First();
            Assert.Equal(6554, savedPicture.PictureProperties.CropFromLeftRaw);
            Assert.Equal(13107, savedPicture.PictureProperties.CropFromTopRaw);
            Assert.Equal(3277, savedPicture.PictureProperties.CropFromRightRaw);
            Assert.Equal(6554, savedPicture.PictureProperties.CropFromBottomRaw);
            Assert.Equal(picturesStream,
                saved.Package.CopyCompoundStreams()["Pictures"]);
        }

        [Fact]
        public void ImportedNegativeCropSurvivesPictureEffectEdit() {
            using PowerPointPresentation presentation = PowerPointPresentation
                .Load(CroppedPictureFixturePath);
            PowerPointPicture picture = presentation.Slides[0].Pictures
                .OrderBy(item => item.Left)
                .Last();
            picture.GrayScale = true;

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptShape savedPicture = Assert.Single(saved.Slides).Shapes
                .Where(shape => shape.Kind == LegacyPptShapeKind.Picture)
                .OrderBy(shape => shape.Bounds.Left)
                .Last();
            Assert.Equal(-8192,
                savedPicture.PictureProperties.CropFromLeftRaw);
            Assert.Equal(-4094,
                savedPicture.PictureProperties.CropFromTopRaw);
            Assert.True(savedPicture.PictureProperties.Grayscale);
        }

        [Fact]
        public void UnmodifiedPictureBinarySave_PreservesOriginalPackageExactly() {
            byte[] source = File.ReadAllBytes(PictureFixturePath);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(PictureFixturePath);

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            Assert.Equal(source, presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ImportedPictureGeometryEdit_UsesIncrementalAnchorRewriteAndPreservesImageStream() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(PictureFixturePath);
            LegacyPptShape sourcePicture = Assert.Single(original.Slides[0].Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Picture);
            byte[] picturesStream = original.Package.CopyCompoundStreams()["Pictures"];

            using PowerPointPresentation presentation = PowerPointPresentation.Load(PictureFixturePath);
            PowerPointPicture picture = Assert.Single(presentation.Slides[0].Pictures);
            picture.Left += 15875;
            picture.Width += 3175;

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptShape savedPicture = Assert.Single(saved.Slides[0].Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Picture);
            Assert.Equal(sourcePicture.Bounds.Left + 10, savedPicture.Bounds.Left);
            Assert.Equal(sourcePicture.Bounds.Width + 2, savedPicture.Bounds.Width);
            Assert.Equal(picturesStream, saved.Package.CopyCompoundStreams()["Pictures"]);
            Assert.Equal(original.BlipStoreEntries[0].ImageBytes, saved.BlipStoreEntries[0].ImageBytes);
        }

        [Fact]
        public void ImportedPictureImageReplacement_RemainsLossBlocked() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(PictureFixturePath);
            PowerPointPicture picture = Assert.Single(presentation.Slides[0].Pictures);
            using var replacement = new MemoryStream(picture.GetImageBytes(), writable: false);

            picture.UpdateImage(replacement, ImagePartType.Png);

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding => finding.Code == "PPT-WRITE-IMPORT-LOSS");
            Assert.Throws<NotSupportedException>(() => presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void UnmodifiedBinarySave_PreservesTheOriginalPackageExactly() {
            byte[] source = File.ReadAllBytes(FixturePath);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            byte[] saved = presentation.ToBytes(PowerPointFileFormat.Ppt);

            Assert.True(preflight.CanWrite);
            Assert.Equal(source, saved);
        }

        [Fact]
        public void BinaryPreservationFingerprint_DetectsCorePropertyChanges() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);
            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);

            presentation.BuiltinDocumentProperties.Title = "Changed title";

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings,
                finding => finding.Code == "PPT-WRITE-IMPORT-LOSS");
        }

        [Fact]
        public void ImportedBinaryGeometryEdit_AppendsIncrementalEditAndPreservesOpaqueStreams() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(FixturePath);
            LegacyPptShape originalTitle = original.Slides[0].Shapes.Single(shape =>
                shape.Text == "OfficeIMO PowerPoint Basics");
            IReadOnlyDictionary<string, byte[]> originalStreams = original.Package.CopyCompoundStreams();

            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);
            PowerPointTextBox title = presentation.Slides[0].TextBoxes.Single(textBox =>
                textBox.Text == "OfficeIMO PowerPoint Basics");
            title.Left += 15875;
            title.Width += 3175;

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite);
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);

            LegacyPptPresentation saved = LegacyPptPresentation.Load(bytes);
            LegacyPptShape savedTitle = saved.Slides[0].Shapes.Single(shape =>
                shape.Text == "OfficeIMO PowerPoint Basics");
            Assert.Equal(originalTitle.Bounds.Left + 10, savedTitle.Bounds.Left);
            Assert.Equal(originalTitle.Bounds.Width + 2, savedTitle.Bounds.Width);
            Assert.Equal(original.Package.UserEdits.Count + 1, saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0, original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            IReadOnlyDictionary<string, byte[]> savedStreams = saved.Package.CopyCompoundStreams();
            Assert.Equal(originalStreams.Keys.OrderBy(value => value), savedStreams.Keys.OrderBy(value => value));
            foreach (KeyValuePair<string, byte[]> stream in originalStreams) {
                if (stream.Key.Equals("PowerPoint Document", StringComparison.OrdinalIgnoreCase)
                    || stream.Key.Equals("Current User", StringComparison.OrdinalIgnoreCase)) continue;
                Assert.Equal(stream.Value, savedStreams[stream.Key]);
            }
        }

        [Fact]
        public void ImportedRichBinaryText_SameLengthEditPreservesFormattingRecords() {
            const string originalText = "OfficeIMO PowerPoint Basics";
            const string replacementText = "OfficeIMO BinaryDeck Basics";
            Assert.Equal(originalText.Length, replacementText.Length);

            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);
            PowerPointTextBox title = presentation.Slides[0].TextBoxes.Single(textBox =>
                textBox.Text == originalText);
            title.Text = replacementText;

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);

            LegacyPptPresentation saved = LegacyPptPresentation.Load(bytes);
            Assert.Contains(saved.Slides[0].Shapes, shape => shape.Text == replacementText);
            Assert.DoesNotContain(saved.Diagnostics, diagnostic =>
                diagnostic.Code == "PPT-TEXT-MASTER-STYLE-PRESERVE-ONLY"
                    || diagnostic.Code == "PPT-TEXT-MASTER-STYLE-TRUNCATED"
                    || diagnostic.Code == "PPT-TEXT-MASTER-STYLE-PARTIAL");
        }

        [Fact]
        public void ImportedRichBinaryText_LengthChangingEditUsesIncrementalSave() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                FixturePath);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);
            PowerPointTextBox title = presentation.Slides[0].TextBoxes.Single(textBox =>
                textBox.Text == "OfficeIMO PowerPoint Basics");
            title.Text += "!";

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            Assert.Contains(saved.Slides[0].Shapes,
                shape => shape.Text == "OfficeIMO PowerPoint Basics!");
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void ImportedPlainBinaryText_ArbitraryLengthEditUsesIncrementalSave() {
            byte[] source;
            using (PowerPointPresentation created = PowerPointPresentation.Create()) {
                created.AddSlide(P.SlideLayoutValues.Blank)
                    .AddTextBox("Short text", 100000, 100000, 3000000, 600000);
                source = created.ToBytes(PowerPointFileFormat.Ppt);
            }

            using var input = new MemoryStream(source);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(input);
            presentation.Slides[0].TextBoxes.Single().Text = "A substantially longer binary text value";

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            byte[] savedBytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            Assert.Contains(saved.Slides[0].Shapes,
                shape => shape.Text == "A substantially longer binary text value");
            Assert.Equal(2, saved.Package.UserEdits.Count);
        }

        [Fact]
        public void NormalLoad_RoutesPptAndProjectsEditablePptxModel() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(FixturePath);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);

            Assert.Equal(PowerPointFileFormat.Ppt, presentation.SourceFormat);
            Assert.Equal(FixturePath, presentation.SourcePath);
            Assert.Single(presentation.LegacyPptProjectionMap!.Slides);
            Assert.Equal(3, presentation.LegacyPptProjectionMap.Slides[0].Shapes.Count);
            Assert.Equal(legacy.Slides[0].MasterId,
                presentation.LegacyPptProjectionMap.Slides[0].MasterId);
            Assert.Equal(legacy.Masters.Count(master => master.IsMainMaster),
                presentation.OpenXmlDocument.PresentationPart!.SlideMasterParts.Count());
            Assert.Empty(presentation.ValidateDocument());
            PowerPointSlide slide = Assert.Single(presentation.Slides);
            Assert.Equal("FFFFFF", presentation.GetThemeColor(PowerPointThemeColor.Light1));
            Assert.Equal("000000", presentation.GetThemeColor(PowerPointThemeColor.Dark1));
            Assert.Equal("3333CC", presentation.GetThemeColor(PowerPointThemeColor.Accent1));
            Assert.Null(slide.SlidePart.ThemeOverridePart);
            Assert.Equal(P.SlideLayoutValues.Title,
                slide.SlidePart.SlideLayoutPart!.SlideLayout!.Type!.Value);
            PowerPointLayoutPlaceholderInfo[] layoutPlaceholders = slide.GetLayoutPlaceholders()
                .OrderBy(placeholder => placeholder.PlaceholderIndex)
                .ToArray();
            Assert.Equal(2, layoutPlaceholders.Length);
            Assert.Equal(P.PlaceholderValues.Title, layoutPlaceholders[0].PlaceholderType);
            Assert.Equal(0U, layoutPlaceholders[0].PlaceholderIndex);
            Assert.Equal(P.PlaceholderValues.SubTitle, layoutPlaceholders[1].PlaceholderType);
            Assert.Equal(1U, layoutPlaceholders[1].PlaceholderIndex);
            PowerPointSlideBackground background = slide.GetBackground();
            Assert.Equal(PowerPointSlideBackgroundKind.SolidColor, background.Kind);
            Assert.Equal("FFFFFF", background.Color);
            var masterShapes = slide.SlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster!
                .CommonSlideData!.ShapeTree!.Elements<DocumentFormat.OpenXml.Presentation.Shape>().ToArray();
            Assert.Contains(masterShapes, shape => shape.TextBody?.InnerText == "Click to edit the title text format");
            Assert.Equal(3, slide.TextBoxes.Count());
            Assert.Contains(slide.TextBoxes, textBox => textBox.Text == "OfficeIMO PowerPoint Basics");
            PowerPointTextBox projectedTitle = Assert.Single(slide.TextBoxes,
                textBox => textBox.Text == "OfficeIMO PowerPoint Basics");
            Assert.Equal(0U, projectedTitle.ShapePlaceholderIndex);
            Assert.Equal(P.PlaceholderSizeValues.Full, projectedTitle.ShapePlaceholderSize);
            DocumentFormat.OpenXml.Presentation.Shape titleShape = slide.SlidePart.Slide!.CommonSlideData!
                .ShapeTree!.Elements<DocumentFormat.OpenXml.Presentation.Shape>()
                .Single(shape => shape.TextBody?.InnerText == "OfficeIMO PowerPoint Basics");
            Assert.NotNull(titleShape.ShapeProperties!.GetFirstChild<A.NoFill>());
            Assert.NotNull(titleShape.ShapeProperties.GetFirstChild<A.Outline>()?.GetFirstChild<A.NoFill>());
            slide.AddTextBox("Edited after binary import");

            using var pptx = presentation.ToStream();
            using PowerPointPresentation reopened = PowerPointPresentation.Load(pptx);
            Assert.Contains(reopened.Slides[0].TextBoxes, textBox => textBox.Text == "Edited after binary import");
        }

        [Fact]
        public void OfficeArtStyleProjection_MapsSupportedSolidAndLineProperties() {
            OfficeArtShapeStyle style = OfficeArtShapeStyle.Decode(new[] {
                new OfficeArtProperty(0, 0x0181, 0x00332211U),
                new OfficeArtProperty(1, 0x0182, 0x00008000U),
                new OfficeArtProperty(2, 0x01BF, 0x00100010U),
                new OfficeArtProperty(3, 0x01C0, 0x00665544U),
                new OfficeArtProperty(4, 0x01C1, 0x0000C000U),
                new OfficeArtProperty(5, 0x01CB, 12700U),
                new OfficeArtProperty(6, 0x01CE, 3U),
                new OfficeArtProperty(7, 0x01D0, 1U),
                new OfficeArtProperty(8, 0x01D2, 0U),
                new OfficeArtProperty(9, 0x01D3, 2U),
                new OfficeArtProperty(10, 0x01D6, 1U),
                new OfficeArtProperty(11, 0x01D7, 2U),
                new OfficeArtProperty(12, 0x01FF, 0x00080008U)
            });
            var source = new LegacyPptShape(LegacyPptShapeKind.Rectangle, 1, 1, 0,
                new LegacyPptBounds(0, 0, 100, 100), string.Empty, placeholder: null,
                style, "112233", "445566");
            var properties = new DocumentFormat.OpenXml.Presentation.ShapeProperties(
                new A.Transform2D(),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle });

            PowerPointPresentation.ApplyLegacyShapeStyle(properties, source);

            A.RgbColorModelHex fill = properties.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!;
            Assert.Equal("112233", fill.Val!.Value);
            Assert.Equal(50000, fill.GetFirstChild<A.Alpha>()!.Val!.Value);
            A.Outline outline = properties.GetFirstChild<A.Outline>()!;
            Assert.Equal(12700, outline.Width!.Value);
            Assert.Equal(A.LineCapValues.Flat, outline.CapType!.Value);
            Assert.Equal("445566", outline.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            Assert.Equal(75000, outline.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!
                .GetFirstChild<A.Alpha>()!.Val!.Value);
            Assert.Equal(A.PresetLineDashValues.SystemDashDot,
                outline.GetFirstChild<A.PresetDash>()!.Val!.Value);
            Assert.NotNull(outline.GetFirstChild<A.Miter>());
            A.HeadEnd head = outline.GetFirstChild<A.HeadEnd>()!;
            Assert.Equal(A.LineEndValues.Triangle, head.Type!.Value);
            Assert.Equal(A.LineEndWidthValues.Small, head.Width!.Value);
            Assert.Equal(A.LineEndLengthValues.Large, head.Length!.Value);
        }

        [Theory]
        [InlineData(".pot", PowerPointFileFormat.Pot)]
        [InlineData(".pps", PowerPointFileFormat.Pps)]
        public void NormalLoad_PreservesLegacyExtensionSemantics(string extension, PowerPointFileFormat format) {
            string copy = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + extension);
            try {
                File.Copy(FixturePath, copy);
                using PowerPointPresentation presentation = PowerPointPresentation.Load(copy);
                Assert.Equal(format, presentation.SourceFormat);
                Assert.Single(presentation.Slides);
            } finally {
                if (File.Exists(copy)) File.Delete(copy);
            }
        }

        [Fact]
        public void NormalLoad_RejectsOtherLegacyOfficeFamiliesClearly() {
            string doc = Path.Combine(AppContext.BaseDirectory, "Documents", "LegacyDocCorpus", "ComSimpleParagraphs.doc");
            string xls = Path.Combine(AppContext.BaseDirectory, "Documents", "LegacyXlsCorpus",
                "openpreserve-format-corpus", "valid.xls");

            InvalidDataException wordError = Assert.Throws<InvalidDataException>(() => PowerPointPresentation.Load(doc));
            InvalidDataException excelError = Assert.Throws<InvalidDataException>(() => PowerPointPresentation.Load(xls));

            Assert.Contains("legacy Word", wordError.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("legacy Excel", excelError.Message, StringComparison.OrdinalIgnoreCase);

            var boundedOptions = new PowerPointLoadOptions {
                LegacyPptImportOptions = new LegacyPptImportOptions {
                    MaxInputBytes = 1
                }
            };
            InvalidDataException boundedWord = Assert.Throws<
                InvalidDataException>(() => PowerPointPresentation.Load(doc,
                    boundedOptions));
            InvalidDataException boundedExcel = Assert.Throws<
                InvalidDataException>(() => PowerPointPresentation.Load(xls,
                    boundedOptions));
            Assert.Contains("exceeds", boundedWord.Message,
                StringComparison.OrdinalIgnoreCase);
            Assert.Contains("exceeds", boundedExcel.Message,
                StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void NormalLoad_PrefersZipContentOverMisleadingPptExtension() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".ppt");
            try {
                using (PowerPointPresentation created = PowerPointPresentation.Create()) {
                    created.AddSlide().AddTextBox("Actually Open XML");
                    File.WriteAllBytes(path, created.ToBytes());
                }

                using PowerPointPresentation loaded = PowerPointPresentation.Load(path);
                Assert.Equal(PowerPointFileFormat.Pptx, loaded.SourceFormat);
                Assert.Contains(loaded.Slides[0].TextBoxes, textBox => textBox.Text == "Actually Open XML");
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void NativeWriter_RoundTripsTextAndBasicShapesAcrossTwoSlides() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide first = presentation.AddSlide();
            first.AddTitle("Binary title", 540000, 360000, 7200000, 700000);
            first.AddTextBox("Line one\nLine two", 540000, 1200000, 6000000, 1800000);
            first.AddRectangle(600000, 3300000, 1200000, 700000);
            first.AddEllipse(2100000, 3300000, 1200000, 700000);
            first.AddShape(A.ShapeTypeValues.Line, 3600000, 3300000, 1600000, 400000);
            presentation.AddSlide().AddTextBox("Second slide", 700000, 700000, 5000000, 1000000);

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite);
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);

            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);
            Assert.Equal(2, binary.Slides.Count);
            Assert.Contains(binary.Slides[0].Shapes, shape => shape.Text == "Binary title");
            Assert.Contains(binary.Slides[0].Shapes, shape => shape.Kind == LegacyPptShapeKind.Rectangle);
            Assert.Contains(binary.Slides[0].Shapes, shape => shape.Kind == LegacyPptShapeKind.Ellipse);
            Assert.Contains(binary.Slides[0].Shapes, shape => shape.Kind == LegacyPptShapeKind.Line);

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation projected = PowerPointPresentation.Load(stream);
            Assert.Equal(PowerPointFileFormat.Ppt, projected.SourceFormat);
            Assert.Equal(2, projected.Slides.Count);
            Assert.Contains(projected.Slides[1].TextBoxes, textBox => textBox.Text == "Second slide");
        }

        [Fact]
        public void NativeWriter_RoundTripsHiddenSlideState() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.AddSlide().AddTextBox("Visible");
            PowerPointSlide hidden = presentation.AddSlide();
            hidden.AddTextBox("Hidden");
            hidden.Hidden = true;

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            Assert.False(legacy.Slides[0].Hidden);
            Assert.True(legacy.Slides[1].Hidden);
            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened = PowerPointPresentation.Load(stream);
            Assert.False(reopened.Slides[0].Hidden);
            Assert.True(reopened.Slides[1].Hidden);
        }

        [Fact]
        public void ImportedBinaryHiddenState_TogglesThroughIncrementalEdits() {
            byte[] hiddenBytes;
            using (PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath)) {
                presentation.Slides[0].Hidden = true;
                Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
                hiddenBytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation hidden = LegacyPptPresentation.Load(hiddenBytes);
            Assert.True(hidden.Slides[0].Hidden);

            using var input = new MemoryStream(hiddenBytes);
            using PowerPointPresentation reopened = PowerPointPresentation.Load(input);
            reopened.Slides[0].Hidden = false;
            Assert.True(reopened.AnalyzeLegacyPptWrite().CanWrite);
            LegacyPptPresentation visible = LegacyPptPresentation.Load(
                reopened.ToBytes(PowerPointFileFormat.Ppt));
            Assert.False(visible.Slides[0].Hidden);
            Assert.Equal(hidden.Package.UserEdits.Count + 1, visible.Package.UserEdits.Count);
        }

        [Fact]
        public void ImportedBinarySlideOrder_RequiresExplicitFreshRewrite() {
            byte[] sourceBytes;
            using (PowerPointPresentation created = PowerPointPresentation.Create()) {
                created.AddSlide(P.SlideLayoutValues.Blank).AddTextBox("First");
                created.AddSlide(P.SlideLayoutValues.Blank).AddTextBox("Second");
                created.AddSlide(P.SlideLayoutValues.Blank).AddTextBox("Third");
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation source = LegacyPptPresentation.Load(sourceBytes);
            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(input);
            presentation.MoveSlide(0, 2);

            LegacyPptWritePreflightReport report = presentation
                .AnalyzeLegacyPptWrite();
            Assert.False(report.CanWrite);
            Assert.Contains(report.Findings, finding =>
                finding.Code == "PPT-WRITE-IMPORT-LOSS");
            Assert.Throws<NotSupportedException>(() =>
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt,
                    new PowerPointSaveOptions {
                        LossPolicy = PowerPointConversionLossPolicy.Allow
                    }));
            Assert.Equal(new[] { "Second", "Third", "First" }, saved.Slides.Select(slide =>
                slide.Shapes.Single(shape => shape.Kind == LegacyPptShapeKind.TextBox).Text));
            Assert.False(saved.Package.DocumentStream.AsSpan(0,
                    Math.Min(saved.Package.DocumentStream.Length,
                        source.Package.DocumentStream.Length))
                .SequenceEqual(source.Package.DocumentStream.AsSpan(0,
                    Math.Min(saved.Package.DocumentStream.Length,
                        source.Package.DocumentStream.Length))));
        }

        [Fact]
        public void ImportedBinarySlideDeletion_RequiresExplicitFreshRewrite() {
            byte[] sourceBytes;
            using (PowerPointPresentation created = PowerPointPresentation.Create()) {
                created.AddSlide(P.SlideLayoutValues.Blank).AddTextBox("Keep first");
                created.AddSlide(P.SlideLayoutValues.Blank).AddTextBox("Delete middle");
                created.AddSlide(P.SlideLayoutValues.Blank).AddTextBox("Keep last");
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(input);
            presentation.RemoveSlide(1);

            LegacyPptWritePreflightReport report = presentation
                .AnalyzeLegacyPptWrite();
            Assert.False(report.CanWrite);
            Assert.Contains(report.Findings, finding =>
                finding.Code == "PPT-WRITE-IMPORT-LOSS");
            Assert.Throws<NotSupportedException>(() =>
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt,
                    new PowerPointSaveOptions {
                        LossPolicy = PowerPointConversionLossPolicy.Allow
                    }));
            Assert.Equal(new[] { "Keep first", "Keep last" }, saved.Slides.Select(slide =>
                slide.Shapes.Single(shape => shape.Kind == LegacyPptShapeKind.TextBox).Text));
            Assert.False(saved.Slides.Any(slide => slide.Shapes.Any(shape =>
                shape.Text == "Delete middle")));
        }

        [Fact]
        public void ImportedBinarySlideAddition_RequiresExplicitFreshRewrite() {
            LegacyPptPresentation source = LegacyPptPresentation.Load(FixturePath);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);
            PowerPointSlide added = presentation.AddSlide();
            added.AddTextBox("Incremental new slide", 200000, 200000, 4000000, 700000);
            added.AddRectangle(300000, 1200000, 1200000, 600000);
            added.Hidden = true;

            LegacyPptWritePreflightReport report = presentation
                .AnalyzeLegacyPptWrite();
            Assert.False(report.CanWrite);
            Assert.Contains(report.Findings, finding =>
                finding.Code == "PPT-WRITE-IMPORT-LOSS");
            Assert.Throws<NotSupportedException>(() =>
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt,
                    new PowerPointSaveOptions {
                        LossPolicy = PowerPointConversionLossPolicy.Allow
                    }));

            Assert.Equal(2, saved.Slides.Count);
            Assert.Contains(saved.Slides[0].Shapes, shape => shape.Text == "OfficeIMO PowerPoint Basics");
            Assert.Contains(saved.Slides[1].Shapes, shape => shape.Text == "Incremental new slide");
            Assert.Contains(saved.Slides[1].Shapes, shape => shape.Kind == LegacyPptShapeKind.Rectangle);
            Assert.True(saved.Slides[1].Hidden);
            Assert.Equal(source.Slides[0].MasterId, saved.Slides[1].MasterId);
            Assert.Equal(LegacyPptSlideLayoutType.TitleSlide, saved.Slides[1].Layout);
            Assert.Equal(LegacyPptPlaceholderKind.Title,
                saved.Slides[1].LayoutPlaceholderTypes[0]);
            Assert.Equal(LegacyPptPlaceholderKind.Subtitle,
                saved.Slides[1].LayoutPlaceholderTypes[1]);
            Assert.False(saved.Package.DocumentStream.AsSpan(0,
                    Math.Min(saved.Package.DocumentStream.Length,
                        source.Package.DocumentStream.Length))
                .SequenceEqual(source.Package.DocumentStream.AsSpan(0,
                    Math.Min(saved.Package.DocumentStream.Length,
                        source.Package.DocumentStream.Length))));
        }

        [Fact]
        public void ImportedBinarySlideAddition_PreservesSelectedBinaryMaster() {
            LegacyPptPresentation source = LegacyPptPresentation.Load(FixturePath);
            LegacyPptMaster selectedMaster = source.Masters.Where(master => master.IsMainMaster).ElementAt(1);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);
            PowerPointSlide added = presentation.AddSlide(masterIndex: 1, layoutIndex: 0);
            added.AddTextBox("Uses second binary master", 200000, 200000, 4000000, 700000);

            Assert.Equal($"Binary Main Master {selectedMaster.MasterId:X8}",
                added.SlidePart.SlideLayoutPart!.SlideLayout!.CommonSlideData!.Name!.Value);
            Assert.True(presentation.LegacyPptProjectionMap!.TryGetMasterId(added,
                out uint projectedMasterId));
            Assert.Equal(selectedMaster.MasterId, projectedMasterId);

            LegacyPptWritePreflightReport report = presentation
                .AnalyzeLegacyPptWrite();
            Assert.False(report.CanWrite);
            Assert.Contains(report.Findings, finding =>
                finding.Code == "PPT-WRITE-IMPORT-LOSS");
            Assert.Throws<NotSupportedException>(() =>
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt,
                    new PowerPointSaveOptions {
                        LossPolicy = PowerPointConversionLossPolicy.Allow
                    }));

            Assert.Equal(selectedMaster.MasterId, saved.Slides[1].MasterId);
            Assert.Contains(saved.Slides[1].Shapes, shape => shape.Text == "Uses second binary master");
        }

        [Fact]
        public void NativeWriter_BlocksKnownLossUnlessExplicitlyAllowed() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.AddSlide().AddChart();

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.True(preflight.HasConversionLoss);
            Assert.Throws<NotSupportedException>(() => presentation.ToBytes(PowerPointFileFormat.Ppt));

            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt,
                new PowerPointSaveOptions { LossPolicy = PowerPointConversionLossPolicy.Allow });
            Assert.NotEmpty(bytes);
            Assert.Single(LegacyPptPresentation.Load(bytes).Slides);
        }

        [Fact]
        public void NativeWriter_PreflightReportsUnsupportedTransitionAndVisualStyleLoss() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.Transition = SlideTransition.Morph;
            PowerPointAutoShape shape = slide.AddRectangle(
                100000, 100000, 1000000, 500000);
            shape.Fill("FF0000");
            shape.SetGlow("4472C4", radiusPoints: 4D);

            LegacyPptWritePreflightReport report = presentation.AnalyzeLegacyPptWrite();

            Assert.Contains(report.Findings, finding => finding.Code == "PPT-WRITE-TRANSITION");
            Assert.Contains(report.Findings, finding => finding.Code == "PPT-WRITE-SHAPE-STYLE");
        }

        [Fact]
        public void AssociatedLegacyStream_SavePreservesBinaryFormat() {
            byte[] source = File.ReadAllBytes(FixturePath);
            using var stream = new MemoryStream();
            stream.Write(source, 0, source.Length);
            stream.Position = 0;

            using (PowerPointPresentation presentation = PowerPointPresentation.Load(stream)) {
                presentation.Slides[0].AddTextBox("Saved back as binary");
                Assert.Throws<NotSupportedException>(() => presentation.Save());
                presentation.Save(stream, presentation.SourceFormat,
                    new PowerPointSaveOptions { LossPolicy = PowerPointConversionLossPolicy.Allow });
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(stream.ToArray());
            Assert.Contains(legacy.Slides[0].Shapes, shape => shape.Text == "Saved back as binary");
        }

        [Theory]
        [InlineData(".ppt")]
        [InlineData(".pot")]
        [InlineData(".pps")]
        public void SaveCopy_UsesBinaryWriterForLegacyExtensions(string extension) {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + extension);
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create();
                presentation.AddSlide().AddTextBox("Legacy extension");
                presentation.SaveCopy(path);

                LegacyPptPresentation binary = LegacyPptPresentation.Load(path);
                Assert.Single(binary.Slides);
                Assert.Contains(binary.Slides[0].Shapes, shape => shape.Text == "Legacy extension");
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public async Task NativeWriter_AsyncPathAndStreamOverloadsProduceBinaryFiles() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".ppt");
            try {
                await using PowerPointPresentation presentation = PowerPointPresentation.Create(path);
                presentation.AddSlide().AddTextBox("Async binary");

                await presentation.SaveAsync();
                Assert.Contains(LegacyPptPresentation.Load(path).Slides[0].Shapes,
                    shape => shape.Text == "Async binary");

                using var stream = new MemoryStream();
                await presentation.SaveAsync(stream, PowerPointFileFormat.Pps);
                Assert.Contains(LegacyPptPresentation.Load(stream.ToArray()).Slides[0].Shapes,
                    shape => shape.Text == "Async binary");
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }
}
