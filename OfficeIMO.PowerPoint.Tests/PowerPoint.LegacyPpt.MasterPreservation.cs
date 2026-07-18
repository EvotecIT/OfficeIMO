using System.Buffers.Binary;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using OfficeIMO.Tests.Pdf;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptMasterTests {
        [Fact]
        public void ImportedMainMasterShapeMove_AppendsPreservingIncrementalRecord() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(FixturePath);
            LegacyPptMaster[] mainMasters = original.Masters
                .Where(master => master.IsMainMaster).ToArray();
            int masterIndex = Array.FindIndex(mainMasters,
                master => master.Shapes.Count > 0);
            Assert.True(masterIndex >= 0);
            LegacyPptMaster originalMaster = mainMasters[masterIndex];

            using PowerPointPresentation imported = PowerPointPresentation.Load(
                FixturePath);
            SlideMasterPart masterPart = imported.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.ElementAt(masterIndex);
            IReadOnlyList<PowerPointShape> projectedShapes = LegacyPptWriter
                .ReadMasterShapesForWrite(masterPart, out string? reason);
            Assert.Null(reason);
            Assert.Equal(originalMaster.Shapes.Count, projectedShapes.Count);
            PowerPointShape projectedShape = projectedShapes[0];
            long expectedLeft = projectedShape.Left + 15875L;
            projectedShape.Left = expectedLeft;
            masterPart.SlideMaster!.CommonSlideData!.Background =
                CreateSolidBackground("0F1E2D");

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                imported.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptMaster savedMaster = Assert.Single(saved.Masters,
                master => master.MasterId == originalMaster.MasterId);

            Assert.Equal(originalMaster.Shapes[0].Bounds.Left + 10,
                savedMaster.Shapes[0].Bounds.Left);
            Assert.Equal("0F1E2D", Assert.IsType<LegacyPptBackground>(
                savedMaster.Background).ForegroundColor);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
            AssertUnrelatedMasterChildrenEqual(original, saved,
                originalMaster.PersistId);

            using var stream = new MemoryStream(saved.Package.CopyOriginalBytes());
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(stream);
            SlideMasterPart reopenedMaster = reopened.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.ElementAt(masterIndex);
            PowerPointShape reopenedShape = LegacyPptWriter
                .ReadMasterShapesForWrite(reopenedMaster, out _)[0];
            Assert.Equal(expectedLeft, reopenedShape.Left);
            Assert.Equal("0F1E2D", reopenedMaster.SlideMaster!
                .CommonSlideData!.Background!.BackgroundProperties!
                .GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedMainMasterTextAndPlaceholderEdits_AppendPreservingRecords() {
            byte[] sourceBytes = CreateBinaryWithEditableMasterText();
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);
            LegacyPptMaster originalMaster = Assert.Single(original.Masters);

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            SlideMasterPart masterPart = imported.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single();
            PowerPointTextBox textBox = Assert.IsType<PowerPointTextBox>(
                Assert.Single(LegacyPptWriter.ReadMasterShapesForWrite(
                    masterPart, out _)));
            textBox.Text = "Edited label";
            P.Shape textShape = Assert.IsType<P.Shape>(textBox.Element);
            A.Run editedRun = Assert.Single(textShape.TextBody!
                .Descendants<A.Run>());
            editedRun.RunProperties = new A.RunProperties(
                new A.LatinFont { Typeface = "OfficeIMO Master" }) {
                Bold = true,
                FontSize = 2800
            };
            textShape.TextBody.Elements<A.Paragraph>().Single()
                .ParagraphProperties = new A.ParagraphProperties {
                Alignment = A.TextAlignmentTypeValues.Center
            };
            textBox.SetTextMarginsPoints(5, 6, 7, 8);
            textBox.TextVerticalAlignment =
                A.TextAnchoringTypeValues.Bottom;
            textBox.TextDirection = A.TextVerticalValues.Vertical;
            textBox.TextAutoFit = PowerPointTextAutoFit.None;
            textShape.TextBody.BodyProperties!.Wrap =
                A.TextWrappingValues.None;
            textBox.PlaceholderType = P.PlaceholderValues.Title;
            textBox.PlaceholderIndex = 3;
            textBox.PlaceholderSize = P.PlaceholderSizeValues.Half;

            Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
            byte[] savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            LegacyPptMaster savedMaster = Assert.Single(saved.Masters);

            Assert.Equal("Edited label", Assert.Single(savedMaster.Shapes).Text);
            LegacyPptShape savedShape = Assert.Single(savedMaster.Shapes);
            LegacyPptCharacterRun savedRun = Assert.Single(savedShape.TextBody
                .CharacterRuns, run => run.Text == "Edited label");
            Assert.True(savedRun.Bold);
            Assert.Equal((short)28, savedRun.FontSizePoints);
            Assert.Equal("OfficeIMO Master", savedRun.Typeface);
            Assert.Equal(LegacyPptTextAlignment.Center,
                Assert.Single(savedShape.TextBody.ParagraphRuns).Alignment);
            Assert.Equal(63500, savedShape.TextFrame.LeftInsetEmus);
            Assert.Equal(76200, savedShape.TextFrame.TopInsetEmus);
            Assert.Equal(88900, savedShape.TextFrame.RightInsetEmus);
            Assert.Equal(101600, savedShape.TextFrame.BottomInsetEmus);
            Assert.Equal(2U, savedShape.TextFrame.WrapMode);
            Assert.Equal(2U, savedShape.TextFrame.AnchorMode);
            Assert.Equal(1U, savedShape.TextFrame.TextFlow);
            Assert.Equal(false, savedShape.TextFrame.FitShapeToText);
            Assert.Contains(saved.Fonts,
                font => font.Typeface == "OfficeIMO Master");
            LegacyPptPlaceholder placeholder = Assert.IsType<
                LegacyPptPlaceholder>(Assert.Single(savedMaster.Shapes)
                    .Placeholder);
            Assert.Equal(3, placeholder.Position);
            Assert.Equal(LegacyPptPlaceholderKind.MasterTitle,
                placeholder.Kind);
            Assert.Equal(LegacyPptPlaceholderSize.Half, placeholder.Size);
            Assert.Equal(LegacyPptPlaceholderKind.MasterTitle,
                savedMaster.LayoutPlaceholderTypes[3]);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
            AssertUnrelatedMasterChildrenEqual(original, saved,
                originalMaster.PersistId, 0x03EF);

            using var reopenedInput = new MemoryStream(savedBytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            SlideMasterPart reopenedMaster = reopened.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single();
            PowerPointTextBox reopenedText = Assert.IsType<PowerPointTextBox>(
                Assert.Single(LegacyPptWriter.ReadMasterShapesForWrite(
                    reopenedMaster, out _)));
            Assert.Equal("Edited label", reopenedText.Text);
            Assert.Equal(P.PlaceholderValues.Title,
                reopenedText.PlaceholderType);
            Assert.Equal(3U, reopenedText.PlaceholderIndex);
            Assert.Equal(P.PlaceholderSizeValues.Half,
                reopenedText.PlaceholderSize);
            Assert.Equal(5D, reopenedText.TextMarginLeftPoints);
            Assert.Equal(6D, reopenedText.TextMarginTopPoints);
            Assert.Equal(7D, reopenedText.TextMarginRightPoints);
            Assert.Equal(8D, reopenedText.TextMarginBottomPoints);
            Assert.Equal(A.TextAnchoringTypeValues.Bottom,
                reopenedText.TextVerticalAlignment);
            Assert.Equal(A.TextVerticalValues.Vertical,
                reopenedText.TextDirection);
            Assert.Equal(PowerPointTextAutoFit.None,
                reopenedText.TextAutoFit);
            P.Shape reopenedShape = Assert.IsType<P.Shape>(
                reopenedText.Element);
            Assert.Equal(A.TextWrappingValues.None,
                reopenedShape.TextBody!.BodyProperties!.Wrap!.Value);
            A.Run reopenedRun = Assert.Single(reopenedShape.TextBody!
                .Descendants<A.Run>());
            Assert.True(reopenedRun.RunProperties!.Bold!.Value);
            Assert.Equal(2800, reopenedRun.RunProperties.FontSize!.Value);
            Assert.Equal("OfficeIMO Master", reopenedRun.RunProperties
                .GetFirstChild<A.LatinFont>()!.Typeface!.Value);
            Assert.Equal(A.TextAlignmentTypeValues.Center,
                reopenedShape.TextBody.Elements<A.Paragraph>().Single()
                    .ParagraphProperties!.Alignment!.Value);
            Assert.Empty(reopened.ValidateDocument());

            reopenedText.PlaceholderIndex = null;
            reopenedText.PlaceholderSize = null;
            reopenedText.PlaceholderType = null;
            Assert.False(reopenedText.IsPlaceholder);
            Assert.True(reopened.AnalyzeLegacyPptWrite().CanWrite);
            byte[] removedBytes = reopened.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation removed = LegacyPptPresentation.Load(
                removedBytes);
            LegacyPptMaster removedMaster = Assert.Single(removed.Masters);

            Assert.Null(Assert.Single(removedMaster.Shapes).Placeholder);
            Assert.DoesNotContain(removedMaster.LayoutPlaceholderTypes,
                value => value != LegacyPptPlaceholderKind.None);
            Assert.Equal(saved.Package.UserEdits.Count + 1,
                removed.Package.UserEdits.Count);
            Assert.True(removed.Package.DocumentStream.AsSpan(0,
                    saved.Package.DocumentStream.Length)
                .SequenceEqual(saved.Package.DocumentStream));

            using var removedInput = new MemoryStream(removedBytes);
            using PowerPointPresentation removedProjection =
                PowerPointPresentation.Load(removedInput);
            PowerPointTextBox removedText = Assert.IsType<PowerPointTextBox>(
                Assert.Single(LegacyPptWriter.ReadMasterShapesForWrite(
                    removedProjection.OpenXmlDocument.PresentationPart!
                        .SlideMasterParts.Single(), out _)));
            Assert.False(removedText.IsPlaceholder);
            Assert.Empty(removedProjection.ValidateDocument());
        }

        [Fact]
        public void ImportedMainMasterTransformEdit_PreservesRichStyle() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                FixturePath);
            using PowerPointPresentation imported = PowerPointPresentation.Load(
                FixturePath);
            P.Shape shape = imported.OpenXmlDocument.PresentationPart!
                .SlideMasterParts.SelectMany(master => master.SlideMaster!
                    .CommonSlideData!.ShapeTree!.Descendants<P.Shape>())
                .First();
            shape.ShapeProperties!.Transform2D!.Rotation = 60000;

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();

            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                imported.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptShape savedShape = saved.Masters
                .SelectMany(master => master.Shapes)
                .First();
            LegacyPptShape originalShape = original.Masters
                .SelectMany(master => master.Shapes)
                .First();
            Assert.Equal(1D, savedShape.Transform.RotationDegrees);
            Assert.Equal(originalShape.Style.FillType,
                savedShape.Style.FillType);
            Assert.Equal(originalShape.FillColor, savedShape.FillColor);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
        }

        [Fact]
        public void ImportedNotesMasterShapeMove_AppendsPreservingIncrementalRecord() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(FixturePath);
            LegacyPptSpecialMaster originalMaster = Assert.IsType<
                LegacyPptSpecialMaster>(original.NotesMaster);
            Assert.NotEmpty(originalMaster.Shapes);
            using PowerPointPresentation imported = PowerPointPresentation.Load(
                FixturePath);
            NotesMasterPart notesPart = imported.OpenXmlDocument.PresentationPart!
                .NotesMasterPart!;
            PowerPointShape shape = LegacyPptWriter.ReadMasterShapesForWrite(
                notesPart, out _)[0];
            long expectedLeft = shape.Left + 15875L;
            shape.Left = expectedLeft;
            A.Accent6Color accent6 = notesPart.ThemePart!.Theme!
                .ThemeElements!.ColorScheme!.GetFirstChild<A.Accent6Color>()!;
            accent6.RemoveAllChildren();
            accent6.Append(new A.RgbColorModelHex { Val = "2468AC" });

            Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
            byte[] savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            LegacyPptSpecialMaster savedMaster = Assert.IsType<
                LegacyPptSpecialMaster>(saved.NotesMaster);

            Assert.Equal(originalMaster.Shapes[0].Bounds.Left + 10,
                savedMaster.Shapes[0].Bounds.Left);
            Assert.Equal("2468AC", savedMaster.RoundTripTheme?
                .Colors[PowerPointThemeColor.Accent6]);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
            AssertUnrelatedMasterChildrenEqual(original, saved,
                originalMaster.PersistId, 0x040E, 0x040F);

            using var reopenedInput = new MemoryStream(savedBytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            PowerPointShape reopenedShape = LegacyPptWriter
                .ReadMasterShapesForWrite(reopened.OpenXmlDocument
                    .PresentationPart!.NotesMasterPart!, out _)[0];
            Assert.Equal(expectedLeft, reopenedShape.Left);
            Assert.Equal("2468AC", reopened.OpenXmlDocument.PresentationPart!
                .NotesMasterPart!.ThemePart!.Theme!.ThemeElements!
                .ColorScheme!.GetFirstChild<A.Accent6Color>()!
                .GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedHandoutMasterShapeAndThemeEdit_AppendsPreservingRecord() {
            byte[] sourceBytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                HandoutMasterPart handoutPart = CreateHandoutMaster(created);
                handoutPart.HandoutMaster!.CommonSlideData!.ShapeTree!.Append(
                    CreateNotesMasterShape(2U, "Handout marker",
                        new PowerPointLayoutBox(300000, 400000, 500000, 500000),
                        placeholder: null, text: null,
                        shapeType: A.ShapeTypeValues.Ellipse));
                created.AddSlide(P.SlideLayoutValues.Blank);
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);
            LegacyPptSpecialMaster originalMaster = Assert.IsType<
                LegacyPptSpecialMaster>(original.HandoutMaster);

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            HandoutMasterPart handoutMasterPart = imported.OpenXmlDocument
                .PresentationPart!.HandoutMasterPart!;
            PowerPointShape shape = Assert.Single(LegacyPptWriter
                .ReadMasterShapesForWrite(handoutMasterPart, out _));
            shape.Left += 15875L;
            A.Accent5Color accent5 = handoutMasterPart.ThemePart!.Theme!
                .ThemeElements!.ColorScheme!.GetFirstChild<A.Accent5Color>()!;
            accent5.RemoveAllChildren();
            accent5.Append(new A.RgbColorModelHex { Val = "13579B" });
            handoutMasterPart.HandoutMaster!.CommonSlideData!.Background =
                CreateSolidBackground("ABCDEF");

            Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
            byte[] savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            LegacyPptSpecialMaster savedMaster = Assert.IsType<
                LegacyPptSpecialMaster>(saved.HandoutMaster);

            Assert.Equal(originalMaster.Shapes[0].Bounds.Left + 10,
                savedMaster.Shapes[0].Bounds.Left);
            Assert.Equal("13579B", savedMaster.RoundTripTheme?
                .Colors[PowerPointThemeColor.Accent5]);
            Assert.Equal("ABCDEF", Assert.IsType<LegacyPptBackground>(
                savedMaster.Background).ForegroundColor);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            using var reopenedInput = new MemoryStream(savedBytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            HandoutMasterPart reopenedPart = reopened.OpenXmlDocument
                .PresentationPart!.HandoutMasterPart!;
            Assert.Equal(shape.Left, Assert.Single(LegacyPptWriter
                .ReadMasterShapesForWrite(reopenedPart, out _)).Left);
            Assert.Equal("13579B", reopenedPart.ThemePart!.Theme!
                .ThemeElements!.ColorScheme!.GetFirstChild<A.Accent5Color>()!
                .GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.Equal("ABCDEF", reopenedPart.HandoutMaster!
                .CommonSlideData!.Background!.BackgroundProperties!
                .GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedTitleMasterShapeThemeAndBackgroundEdit_AppendsPreservingRecord() {
            byte[] imageBytes = PdfPngTestImages.CreateRgbPng(33, 99, 165);
            byte[] sourceBytes = CreateBinaryWithEditableTitleMaster();
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);
            LegacyPptMaster mainMaster = Assert.Single(original.Masters,
                master => master.IsMainMaster);
            LegacyPptMaster originalTitleMaster = Assert.Single(original.Masters,
                master => !master.IsMainMaster);
            Assert.Equal(mainMaster.MasterId, originalTitleMaster.ParentMasterId);
            Assert.NotEmpty(originalTitleMaster.Shapes);

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            string titleName =
                $"Binary Title Master {originalTitleMaster.MasterId:X8}";
            SlideLayoutPart titlePart = imported.OpenXmlDocument.PresentationPart!
                .SlideMasterParts.SelectMany(master => master.SlideLayoutParts)
                .Single(layout => string.Equals(layout.SlideLayout?
                    .CommonSlideData?.Name?.Value, titleName,
                    StringComparison.Ordinal));
            Assert.False(titlePart.SlideLayout!.ShowMasterShapes!.Value);
            Assert.DoesNotContain(titlePart.SlideLayout.CommonSlideData!
                .GetAttributes(), attribute =>
                    attribute.LocalName == "showMasterSp");
            Assert.Empty(imported.ValidateDocument());
            titlePart.SlideLayout.ShowMasterShapes = true;
            PowerPointShape shape = Assert.Single(LegacyPptWriter
                .ReadMasterShapesForWrite(titlePart, out string? shapeReason));
            Assert.Null(shapeReason);
            long expectedLeft = shape.Left + 15875L;
            shape.Left = expectedLeft;
            A.ThemeOverride theme = Assert.IsType<A.ThemeOverride>(
                titlePart.ThemeOverridePart?.ThemeOverride);
            A.Accent5Color accent5 = Assert.IsType<A.Accent5Color>(
                theme.ColorScheme?.GetFirstChild<A.Accent5Color>());
            accent5.RemoveAllChildren();
            accent5.Append(new A.RgbColorModelHex { Val = "5A6B7C" });
            SetPictureBackground(titlePart,
                titlePart.SlideLayout!.CommonSlideData!, imageBytes);

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            LegacyPptMaster savedTitleMaster = Assert.Single(saved.Masters,
                master => master.MasterId == originalTitleMaster.MasterId);

            Assert.False(savedTitleMaster.IsMainMaster);
            Assert.True(savedTitleMaster.FollowsMasterObjects);
            Assert.Equal(originalTitleMaster.Shapes[0].Bounds.Left + 10,
                savedTitleMaster.Shapes[0].Bounds.Left);
            Assert.Equal("5A6B7C", savedTitleMaster.RoundTripTheme?
                .Colors[PowerPointThemeColor.Accent5]);
            LegacyPptBackground savedBackground = Assert.IsType<
                LegacyPptBackground>(savedTitleMaster.Background);
            Assert.Equal(LegacyPptBackgroundKind.Picture,
                savedBackground.Kind);
            Assert.Equal(imageBytes, savedBackground.Picture!.ImageBytes);
            Assert.Equal(1U, Assert.Single(saved.BlipStoreEntries)
                .ReferenceCount);
            LegacyPptSlide savedTitleSlide = Assert.Single(saved.Slides,
                slide => slide.MasterId == originalTitleMaster.MasterId);
            Assert.True(savedTitleSlide.FollowsMasterBackground);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
            AssertUnrelatedMasterChildrenEqual(original, saved,
                originalTitleMaster.PersistId, 0x03EF, 0x040E, 0x040F);

            using var reopenedInput = new MemoryStream(savedBytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            SlideLayoutPart reopenedTitlePart = reopened.OpenXmlDocument
                .PresentationPart!.SlideMasterParts
                .SelectMany(master => master.SlideLayoutParts)
                .Single(layout => string.Equals(layout.SlideLayout?
                    .CommonSlideData?.Name?.Value, titleName,
                    StringComparison.Ordinal));
            Assert.True(reopenedTitlePart.SlideLayout!.ShowMasterShapes?.Value
                != false);
            Assert.Equal(expectedLeft, Assert.Single(LegacyPptWriter
                .ReadMasterShapesForWrite(reopenedTitlePart, out _)).Left);
            Assert.Equal("5A6B7C", reopenedTitlePart.ThemeOverridePart!
                .ThemeOverride!.ColorScheme!.GetFirstChild<A.Accent5Color>()!
                .GetFirstChild<A.RgbColorModelHex>()!.Val!.Value);
            Assert.NotNull(reopenedTitlePart.SlideLayout!.CommonSlideData!
                .Background!.BackgroundProperties!
                .GetFirstChild<A.BlipFill>());
            PowerPointSlide reopenedTitleSlide = Assert.Single(reopened.Slides,
                slide => ReferenceEquals(slide.SlidePart.SlideLayoutPart,
                    reopenedTitlePart));
            Assert.Null(reopenedTitleSlide.SlidePart.Slide!.CommonSlideData!
                .Background);
            Assert.Equal(imageBytes,
                reopenedTitleSlide.GetBackground().ImageBytes);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedTitleMasterSlidePictureBackgroundEdit_RemainsSlideSpecific() {
            byte[] imageBytes = PdfPngTestImages.CreateRgbPng(85, 145, 205);
            byte[] sourceBytes = CreateBinaryWithEditableTitleMaster();
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);
            LegacyPptMaster titleMaster = Assert.Single(original.Masters,
                master => !master.IsMainMaster);

            using var input = new MemoryStream(sourceBytes, writable: false);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            string titleName = $"Binary Title Master {titleMaster.MasterId:X8}";
            SlideLayoutPart titlePart = imported.OpenXmlDocument.PresentationPart!
                .SlideMasterParts.SelectMany(master => master.SlideLayoutParts)
                .Single(layout => string.Equals(layout.SlideLayout?
                    .CommonSlideData?.Name?.Value, titleName,
                    StringComparison.Ordinal));
            PowerPointSlide titleSlide = Assert.Single(imported.Slides,
                slide => ReferenceEquals(slide.SlidePart.SlideLayoutPart,
                    titlePart));
            Assert.Null(titleSlide.SlidePart.Slide!.CommonSlideData!.Background);
            using (var image = new MemoryStream(imageBytes, writable: false)) {
                titleSlide.SetBackgroundImage(image,
                    OfficeIMO.PowerPoint.ImagePartType.Png);
            }

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            LegacyPptSlide savedTitleSlide = Assert.Single(saved.Slides,
                slide => slide.MasterId == titleMaster.MasterId);

            Assert.False(savedTitleSlide.FollowsMasterBackground);
            LegacyPptBackground savedBackground = Assert.IsType<
                LegacyPptBackground>(savedTitleSlide.Background);
            Assert.Equal(LegacyPptBackgroundKind.Picture,
                savedBackground.Kind);
            Assert.Equal(imageBytes, savedBackground.Picture!.ImageBytes);
            Assert.Equal(1U, Assert.Single(saved.BlipStoreEntries)
                .ReferenceCount);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            using var reopenedInput = new MemoryStream(savedBytes,
                writable: false);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            PowerPointSlide reopenedTitleSlide = Assert.Single(reopened.Slides,
                slide => slide.SlidePart.SlideLayoutPart?.SlideLayout?
                    .CommonSlideData?.Name?.Value == titleName);
            Assert.NotNull(reopenedTitleSlide.SlidePart.Slide!
                .CommonSlideData!.Background!.BackgroundProperties!
                .GetFirstChild<A.BlipFill>());
            Assert.Equal(imageBytes,
                reopenedTitleSlide.GetBackground().ImageBytes);
            Assert.Empty(reopened.ValidateDocument());
        }

        private static byte[] CreateBinaryWithEditableMasterText() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            SlideMasterPart masterPart = presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single();
            P.ShapeTree tree = masterPart.SlideMaster!.CommonSlideData!
                .ShapeTree!;
            tree.Append(new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties {
                        Id = 2U,
                        Name = "Editable master label"
                    },
                    new P.NonVisualShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 400000, Y = 500000 },
                        new A.Extents { Cx = 3000000, Cy = 600000 }),
                    new A.PresetGeometry(new A.AdjustValueList()) {
                        Preset = A.ShapeTypeValues.Rectangle
                    }),
                new P.TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(new A.Text("Master label")),
                        new A.EndParagraphRunProperties()))));
            presentation.AddSlide(P.SlideLayoutValues.Blank);
            return presentation.ToBytes(PowerPointFileFormat.Ppt);
        }

        private static byte[] CreateBinaryWithEditableTitleMaster() {
            using PowerPointPresentation target = PowerPointPresentation.Create();
            target.SetThemeColor(PowerPointThemeColor.Accent1, "102030");
            target.AddSlide();

            using PowerPointPresentation source = PowerPointPresentation.Create();
            source.SetThemeColor(PowerPointThemeColor.Accent1, "A0B0C0");
            SlideMasterPart sourceMaster = source.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single();
            sourceMaster.ThemePart!.Theme!.Save();
            SlideLayoutPart sourceLayout = sourceMaster.SlideLayoutParts.First();
            sourceLayout.SlideLayout!.CommonSlideData!.Name =
                "Imported editable title-master layout";
            sourceLayout.SlideLayout.Save();
            sourceMaster.SlideMaster!.CommonSlideData!.ShapeTree!.Append(
                new P.Shape(
                    new P.NonVisualShapeProperties(
                        new P.NonVisualDrawingProperties {
                            Id = 2U,
                            Name = "Editable title-master marker"
                        },
                        new P.NonVisualShapeDrawingProperties(),
                        new P.ApplicationNonVisualDrawingProperties()),
                    new P.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 400000, Y = 500000 },
                            new A.Extents { Cx = 1000000, Cy = 700000 }),
                        new A.PresetGeometry(new A.AdjustValueList()) {
                            Preset = A.ShapeTypeValues.Ellipse
                        })));
            sourceMaster.SlideMaster.Save();
            source.AddSlide();
            target.ImportSlide(source, 0);

            LegacyPptPresentation generated = LegacyPptPresentation.Load(
                target.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptMaster[] generatedMasters = generated.Masters.ToArray();
            Assert.Equal(2, generatedMasters.Length);
            LegacyPptMaster parent = generatedMasters[0];
            LegacyPptMaster title = generatedMasters[1];
            Assert.All(generatedMasters, master => Assert.True(master.IsMainMaster));
            Assert.NotEmpty(title.Shapes);

            LegacyPptPersistObject persistObject = generated.Package
                .PersistObjects[title.PersistId];
            LegacyPptRecord titleRecord = LegacyPptRecordReader.ReadSingle(
                persistObject.RecordBytes, 0, new LegacyPptImportOptions());
            LegacyPptRecord slideAtom = Assert.Single(titleRecord.Children,
                record => record.Type == 0x03EF);
            Assert.True(slideAtom.PayloadLength >= 24);
            byte[] documentStream = (byte[])generated.Package.DocumentStream.Clone();
            int recordOffset = checked((int)persistObject.StreamOffset);
            BinaryPrimitives.WriteUInt16LittleEndian(
                documentStream.AsSpan(recordOffset + 2, 2), 0x03EE);
            int slideAtomPayloadOffset = checked(recordOffset
                + slideAtom.PayloadOffset);
            BinaryPrimitives.WriteUInt32LittleEndian(
                documentStream.AsSpan(slideAtomPayloadOffset + 12, 4),
                parent.MasterId);
            BinaryPrimitives.WriteUInt16LittleEndian(
                documentStream.AsSpan(slideAtomPayloadOffset + 20, 2), 0x0000);
            return generated.Package.RewriteCompoundStreams(
                new Dictionary<string, byte[]> {
                    ["PowerPoint Document"] = documentStream
                });
        }

        private static P.Background CreateSolidBackground(string color) => new(
            new P.BackgroundProperties(
                new A.SolidFill(new A.RgbColorModelHex { Val = color })));

        private static void AssertUnrelatedMasterChildrenEqual(
            LegacyPptPresentation original, LegacyPptPresentation saved,
            uint persistId, params ushort[] additionallyExcludedTypes) {
            IReadOnlyList<byte[]> originalChildren = ReadMasterChildrenExceptDrawing(
                original, persistId, additionallyExcludedTypes);
            IReadOnlyList<byte[]> savedChildren = ReadMasterChildrenExceptDrawing(
                saved, persistId, additionallyExcludedTypes);
            Assert.Equal(originalChildren.Count, savedChildren.Count);
            for (int index = 0; index < originalChildren.Count; index++) {
                Assert.True(originalChildren[index]
                    .SequenceEqual(savedChildren[index]));
            }
        }

        private static IReadOnlyList<byte[]> ReadMasterChildrenExceptDrawing(
            LegacyPptPresentation presentation, uint persistId,
            IReadOnlyCollection<ushort> additionallyExcludedTypes) {
            LegacyPptPersistObject persistObject =
                presentation.Package.PersistObjects[persistId];
            LegacyPptRecord record = LegacyPptRecordReader.ReadSingle(
                persistObject.RecordBytes, 0, new LegacyPptImportOptions());
            return record.Children.Where(child => child.Type != 0x040C
                    && !additionallyExcludedTypes.Contains(child.Type))
                .Select(child => child.CopyRecordBytes()).ToArray();
        }
    }
}
