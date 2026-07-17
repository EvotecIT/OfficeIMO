using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using OfficeIMO.Tests.Pdf;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptMasterTests {
        [Fact]
        public void NativeWriter_WritesDeduplicatedPicturesOnAllMasterTypes() {
            byte[] imageBytes = PdfPngTestImages.CreateRgbPng(62, 122, 202);
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PresentationPart presentationPart = presentation.OpenXmlDocument
                .PresentationPart!;
            SlideMasterPart mainPart = presentationPart.SlideMasterParts
                .Single();
            NotesMasterPart notesPart = presentationPart.NotesMasterPart!;
            HandoutMasterPart handoutPart = CreateHandoutMaster(presentation);
            P.Picture hiddenMainPicture = AddPictureShape(mainPart,
                mainPart.SlideMaster!.CommonSlideData!.ShapeTree!, imageBytes,
                100U, 200000L);
            hiddenMainPicture.NonVisualPictureProperties!
                .NonVisualDrawingProperties!.Hidden = true;
            AddPictureShape(notesPart,
                notesPart.NotesMaster!.CommonSlideData!.ShapeTree!, imageBytes,
                101U, 400000L);
            AddPictureShape(handoutPart,
                handoutPart.HandoutMaster!.CommonSlideData!.ShapeTree!,
                imageBytes, 102U, 600000L);
            presentation.AddSlide(P.SlideLayoutValues.Blank);

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);

            AssertMasterPicture(Assert.Single(binary.Masters).Shapes,
                imageBytes, expectedHidden: true);
            AssertMasterPicture(Assert.IsType<LegacyPptSpecialMaster>(
                binary.NotesMaster).Shapes, imageBytes);
            AssertMasterPicture(Assert.IsType<LegacyPptSpecialMaster>(
                binary.HandoutMaster).Shapes, imageBytes);
            Assert.Equal(3U, Assert.Single(binary.BlipStoreEntries)
                .ReferenceCount);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(input);
            AssertProjectedMasterPicture(LegacyPptWriter
                .ReadMasterShapesForWrite(reopened.OpenXmlDocument
                    .PresentationPart!.SlideMasterParts.Single(), out _),
                imageBytes, expectedHidden: true);
            AssertProjectedMasterPicture(LegacyPptWriter
                .ReadMasterShapesForWrite(reopened.OpenXmlDocument
                    .PresentationPart!.NotesMasterPart!, out _), imageBytes);
            AssertProjectedMasterPicture(LegacyPptWriter
                .ReadMasterShapesForWrite(reopened.OpenXmlDocument
                    .PresentationPart!.HandoutMasterPart!, out _), imageBytes);
            Assert.Empty(reopened.ValidateDocument());
            Assert.Equal(bytes,
                reopened.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void NativeWriter_RasterizesUnsupportedTablesOnMastersWhenAllowed() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide(
                P.SlideLayoutValues.Blank);
            PowerPointTable sourceTable = slide.AddTable(2, 2);
            sourceTable.GetCell(0, 0).Text = "Merged master";
            sourceTable.MergeCells(0, 0, 0, 1);
            P.GraphicFrame masterTable = Assert.IsType<P.GraphicFrame>(
                sourceTable.Element.CloneNode(true));
            masterTable.NonVisualGraphicFrameProperties!
                .NonVisualDrawingProperties!.Id = 100U;
            masterTable.NonVisualGraphicFrameProperties
                .NonVisualDrawingProperties.Name = "Merged master table";
            sourceTable.Element.Remove();
            SlideMasterPart masterPart = presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single();
            masterPart.SlideMaster!.CommonSlideData!.ShapeTree!
                .Append(masterTable);
            masterPart.SlideMaster.Save();

            LegacyPptWriteFinding finding = Assert.Single(presentation
                .AnalyzeLegacyPptWrite().Findings, item =>
                    item.Code == "PPT-WRITE-MASTER-TABLE");
            Assert.Equal(LegacyPptFeature.Tables, finding.Feature);
            Assert.Throws<NotSupportedException>(() => presentation.ToBytes(
                PowerPointFileFormat.Ppt));

            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt,
                new PowerPointSaveOptions {
                    LossPolicy = PowerPointConversionLossPolicy.Allow
                });
            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);
            Assert.Contains(Assert.Single(binary.Masters).Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Picture);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(input);
            Assert.Contains(LegacyPptWriter.ReadMasterShapesForWrite(
                reopened.OpenXmlDocument.PresentationPart!
                    .SlideMasterParts.Single(), out _),
                shape => shape is PowerPointPicture);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_MaterializesLayoutPicturesIntoAffectedSlides() {
            byte[] imageBytes = PdfPngTestImages.CreateRgbPng(142, 82, 42);
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            SlideMasterPart masterPart = presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single();
            int layoutIndex = presentation.GetLayoutIndex(
                P.SlideLayoutValues.Blank);
            SlideLayoutPart layoutPart = masterPart.SlideLayoutParts
                .ElementAt(layoutIndex);
            AddPictureShape(layoutPart,
                layoutPart.SlideLayout!.CommonSlideData!.ShapeTree!,
                imageBytes, 100U, 350000L);
            presentation.AddSlide(0, layoutIndex);
            presentation.AddSlide(0, layoutIndex);

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);

            Assert.All(binary.Slides, slide => AssertMasterPicture(
                slide.Shapes, imageBytes));
            Assert.Equal(2U, Assert.Single(binary.BlipStoreEntries)
                .ReferenceCount);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(input);
            Assert.All(reopened.Slides, slide => Assert.Equal(imageBytes,
                Assert.Single(slide.Pictures).GetImageBytes()));
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedMasterPictureVisibilityEdit_IsIncremental() {
            byte[] imageBytes = PdfPngTestImages.CreateRgbPng(82, 42, 122);
            byte[] sourceBytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                SlideMasterPart masterPart = source.OpenXmlDocument
                    .PresentationPart!.SlideMasterParts.Single();
                P.Picture picture = AddPictureShape(masterPart,
                    masterPart.SlideMaster!.CommonSlideData!.ShapeTree!,
                    imageBytes, 100U, 250000L);
                picture.NonVisualPictureProperties!
                    .NonVisualDrawingProperties!.Hidden = true;
                source.AddSlide(P.SlideLayoutValues.Blank);
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);

            byte[] editedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation
                       .Load(input)) {
                PowerPointPicture picture = Assert.IsType<PowerPointPicture>(
                    Assert.Single(LegacyPptWriter.ReadMasterShapesForWrite(
                        imported.OpenXmlDocument.PresentationPart!
                            .SlideMasterParts.Single(), out _), shape =>
                        shape is PowerPointPicture));
                Assert.True(picture.Hidden);
                picture.Hidden = false;

                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                editedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation edited = LegacyPptPresentation.Load(
                editedBytes);
            Assert.False(Assert.Single(Assert.Single(edited.Masters).Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Picture)
                .Style.Hidden ?? false);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                edited.Package.UserEdits.Count);
            Assert.True(edited.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            using var editedInput = new MemoryStream(editedBytes,
                writable: false);
            using PowerPointPresentation reopened = PowerPointPresentation
                .Load(editedInput);
            PowerPointPicture reopenedPicture = Assert.IsType<
                PowerPointPicture>(Assert.Single(LegacyPptWriter
                .ReadMasterShapesForWrite(reopened.OpenXmlDocument
                    .PresentationPart!.SlideMasterParts.Single(), out _),
                shape => shape is PowerPointPicture));
            Assert.False(reopenedPicture.Hidden);
            Assert.Equal(editedBytes,
                reopened.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void NativeWriter_RejectsPicturesOnUnusedOrdinaryLayouts() {
            byte[] imageBytes = PdfPngTestImages.CreateRgbPng(52, 92, 132);
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            SlideMasterPart masterPart = presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single();
            int unusedLayoutIndex = presentation.GetLayoutIndex(
                P.SlideLayoutValues.Title);
            SlideLayoutPart unusedLayout = masterPart.SlideLayoutParts
                .ElementAt(unusedLayoutIndex);
            AddPictureShape(unusedLayout,
                unusedLayout.SlideLayout!.CommonSlideData!.ShapeTree!,
                imageBytes, 100U, 350000L);
            presentation.AddSlide(P.SlideLayoutValues.Blank);

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            LegacyPptWriteFinding finding = Assert.Single(
                preflight.Findings, item =>
                    item.Code == "PPT-WRITE-LAYOUT-PICTURE");
            Assert.Equal(LegacyPptFeature.Layouts, finding.Feature);
            Assert.Contains("does not materialize into any owning slide",
                finding.Description);
        }

        [Fact]
        public void NativeWriter_RejectsLayoutPicturesHiddenByOwningSlides() {
            byte[] imageBytes = PdfPngTestImages.CreateRgbPng(72, 112, 152);
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            SlideMasterPart masterPart = presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single();
            int layoutIndex = presentation.GetLayoutIndex(
                P.SlideLayoutValues.Blank);
            SlideLayoutPart layout = masterPart.SlideLayoutParts
                .ElementAt(layoutIndex);
            AddPictureShape(layout,
                layout.SlideLayout!.CommonSlideData!.ShapeTree!, imageBytes,
                100U, 350000L);
            PowerPointSlide slide = presentation.AddSlide(0, layoutIndex);
            slide.SlidePart.Slide!.ShowMasterShapes = false;

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-LAYOUT-PICTURE"
                && finding.Description.Contains("does not materialize",
                    StringComparison.Ordinal));
        }

        [Fact]
        public void NativeWriter_RejectsLayoutPicturesOverriddenOnAllSlides() {
            byte[] imageBytes = PdfPngTestImages.CreateRgbPng(92, 132, 172);
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            SlideMasterPart masterPart = presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single();
            int layoutIndex = presentation.GetLayoutIndex(
                P.SlideLayoutValues.Blank);
            SlideLayoutPart layout = masterPart.SlideLayoutParts
                .ElementAt(layoutIndex);
            P.Picture layoutPicture = AddPictureShape(layout,
                layout.SlideLayout!.CommonSlideData!.ShapeTree!, imageBytes,
                100U, 350000L);
            AddPicturePlaceholder(layoutPicture, 7U);
            PowerPointSlide slide = presentation.AddSlide(0, layoutIndex);
            using var image = new MemoryStream(imageBytes, writable: false);
            P.Picture slidePicture = (P.Picture)slide.AddPicture(image,
                OfficeIMO.PowerPoint.ImagePartType.Png,
                450000L, 300000L, 800000L, 600000L)
                .Element;
            AddPicturePlaceholder(slidePicture, 7U);

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-LAYOUT-PICTURE");
        }

        [Fact]
        public void NativeWriter_RejectsMediaPicturesOnMasters() {
            byte[] imageBytes = PdfPngTestImages.CreateRgbPng(22, 42, 62);
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            SlideMasterPart masterPart = presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single();
            P.Picture picture = AddPictureShape(masterPart,
                masterPart.SlideMaster!.CommonSlideData!.ShapeTree!,
                imageBytes, 100U, 350000L);
            picture.NonVisualPictureProperties!
                .ApplicationNonVisualDrawingProperties!
                .Append(new A.AudioFromFile { Link = "rIdMissingAudio" });
            presentation.AddSlide(P.SlideLayoutValues.Blank);

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-MASTER-SHAPE");
        }

        [Fact]
        public void NativeWriter_RejectsPictureLocksWithoutExactMapping() {
            byte[] imageBytes = PdfPngTestImages.CreateRgbPng(32, 72, 112);
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            SlideMasterPart masterPart = presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single();
            P.Picture picture = AddPictureShape(masterPart,
                masterPart.SlideMaster!.CommonSlideData!.ShapeTree!,
                imageBytes, 100U, 350000L);
            picture.NonVisualPictureProperties!
                .NonVisualPictureDrawingProperties!
                .GetFirstChild<A.PictureLocks>()!.NoResize = true;
            presentation.AddSlide(P.SlideLayoutValues.Blank);

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-PICTURE"
                && finding.Description.Contains("no-resize",
                    StringComparison.Ordinal));
        }

        private static P.Picture AddPictureShape(OpenXmlPart ownerPart,
            P.ShapeTree tree, byte[] imageBytes, uint id, long left) {
            ImagePart imagePart = ownerPart.AddNewPart<ImagePart>("image/png");
            using (var image = new MemoryStream(imageBytes, writable: false)) {
                imagePart.FeedData(image);
            }
            var picture = new P.Picture(
                new P.NonVisualPictureProperties(
                    new P.NonVisualDrawingProperties {
                        Id = id,
                        Name = $"Master picture {id}"
                    },
                    new P.NonVisualPictureDrawingProperties(
                        new A.PictureLocks {
                            NoGrouping = true,
                            NoSelection = true,
                            NoRotation = true,
                            NoChangeAspect = true,
                            NoMove = true,
                            NoEditPoints = true,
                            NoAdjustHandles = true,
                            NoCrop = true,
                            NoChangeShapeType = true
                        }) {
                        PreferRelativeResize = true
                    },
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.BlipFill(
                    new A.Blip { Embed = ownerPart.GetIdOfPart(imagePart) },
                    new A.Stretch(new A.FillRectangle())),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = left, Y = 300000L },
                        new A.Extents { Cx = 800000L, Cy = 600000L }),
                    new A.PresetGeometry(new A.AdjustValueList()) {
                        Preset = A.ShapeTypeValues.Rectangle
                    }));
            tree.Append(picture);
            return picture;
        }

        private static void AddPicturePlaceholder(P.Picture picture,
            uint index) {
            picture.NonVisualPictureProperties!
                .ApplicationNonVisualDrawingProperties!
                .Append(new P.PlaceholderShape {
                    Type = P.PlaceholderValues.Picture,
                    Index = index
                });
        }

        private static void AssertMasterPicture(
            IReadOnlyList<LegacyPptShape> shapes, byte[] imageBytes,
            bool expectedHidden = false) {
            LegacyPptShape picture = Assert.Single(shapes,
                shape => shape.Kind == LegacyPptShapeKind.Picture);
            Assert.Equal(imageBytes, picture.Picture!.ImageBytes);
            Assert.Equal(expectedHidden, picture.Style.Hidden ?? false);
            OfficeArtShapeProtection protection =
                OfficeArtShapeProtection.Decode(picture.Style.Properties);
            Assert.True(protection.LockAgainstGrouping);
            Assert.True(protection.LockAgainstSelect);
            Assert.True(protection.LockRotation);
            Assert.True(protection.LockAspectRatio);
            Assert.True(protection.LockPosition);
            Assert.True(protection.LockCropping);
            Assert.True(protection.LockVertices);
            Assert.True(protection.LockAdjustHandles);
            Assert.Null(protection.LockAgainstUngrouping);
            Assert.Null(protection.LockText);
            string propertyDump = string.Join(", ", picture.Style.Properties
                .Select(property => $"0x{property.PropertyId:X4}=0x{property.Value:X8}"));
            Assert.True(picture.Style.PreferRelativeResize, propertyDump);
            Assert.True(picture.Style.LockShapeType, propertyDump);
        }

        private static void AssertProjectedMasterPicture(
            IReadOnlyList<PowerPointShape> shapes, byte[] imageBytes,
            bool expectedHidden = false) {
            PowerPointPicture picture = Assert.IsType<PowerPointPicture>(
                Assert.Single(shapes, shape => shape is PowerPointPicture));
            Assert.Equal(imageBytes, picture.GetImageBytes());
            Assert.Equal(expectedHidden, picture.Hidden);
            A.PictureLocks locks = Assert.IsType<A.PictureLocks>(
                ((P.Picture)picture.Element).NonVisualPictureProperties?
                    .NonVisualPictureDrawingProperties?.FirstChild);
            Assert.True(locks.NoGrouping);
            Assert.True(locks.NoSelection);
            Assert.True(locks.NoRotation);
            Assert.True(locks.NoChangeAspect);
            Assert.True(locks.NoMove);
            Assert.True(locks.NoEditPoints);
            Assert.True(locks.NoAdjustHandles);
            Assert.True(locks.NoCrop);
            Assert.True(locks.NoChangeShapeType);
            Assert.True(((P.Picture)picture.Element)
                .NonVisualPictureProperties?
                .NonVisualPictureDrawingProperties?
                .PreferRelativeResize);
        }
    }
}
