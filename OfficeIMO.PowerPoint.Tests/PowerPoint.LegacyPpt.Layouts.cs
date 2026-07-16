using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        [Fact]
        public void LayoutReader_RecognizesEveryDefinedBinaryLayoutType() {
            foreach (LegacyPptSlideLayoutType expected in Enum.GetValues(
                         typeof(LegacyPptSlideLayoutType))) {
                var payload = new byte[24];
                System.Buffers.Binary.BinaryPrimitives.WriteUInt32LittleEndian(
                    payload.AsSpan(0, 4), (uint)expected);
                var record = new LegacyPptRecord(payload, 0, 2, 0, 0x03EF,
                    0, payload.Length);

                Assert.True(LegacyPptLayoutReader.TryReadSlideAtom(record,
                    out LegacyPptSlideAtomData actual));
                Assert.Equal(expected, actual.Layout);
            }
        }

        [Fact]
        public void LayoutReader_DecodesLayoutSignatureAndPlaceholderIdentity() {
            var slidePayload = new byte[24];
            slidePayload[0] = (byte)LegacyPptSlideLayoutType.TwoColumns;
            slidePayload[4] = (byte)LegacyPptPlaceholderKind.Title;
            slidePayload[5] = (byte)LegacyPptPlaceholderKind.Body;
            slidePayload[6] = (byte)LegacyPptPlaceholderKind.Graph;
            System.Buffers.Binary.BinaryPrimitives.WriteUInt32LittleEndian(
                slidePayload.AsSpan(12, 4), 0x12345678U);
            System.Buffers.Binary.BinaryPrimitives.WriteUInt32LittleEndian(
                slidePayload.AsSpan(16, 4), 0x87654321U);
            slidePayload[20] = 0x07;
            var slideAtom = new LegacyPptRecord(slidePayload, 0, 2, 0, 0x03EF,
                0, slidePayload.Length);

            Assert.True(LegacyPptLayoutReader.TryReadSlideAtom(slideAtom,
                out LegacyPptSlideAtomData slide));
            Assert.Equal(LegacyPptSlideLayoutType.TwoColumns, slide.Layout);
            Assert.Equal(0x12345678U, slide.MasterId);
            Assert.Equal(0x87654321U, slide.NotesId);
            Assert.True(slide.FollowsMasterObjects);
            Assert.True(slide.FollowsMasterColorScheme);
            Assert.True(slide.FollowsMasterBackground);
            Assert.Equal(LegacyPptPlaceholderKind.Graph, slide.PlaceholderTypes[2]);

            var placeholderPayload = new byte[8];
            System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(
                placeholderPayload.AsSpan(0, 4), 6);
            placeholderPayload[4] = (byte)LegacyPptPlaceholderKind.VerticalBody;
            placeholderPayload[5] = (byte)LegacyPptPlaceholderSize.Quarter;
            var placeholderAtom = new LegacyPptRecord(placeholderPayload, 0, 0, 0,
                0x0BC3, 0, placeholderPayload.Length);

            LegacyPptPlaceholder placeholder = Assert.IsType<LegacyPptPlaceholder>(
                LegacyPptLayoutReader.ReadPlaceholder(placeholderAtom,
                    out LegacyPptPlaceholderReadStatus status));
            Assert.Equal(LegacyPptPlaceholderReadStatus.Decoded, status);
            Assert.Equal(6, placeholder.Position);
            Assert.Equal(LegacyPptPlaceholderKind.VerticalBody, placeholder.Kind);
            Assert.Equal(LegacyPptPlaceholderSize.Quarter, placeholder.Size);
        }

        [Fact]
        public void LayoutReader_BoundsInvalidPlaceholderRecords() {
            var sentinelPayload = new byte[8];
            System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(
                sentinelPayload.AsSpan(0, 4), -1);
            sentinelPayload[4] = (byte)LegacyPptPlaceholderKind.Title;
            var sentinel = new LegacyPptRecord(sentinelPayload, 0, 0, 0, 0x0BC3,
                0, sentinelPayload.Length);
            Assert.Null(LegacyPptLayoutReader.ReadPlaceholder(sentinel,
                out LegacyPptPlaceholderReadStatus sentinelStatus));
            Assert.Equal(LegacyPptPlaceholderReadStatus.NotPlaceholder, sentinelStatus);

            byte[] invalidPayload = { 0, 0, 0, 0, 0xFF, 0x03, 0, 0 };
            var invalid = new LegacyPptRecord(invalidPayload, 0, 0, 0, 0x0BC3,
                0, invalidPayload.Length);
            Assert.Null(LegacyPptLayoutReader.ReadPlaceholder(invalid,
                out LegacyPptPlaceholderReadStatus invalidStatus));
            Assert.Equal(LegacyPptPlaceholderReadStatus.Invalid, invalidStatus);

            var invalidSlidePayload = new byte[25];
            invalidSlidePayload[0] = 0x03;
            invalidSlidePayload[4] = 0xFF;
            invalidSlidePayload[20] = 0xF8;
            var invalidSlide = new LegacyPptRecord(invalidSlidePayload, 0, 2, 0,
                0x03EF, 0, invalidSlidePayload.Length);
            Assert.True(LegacyPptLayoutReader.TryReadSlideAtom(invalidSlide,
                out LegacyPptSlideAtomData invalidSlideData));
            Assert.Null(invalidSlideData.Layout);
            Assert.True(invalidSlideData.HasInvalidPlaceholderType);
            Assert.True(invalidSlideData.HasReservedFlags);
            Assert.True(invalidSlideData.HasInvalidLength);
        }

        [Fact]
        public void RealCorpus_LayoutAndPlaceholderRecordsDecodeWithoutStructuralWarnings() {
            string corpus = Path.Combine(AppContext.BaseDirectory, "Documents",
                "LegacyPptCorpus");
            foreach (string path in Directory.GetFiles(corpus, "*.ppt")) {
                LegacyPptPresentation presentation = LegacyPptPresentation.Load(path);
                Assert.All(presentation.Slides, slide =>
                    Assert.Equal(8, slide.LayoutPlaceholderTypes.Count));
                Assert.All(presentation.Slides.SelectMany(slide => slide.Shapes)
                    .Where(shape => shape.Placeholder != null), shape =>
                    Assert.True(shape.Placeholder!.Position >= 0));
                Assert.DoesNotContain(presentation.Diagnostics, diagnostic =>
                    diagnostic.Code == "PPT-SLIDE-ATOM-TRUNCATED"
                    || diagnostic.Code == "PPT-MASTER-SLIDE-ATOM-TRUNCATED"
                    || diagnostic.Code == "PPT-SLIDE-ATOM-LENGTH"
                    || diagnostic.Code == "PPT-SLIDE-LAYOUT-TYPE"
                    || diagnostic.Code == "PPT-SLIDE-LAYOUT-PLACEHOLDER"
                    || diagnostic.Code == "PPT-SLIDE-FLAGS-RESERVED"
                    || diagnostic.Code == "PPT-PLACEHOLDER-INVALID");
            }
        }

        [Fact]
        public void NativeWriter_RoundTripsLayoutAndPlaceholderIdentity() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide(P.SlideLayoutValues.Text);
            PowerPointTextBox title = slide.AddTitle("Layout title", 500000, 300000,
                7000000, 800000);
            title.PlaceholderIndex = 0;
            PowerPointTextBox body = slide.AddTextBox("Vertical body", 800000, 1500000,
                5000000, 3000000);
            body.PlaceholderType = P.PlaceholderValues.Body;
            body.PlaceholderIndex = 1;
            body.PlaceholderSize = P.PlaceholderSizeValues.Quarter;
            body.PlaceholderOrientation = P.DirectionValues.Vertical;

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);

            LegacyPptSlide binary = Assert.Single(LegacyPptPresentation.Load(bytes).Slides);
            Assert.Equal(LegacyPptSlideLayoutType.TitleBody, binary.Layout);
            Assert.Equal(LegacyPptPlaceholderKind.Title, binary.LayoutPlaceholderTypes[0]);
            Assert.Equal(LegacyPptPlaceholderKind.VerticalBody,
                binary.LayoutPlaceholderTypes[1]);
            LegacyPptPlaceholder binaryTitle = Assert.IsType<LegacyPptPlaceholder>(
                binary.Shapes.Single(shape => shape.Text == "Layout title").Placeholder);
            LegacyPptPlaceholder binaryBody = Assert.IsType<LegacyPptPlaceholder>(
                binary.Shapes.Single(shape => shape.Text == "Vertical body").Placeholder);
            Assert.Equal(0, binaryTitle.Position);
            Assert.Equal(LegacyPptPlaceholderSize.Full, binaryTitle.Size);
            Assert.Equal(1, binaryBody.Position);
            Assert.Equal(LegacyPptPlaceholderSize.Quarter, binaryBody.Size);

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened = PowerPointPresentation.Load(stream);
            PowerPointTextBox reopenedBody = Assert.Single(reopened.Slides[0].TextBoxes,
                textBox => textBox.Text == "Vertical body");
            Assert.Equal(P.PlaceholderValues.Body, reopenedBody.ShapePlaceholderType);
            Assert.Equal(1U, reopenedBody.ShapePlaceholderIndex);
            Assert.Equal(P.PlaceholderSizeValues.Quarter,
                reopenedBody.ShapePlaceholderSize);
            Assert.Equal(P.DirectionValues.Vertical,
                reopenedBody.ShapePlaceholderOrientation);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_ProjectsDistinctBinaryLayoutsUnderSharedMaster() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide titleSlide = presentation.AddSlide(P.SlideLayoutValues.Title);
            PowerPointTextBox centeredTitle = titleSlide.AddTextBox("Centered", 500000,
                500000, 7000000, 900000);
            centeredTitle.PlaceholderType = P.PlaceholderValues.CenteredTitle;
            centeredTitle.PlaceholderIndex = 0;
            PowerPointTextBox subtitle = titleSlide.AddTextBox("Subtitle", 500000,
                1800000, 7000000, 900000);
            subtitle.PlaceholderType = P.PlaceholderValues.SubTitle;
            subtitle.PlaceholderIndex = 1;

            PowerPointSlide textSlide = presentation.AddSlide(P.SlideLayoutValues.Text);
            PowerPointTextBox title = textSlide.AddTitle("Title and body", 500000,
                300000, 7000000, 800000);
            title.PlaceholderIndex = 0;
            PowerPointTextBox body = textSlide.AddTextBox("Body", 700000, 1500000,
                6000000, 3000000);
            body.PlaceholderType = P.PlaceholderValues.Body;
            body.PlaceholderIndex = 1;

            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);
            Assert.Equal(binary.Slides[0].MasterId, binary.Slides[1].MasterId);
            Assert.Equal(LegacyPptSlideLayoutType.TitleSlide, binary.Slides[0].Layout);
            Assert.Equal(LegacyPptSlideLayoutType.TitleBody, binary.Slides[1].Layout);

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation projected = PowerPointPresentation.Load(stream);
            Assert.Equal(P.SlideLayoutValues.Title,
                projected.Slides[0].SlidePart.SlideLayoutPart!.SlideLayout!.Type!.Value);
            Assert.Equal(P.SlideLayoutValues.Text,
                projected.Slides[1].SlidePart.SlideLayoutPart!.SlideLayout!.Type!.Value);
            Assert.NotEqual(projected.Slides[0].SlidePart.SlideLayoutPart!.Uri,
                projected.Slides[1].SlidePart.SlideLayoutPart!.Uri);
            Assert.Same(projected.Slides[0].SlidePart.SlideLayoutPart!.SlideMasterPart,
                projected.Slides[1].SlidePart.SlideLayoutPart!.SlideMasterPart);
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_MaterializesLayoutPlaceholderGeometry() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            int layoutIndex = presentation.GetLayoutIndex(P.SlideLayoutValues.Text);
            var titleBounds = new PowerPointLayoutBox(500000, 250000,
                7200000, 900000);
            var bodyBounds = new PowerPointLayoutBox(850000, 1450000,
                6500000, 4100000);
            presentation.SetLayoutPlaceholderBounds(0, layoutIndex,
                P.PlaceholderValues.Title, titleBounds, index: 0);
            presentation.SetLayoutPlaceholderBounds(0, layoutIndex,
                P.PlaceholderValues.Body, bodyBounds, index: 1);
            PowerPointSlide slide = presentation.AddSlide(0, layoutIndex);
            Assert.Empty(slide.Shapes);

            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptSlide binary = Assert.Single(
                LegacyPptPresentation.Load(bytes).Slides);
            Assert.Equal(2, binary.Shapes.Count);
            LegacyPptShape title = binary.Shapes.Single(shape =>
                shape.Placeholder?.Kind == LegacyPptPlaceholderKind.Title);
            LegacyPptShape body = binary.Shapes.Single(shape =>
                shape.Placeholder?.Kind == LegacyPptPlaceholderKind.Body);
            Assert.Equal(LayoutToMasterUnits(titleBounds.Left), title.Bounds.Left);
            Assert.Equal(LayoutToMasterUnits(titleBounds.Top), title.Bounds.Top);
            Assert.Equal(LayoutToMasterUnits(titleBounds.Width), title.Bounds.Width);
            Assert.Equal(LayoutToMasterUnits(titleBounds.Height), title.Bounds.Height);
            Assert.Equal(LayoutToMasterUnits(bodyBounds.Left), body.Bounds.Left);
            Assert.Equal(LayoutToMasterUnits(bodyBounds.Top), body.Bounds.Top);
            Assert.Equal(LayoutToMasterUnits(bodyBounds.Width), body.Bounds.Width);
            Assert.Equal(LayoutToMasterUnits(bodyBounds.Height), body.Bounds.Height);

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened = PowerPointPresentation.Load(stream);
            PowerPointSlide projected = Assert.Single(reopened.Slides);
            Assert.Equal(2, projected.Shapes.Count);
            PowerPointTextBox projectedTitle = Assert.IsType<PowerPointTextBox>(
                projected.Shapes.Single(shape =>
                    shape.ShapePlaceholderType == P.PlaceholderValues.Title));
            PowerPointTextBox projectedBody = Assert.IsType<PowerPointTextBox>(
                projected.Shapes.Single(shape =>
                    shape.ShapePlaceholderType == P.PlaceholderValues.Body));
            Assert.Equal(LayoutToEmus(title.Bounds.Left), projectedTitle.Left);
            Assert.Equal(LayoutToEmus(title.Bounds.Top), projectedTitle.Top);
            Assert.Equal(LayoutToEmus(body.Bounds.Left), projectedBody.Left);
            Assert.Equal(LayoutToEmus(body.Bounds.Top), projectedBody.Top);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_PrefersSlidePlaceholderOverLayoutPlaceholder() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            int layoutIndex = presentation.GetLayoutIndex(P.SlideLayoutValues.Text);
            var layoutTitleBounds = new PowerPointLayoutBox(500000, 250000,
                7200000, 900000);
            var slideTitleBounds = new PowerPointLayoutBox(900000, 450000,
                6400000, 700000);
            presentation.SetLayoutPlaceholderBounds(0, layoutIndex,
                P.PlaceholderValues.Title, layoutTitleBounds, index: 0);
            PowerPointSlide slide = presentation.AddSlide(0, layoutIndex);
            PowerPointTextBox title = slide.AddTitle("Slide override",
                slideTitleBounds.Left, slideTitleBounds.Top,
                slideTitleBounds.Width, slideTitleBounds.Height);
            title.PlaceholderIndex = 0;

            LegacyPptSlide binary = Assert.Single(LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt)).Slides);

            Assert.Equal(2, binary.Shapes.Count);
            LegacyPptShape binaryTitle = Assert.Single(binary.Shapes, shape =>
                shape.Placeholder?.Kind == LegacyPptPlaceholderKind.Title);
            Assert.Equal("Slide override", binaryTitle.Text);
            Assert.Equal(LayoutToMasterUnits(slideTitleBounds.Left),
                binaryTitle.Bounds.Left);
            Assert.DoesNotContain(binary.Shapes, shape =>
                shape.Placeholder?.Kind == LegacyPptPlaceholderKind.Title
                && shape.Bounds.Left == LayoutToMasterUnits(layoutTitleBounds.Left));
            Assert.Single(binary.Shapes, shape =>
                shape.Placeholder?.Kind == LegacyPptPlaceholderKind.Body);
        }

        [Fact]
        public void NativeWriter_MaterializesLayoutDecorationAndHonorsVisibility() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            int layoutIndex = presentation.GetLayoutIndex(P.SlideLayoutValues.Text);
            PowerPointSlide slide = presentation.AddSlide(0, layoutIndex);
            var decorationBounds = new PowerPointLayoutBox(200000, 300000,
                1400000, 240000);
            P.ShapeTree tree = slide.SlidePart.SlideLayoutPart!.SlideLayout!
                .CommonSlideData!.ShapeTree!;
            tree.Append(new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties {
                        Id = 701U,
                        Name = "Layout decoration"
                    },
                    new P.NonVisualShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset {
                            X = decorationBounds.Left,
                            Y = decorationBounds.Top
                        },
                        new A.Extents {
                            Cx = decorationBounds.Width,
                            Cy = decorationBounds.Height
                        }),
                    new A.PresetGeometry(new A.AdjustValueList()) {
                        Preset = A.ShapeTypeValues.Rectangle
                    })));

            LegacyPptSlide visible = Assert.Single(LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt)).Slides);
            LegacyPptShape decoration = Assert.Single(visible.Shapes,
                shape => shape.Placeholder == null);
            Assert.Equal(LayoutToMasterUnits(decorationBounds.Left),
                decoration.Bounds.Left);
            Assert.Equal(LayoutToMasterUnits(decorationBounds.Top),
                decoration.Bounds.Top);

            slide.SlidePart.Slide!.ShowMasterShapes = false;
            byte[] hiddenBytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptSlide hidden = Assert.Single(
                LegacyPptPresentation.Load(hiddenBytes).Slides);
            Assert.Empty(hidden.Shapes);
            Assert.False(hidden.FollowsMasterObjects);

            using var hiddenStream = new MemoryStream(hiddenBytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(hiddenStream);
            Assert.False(reopened.Slides[0].SlidePart.Slide!
                .ShowMasterShapes!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedSlideMasterShapeVisibilityEdit_AppendsPreservingRecord() {
            byte[] sourceBytes;
            using (PowerPointPresentation created =
                   PowerPointPresentation.Create()) {
                created.AddSlide(P.SlideLayoutValues.Blank);
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original =
                LegacyPptPresentation.Load(sourceBytes);

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            imported.Slides[0].SlidePart.Slide!.ShowMasterShapes = false;

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved =
                LegacyPptPresentation.Load(savedBytes);

            Assert.False(Assert.Single(saved.Slides).FollowsMasterObjects);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            using var reopenedInput = new MemoryStream(savedBytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            Assert.False(reopened.Slides[0].SlidePart.Slide!
                .ShowMasterShapes!.Value);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedLayoutRelationshipEdit_RemainsLossBlocked() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);
            PowerPointSlide slide = Assert.Single(presentation.Slides);
            slide.SlidePart.SlideLayoutPart!.SlideLayout!.Type = P.SlideLayoutValues.Blank;

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings,
                finding => finding.Code == "PPT-WRITE-IMPORT-LOSS");
        }

        [Fact]
        public void ImportedSlidePlaceholderContractEdit_AppendsPreservingRecord() {
            LegacyPptPresentation original =
                LegacyPptPresentation.Load(FixturePath);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(FixturePath);
            PowerPointTextBox title = Assert.Single(imported.Slides[0]
                .TextBoxes, textBox =>
                    textBox.Text == "OfficeIMO PowerPoint Basics");
            title.PlaceholderType = P.PlaceholderValues.Body;
            title.PlaceholderIndex = 7;
            title.PlaceholderSize = P.PlaceholderSizeValues.Quarter;
            title.PlaceholderOrientation = P.DirectionValues.Vertical;

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            byte[] savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved =
                LegacyPptPresentation.Load(savedBytes);
            LegacyPptSlide savedSlide = Assert.Single(saved.Slides);
            LegacyPptPlaceholder placeholder = Assert.IsType<
                LegacyPptPlaceholder>(Assert.Single(savedSlide.Shapes,
                    shape => shape.Text == "OfficeIMO PowerPoint Basics")
                    .Placeholder);

            Assert.Equal(7, placeholder.Position);
            Assert.Equal(LegacyPptPlaceholderKind.VerticalBody,
                placeholder.Kind);
            Assert.Equal(LegacyPptPlaceholderSize.Quarter,
                placeholder.Size);
            Assert.Equal(LegacyPptPlaceholderKind.VerticalBody,
                savedSlide.LayoutPlaceholderTypes[7]);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            using var reopenedInput = new MemoryStream(savedBytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            PowerPointTextBox reopenedTitle = Assert.Single(
                reopened.Slides[0].TextBoxes, textBox =>
                    textBox.Text == "OfficeIMO PowerPoint Basics");
            Assert.Equal(P.PlaceholderValues.Body,
                reopenedTitle.PlaceholderType);
            Assert.Equal(7U, reopenedTitle.PlaceholderIndex);
            Assert.Equal(P.PlaceholderSizeValues.Quarter,
                reopenedTitle.PlaceholderSize);
            Assert.Equal(P.DirectionValues.Vertical,
                reopenedTitle.PlaceholderOrientation);
            Assert.Empty(reopened.ValidateDocument());
        }

        private static int LayoutToMasterUnits(long emus) => checked((int)Math.Round(
            emus / 1587.5D, MidpointRounding.AwayFromZero));

        private static long LayoutToEmus(int masterUnits) => checked((long)Math.Round(
            masterUnits * 1587.5D, MidpointRounding.AwayFromZero));
    }
}
