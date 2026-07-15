using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        private static string AccessibilityFixturePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "AccessibilityPowerPoint.ppt");

        [Fact]
        public void CurrentUserAtomReader_AcceptsMicrosoftFourBytePayloadOverstatement() {
            byte[] stream = {
                0x00, 0x00, 0xF6, 0x0F, 0x1C, 0x00, 0x00, 0x00,
                0x14, 0x00, 0x00, 0x00, 0x5F, 0xC0, 0x91, 0xE3,
                0xDE, 0x9D, 0x00, 0x00, 0x00, 0x00, 0xF4, 0x03,
                0x03, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00
            };

            LegacyPptCurrentUserAtom currentUser = LegacyPptCurrentUserAtom.Read(stream);

            Assert.Equal(0xE391C05FU, currentUser.HeaderToken);
            Assert.Equal(0x00009DDEU, currentUser.CurrentEditOffset);
            Assert.True(currentUser.HasFourBytePayloadOverstatement);
        }

        [Fact]
        public void CurrentUserAtomReader_RejectsOtherTruncatedPayloads() {
            byte[] stream = {
                0x00, 0x00, 0xF6, 0x0F, 0x1D, 0x00, 0x00, 0x00,
                0x14, 0x00, 0x00, 0x00, 0x5F, 0xC0, 0x91, 0xE3,
                0xDE, 0x9D, 0x00, 0x00, 0x00, 0x00, 0xF4, 0x03,
                0x03, 0x00, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00
            };

            Assert.Throws<InvalidDataException>(() => LegacyPptCurrentUserAtom.Read(stream));
        }

        [Fact]
        public void LegacyShapeFactory_ProjectsGroupAndChildAccessibilityMetadata() {
            OfficeArtShapeStyle groupStyle = OfficeArtShapeStyle.Decode(new[] {
                new OfficeArtProperty(0, 0x8380, 24U, complexText: "Named Group"),
                new OfficeArtProperty(1, 0x8381, 36U, complexText: "Group description")
            });
            OfficeArtShapeStyle childStyle = OfficeArtShapeStyle.Decode(new[] {
                new OfficeArtProperty(0, 0x8380, 24U, complexText: "Named Child"),
                new OfficeArtProperty(1, 0x8381, 36U, complexText: "Child description")
            });
            var child = new LegacyPptShape(LegacyPptShapeKind.Rectangle, 1, 20, 0,
                new LegacyPptBounds(0, 0, 100, 100), string.Empty,
                placeholder: null, childStyle, null, null);
            var source = new LegacyPptShape(LegacyPptShapeKind.Group, 0, 10, 0,
                new LegacyPptBounds(100, 200, 400, 300), string.Empty,
                placeholder: null, groupStyle, null, null,
                groupCoordinateBounds: new LegacyPptBounds(0, 0, 400, 300),
                children: new[] { child });
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            uint nextShapeId = 2;

            P.GroupShape group = Assert.IsType<P.GroupShape>(
                PowerPointPresentation.CreateLegacyOpenXmlShape(slide.SlidePart, source,
                    ref nextShapeId));
            P.Shape projectedChild = Assert.Single(group.Elements<P.Shape>());

            Assert.Equal("Named Group", group.NonVisualGroupShapeProperties!
                .NonVisualDrawingProperties!.Name!.Value);
            Assert.Equal("Group description", group.NonVisualGroupShapeProperties
                .NonVisualDrawingProperties!.Description!.Value);
            Assert.Equal("Named Child", projectedChild.NonVisualShapeProperties!
                .NonVisualDrawingProperties!.Name!.Value);
            Assert.Equal("Child description", projectedChild.NonVisualShapeProperties
                .NonVisualDrawingProperties!.Description!.Value);
        }

        [Fact]
        public void NeutralReader_DecodesMicrosoftObjectNamesAndDescriptions() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(AccessibilityFixturePath);
            LegacyPptSlide slide = Assert.Single(legacy.Slides);

            LegacyPptShape summary = Assert.Single(slide.Shapes, shape => shape.Text == "Summary");
            Assert.Equal("Quarter Summary", summary.Metadata.Name);
            Assert.Equal("Rounded summary of quarter performance", summary.Metadata.Description);

            LegacyPptShape picture = Assert.Single(slide.Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Picture);
            Assert.Equal("Four Colors", picture.Metadata.Name);
            Assert.Equal("Blue field with a yellow circle", picture.Metadata.Description);

            LegacyPptShape connector = Assert.Single(slide.Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Connector);
            Assert.Equal("Trend Link", connector.Metadata.Name);
            Assert.Equal("Connector from summary to picture", connector.Metadata.Description);
            Assert.DoesNotContain(legacy.Diagnostics,
                diagnostic => diagnostic.Code.EndsWith("-TRUNCATED", StringComparison.Ordinal));
        }

        [Fact]
        public void NormalLoad_ProjectsAccessibilityMetadataAndPreservesBinaryExactly() {
            byte[] source = File.ReadAllBytes(AccessibilityFixturePath);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(
                AccessibilityFixturePath);
            PowerPointSlide slide = Assert.Single(presentation.Slides);

            PowerPointTextBox summary = Assert.Single(slide.TextBoxes,
                shape => shape.Text == "Summary");
            Assert.Equal("Quarter Summary", summary.Name);
            Assert.Equal("Rounded summary of quarter performance", summary.Description);

            PowerPointPicture picture = Assert.Single(slide.Pictures);
            Assert.Equal("Four Colors", picture.Name);
            Assert.Equal("Blue field with a yellow circle", picture.Description);

            PowerPointConnectionShape connector = Assert.Single(
                slide.Shapes.OfType<PowerPointConnectionShape>());
            Assert.Equal("Trend Link", connector.Name);
            Assert.Equal("Connector from summary to picture", connector.Description);
            Assert.Empty(presentation.ValidateDocument());
            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            Assert.Equal(source, presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ImportedMicrosoftGeometryEdit_PreservesAccessibilityMetadata() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(AccessibilityFixturePath);
            LegacyPptShape originalSummary = original.Slides[0].Shapes.Single(shape =>
                shape.Text == "Summary");
            using PowerPointPresentation presentation = PowerPointPresentation.Load(
                AccessibilityFixturePath);
            PowerPointTextBox summary = presentation.Slides[0].TextBoxes.Single(shape =>
                shape.Text == "Summary");

            summary.Left += 15875;

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptShape savedSummary = saved.Slides[0].Shapes.Single(shape =>
                shape.Text == "Summary");
            Assert.Equal(originalSummary.Bounds.Left + 10, savedSummary.Bounds.Left);
            Assert.Equal("Quarter Summary", savedSummary.Metadata.Name);
            Assert.Equal("Rounded summary of quarter performance",
                savedSummary.Metadata.Description);
            Assert.Equal(original.Package.UserEdits.Count + 1, saved.Package.UserEdits.Count);
        }

        [Fact]
        public void ImportedAccessibilityMetadataEdit_RemainsLossBlocked() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(
                AccessibilityFixturePath);
            PowerPointPicture picture = Assert.Single(presentation.Slides[0].Pictures);

            picture.Description = "Edited alternative text";

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings,
                finding => finding.Code == "PPT-WRITE-IMPORT-LOSS");
        }
    }
}
