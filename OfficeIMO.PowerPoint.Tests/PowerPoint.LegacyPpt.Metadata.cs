using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using OfficeIMO.Tests.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
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
            Assert.True(summary.Metadata.CanRewrite);

            LegacyPptShape picture = Assert.Single(slide.Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Picture);
            Assert.Equal("Four Colors", picture.Metadata.Name);
            Assert.Equal("Blue field with a yellow circle", picture.Metadata.Description);
            Assert.True(picture.Metadata.CanRewrite);

            LegacyPptShape connector = Assert.Single(slide.Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Connector);
            Assert.Equal("Trend Link", connector.Metadata.Name);
            Assert.Equal("Connector from summary to picture", connector.Metadata.Description);
            Assert.True(connector.Metadata.CanRewrite);
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

        [Theory]
        [InlineData("text")]
        [InlineData("picture")]
        [InlineData("connector")]
        public void ImportedAccessibilityMetadataEditsRoundTripIncrementally(
            string shapeKind) {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(
                AccessibilityFixturePath);
            PowerPointShape target = shapeKind switch {
                "text" => Assert.Single(presentation.Slides[0].TextBoxes,
                    shape => shape.Text == "Summary"),
                "picture" => Assert.Single(presentation.Slides[0].Pictures),
                "connector" => Assert.Single(presentation.Slides[0].Shapes
                    .OfType<PowerPointConnectionShape>()),
                _ => throw new ArgumentOutOfRangeException(nameof(shapeKind))
            };
            string editedName = $"Edited {shapeKind} name";
            string editedDescription = $"Edited {shapeKind} alternative text";
            target.Name = editedName;
            target.Description = editedDescription;

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            LegacyPptPresentation reopened = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            AssertShapeMetadata(reopened.Slides[0].Shapes, editedName,
                editedDescription);
            Assert.Equal(2, reopened.Package.UserEdits.Count);
        }

        [Fact]
        public void ImportedAccessibilityDescriptionCanBeRemoved() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(
                AccessibilityFixturePath);
            PowerPointPicture picture = Assert.Single(
                presentation.Slides[0].Pictures);

            picture.Description = null;

            LegacyPptWritePreflightReport preflight =
                presentation.AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            LegacyPptShape reopened = Assert.Single(
                LegacyPptPresentation.Load(presentation.ToBytes(
                        PowerPointFileFormat.Ppt)).Slides[0].Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Picture);
            Assert.Equal("Four Colors", reopened.Metadata.Name);
            Assert.Null(reopened.Metadata.Description);
        }

        [Fact]
        public void ImportedMainMasterAccessibilityMetadataEditIsPreserved() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                AccessibilityFixturePath);
            LegacyPptMaster[] mainMasters = original.Masters
                .Where(master => master.IsMainMaster).ToArray();
            int masterIndex = Array.FindIndex(mainMasters,
                master => master.Shapes.Count > 0);
            Assert.True(masterIndex >= 0);
            LegacyPptMaster originalMaster = mainMasters[masterIndex];
            using PowerPointPresentation presentation = PowerPointPresentation.Load(
                AccessibilityFixturePath);
            DocumentFormat.OpenXml.Packaging.SlideMasterPart masterPart =
                presentation.OpenXmlDocument.PresentationPart!
                    .SlideMasterParts.ElementAt(masterIndex);
            PowerPointShape target = LegacyPptWriter
                .ReadMasterShapesForWrite(masterPart, out string? reason)[0];
            Assert.Null(reason);

            target.Name = "Edited master object";
            target.Description = "Edited master alternative text";

            LegacyPptWritePreflightReport preflight =
                presentation.AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            LegacyPptPresentation reopened = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptMaster reopenedMaster = Assert.Single(reopened.Masters,
                master => master.MasterId == originalMaster.MasterId);
            AssertShapeMetadata(reopenedMaster.Shapes,
                "Edited master object", "Edited master alternative text");
            Assert.Equal(original.Package.UserEdits.Count + 1,
                reopened.Package.UserEdits.Count);
        }

        [Fact]
        public void NativeWriter_AuthorsShapeNamesAndAlternativeText() {
            byte[] image = PdfPngTestImages.CreateRgbPng(30, 90, 180);
            byte[] bytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                PowerPointTextBox text = slide.AddTextBoxPoints(
                    "Accessible text", 20, 20, 140, 35);
                text.Name = "Summary object";
                text.Description = "Summary alternative text";
                PowerPointPicture picture = slide.AddPicture(
                    new MemoryStream(image), ImagePartType.Png,
                    PowerPointUnits.FromPoints(20),
                    PowerPointUnits.FromPoints(70),
                    PowerPointUnits.FromPoints(90),
                    PowerPointUnits.FromPoints(60));
                picture.Name = "Chart preview";
                picture.Description = "Blue chart preview";
                PowerPointConnectionShape connector = slide
                    .AddConnectionShape(A.ShapeTypeValues.StraightConnector1,
                        PowerPointUnits.FromPoints(25),
                        PowerPointUnits.FromPoints(145),
                        PowerPointUnits.FromPoints(170),
                        PowerPointUnits.FromPoints(145));
                connector.Name = "Trend connector";
                connector.Description = "Connects the summary to detail";
                PowerPointAutoShape child = slide.AddShapePoints(
                    A.ShapeTypeValues.Rectangle, 190, 20, 70, 45);
                child.Name = "Grouped child";
                child.Description = "Child description";
                PowerPointAutoShape secondChild = slide.AddShapePoints(
                    A.ShapeTypeValues.Ellipse, 205, 75, 45, 35);
                secondChild.Name = "Second grouped child";
                secondChild.Description = "Second child description";
                PowerPointGroupShape group = slide.GroupShapes(
                    new PowerPointShape[] { child, secondChild },
                    "Accessible group");
                group.Name = "Accessible group";
                group.Description = "Group description";

                LegacyPptWritePreflightReport preflight =
                    source.AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptSlide binary = Assert.Single(
                LegacyPptPresentation.Load(bytes).Slides);
            AssertShapeMetadata(binary.Shapes, "Summary object",
                "Summary alternative text");
            AssertShapeMetadata(binary.Shapes, "Chart preview",
                "Blue chart preview");
            AssertShapeMetadata(binary.Shapes, "Trend connector",
                "Connects the summary to detail");
            LegacyPptShape binaryGroup = Assert.Single(binary.Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Group);
            Assert.Equal("Accessible group", binaryGroup.Metadata.Name);
            Assert.Equal("Group description",
                binaryGroup.Metadata.Description);
            AssertShapeMetadata(binaryGroup.Children, "Grouped child",
                "Child description");
            AssertShapeMetadata(binaryGroup.Children,
                "Second grouped child", "Second child description");

            using var input = new MemoryStream(bytes);
            using PowerPointPresentation projected =
                PowerPointPresentation.Load(input);
            Assert.Contains(projected.Slides[0].Shapes,
                shape => shape.Name == "Chart preview"
                    && shape.Description == "Blue chart preview");
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_BlocksModernOnlyAccessibilityFlags() {
            using PowerPointPresentation source =
                PowerPointPresentation.Create();
            PowerPointAutoShape shape = source.AddSlide(
                    P.SlideLayoutValues.Blank)
                .AddShapePoints(A.ShapeTypeValues.Rectangle,
                    20, 20, 100, 60);
            shape.Title = "Concise modern title";

            LegacyPptWritePreflightReport titlePreflight =
                source.AnalyzeLegacyPptWrite();
            Assert.Contains(titlePreflight.Findings, finding =>
                finding.Code == "PPT-WRITE-ACCESSIBILITY-METADATA"
                && finding.Description.Contains("titles",
                    StringComparison.OrdinalIgnoreCase));

            shape.Title = null;
            shape.Decorative = true;
            LegacyPptWritePreflightReport decorativePreflight =
                source.AnalyzeLegacyPptWrite();
            Assert.Contains(decorativePreflight.Findings, finding =>
                finding.Code == "PPT-WRITE-ACCESSIBILITY-METADATA"
                && finding.Description.Contains("decorative",
                    StringComparison.OrdinalIgnoreCase));
        }

        private static void AssertShapeMetadata(
            IEnumerable<LegacyPptShape> shapes, string name,
            string description) {
            LegacyPptShape shape = Assert.Single(shapes,
                item => item.Metadata.Name == name);
            Assert.Equal(description, shape.Metadata.Description);
        }
    }
}
