using System.Buffers.Binary;
using OfficeIMO.Drawing.Binary;
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
        public void NativeWriter_AuthorsAutoShapeStylesTransformsAndShadow() {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                PowerPointAutoShape rectangle = slide.AddShapePoints(
                    A.ShapeTypeValues.Rectangle, 24, 24, 120, 70);
                rectangle.Fill("C0504D").Stroke("4F81BD", 2.5D);
                rectangle.FillTransparency = 25;
                rectangle.OutlineDash = A.PresetLineDashValues.DashDot;
                rectangle.Rotation = 30D;
                rectangle.HorizontalFlip = true;
                rectangle.SetShadow("222222", blurPoints: 4D,
                    distancePoints: 3D, angleDegrees: 45D,
                    transparencyPercent: 35);

                PowerPointAutoShape chevron = slide.AddShapePoints(
                    A.ShapeTypeValues.Chevron, 180, 24, 100, 70);
                chevron.Fill("9BBB59").Stroke("8064A2", 1.25D);
                chevron.VerticalFlip = true;

                PowerPointAutoShape line = slide.AddShapePoints(
                    A.ShapeTypeValues.Line, 24, 130, 180, 40);
                line.Stroke("00B050", 3D);
                line.OutlineDash = A.PresetLineDashValues
                    .LargeDashDotDot;
                line.SetLineEnds(A.LineEndValues.Triangle,
                    A.LineEndValues.Diamond,
                    A.LineEndWidthValues.Large,
                    A.LineEndLengthValues.Small);
                line.HorizontalFlip = true;
                line.VerticalFlip = true;

                PowerPointTextBox text = slide.AddTextBoxPoints(
                    "Styled text", 24, 200, 160, 45);
                text.FillColor = "F2F2F2";
                text.OutlineColor = "7F7F7F";
                text.OutlineWidthPoints = 1D;
                text.Rotation = -15D;

                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);
            LegacyPptShape[] shapes = Assert.Single(binary.Slides).Shapes
                .OrderBy(shape => shape.Bounds.Top)
                .ThenBy(shape => shape.Bounds.Left)
                .ToArray();
            Assert.Equal(4, shapes.Length);
            LegacyPptShape rectangleShape = shapes.Single(shape =>
                shape.OfficeArtShapeType == 1);
            Assert.Equal(30D, rectangleShape.Transform.RotationDegrees);
            Assert.True(rectangleShape.Transform.FlipHorizontal);
            Assert.False(rectangleShape.Transform.FlipVertical);
            Assert.Equal("C0504D", rectangleShape.FillColor);
            Assert.Equal(0.75D, rectangleShape.Style.FillOpacity);
            Assert.Equal("4F81BD", rectangleShape.LineColor);
            Assert.Equal(31750, rectangleShape.Style.LineWidthEmus);
            Assert.Equal(8U, rectangleShape.Style.LineDashing);
            Assert.True(rectangleShape.Style.ShadowEnabled);
            Assert.Equal("222222", rectangleShape.ShadowColor);
            Assert.Equal(0.65D,
                rectangleShape.Style.ShadowOpacity!.Value, 4);
            Assert.Equal(26941,
                rectangleShape.Style.ShadowOffsetXEmus);
            Assert.Equal(26941,
                rectangleShape.Style.ShadowOffsetYEmus);
            Assert.Equal(50800,
                rectangleShape.Style.ShadowSoftnessEmus);

            LegacyPptShape chevronShape = shapes.Single(shape =>
                shape.OfficeArtShapeType == 55);
            Assert.True(chevronShape.Transform.FlipVertical);
            Assert.Equal("9BBB59", chevronShape.FillColor);
            LegacyPptShape lineShape = shapes.Single(shape =>
                shape.OfficeArtShapeType == 20);
            Assert.True(lineShape.Transform.FlipHorizontal);
            Assert.True(lineShape.Transform.FlipVertical);
            Assert.Equal(10U, lineShape.Style.LineDashing);
            Assert.Equal(1U, lineShape.Style.LineStartArrowhead);
            Assert.Equal(3U, lineShape.Style.LineEndArrowhead);
            Assert.Equal(2U, lineShape.Style.LineStartArrowWidth);
            Assert.Equal(0U, lineShape.Style.LineStartArrowLength);
            LegacyPptShape textShape = shapes.Single(shape =>
                shape.OfficeArtShapeType == 202);
            Assert.Equal(345D, textShape.Transform.RotationDegrees);
            Assert.Equal("Styled text", textShape.Text);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation
                .Load(input);
            Assert.Equal(30D, projected.Slides[0].Shapes
                .Single(shape => shape.FillColor == "C0504D").Rotation);
            Assert.Contains(projected.Slides[0].Shapes
                .OfType<PowerPointAutoShape>(), shape =>
                    shape.ShapeType == A.ShapeTypeValues.Chevron);
            Assert.Empty(projected.ValidateDocument());
            Assert.Equal(bytes,
                projected.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void NativeWriter_AuthorsUnattachedConnectorsWithLineFormatting() {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                PowerPointConnectionShape connector = slide
                    .AddConnectionShape(A.ShapeTypeValues.BentConnector3,
                        PowerPointUnits.FromPoints(24),
                        PowerPointUnits.FromPoints(32),
                        PowerPointUnits.FromPoints(180),
                        PowerPointUnits.FromPoints(90));
                connector.Stroke("4472C4", 2.25D);
                connector.OutlineDash = A.PresetLineDashValues.DashDot;
                connector.SetLineEnds(A.LineEndValues.Oval,
                    A.LineEndValues.Triangle,
                    A.LineEndWidthValues.Large,
                    A.LineEndLengthValues.Small);
                connector.Rotation = 15D;
                connector.HorizontalFlip = true;

                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptShape connectorShape = Assert.Single(Assert.Single(
                LegacyPptPresentation.Load(bytes).Slides).Shapes);
            Assert.Equal(LegacyPptShapeKind.Connector,
                connectorShape.Kind);
            Assert.Equal(34, connectorShape.OfficeArtShapeType);
            Assert.Equal("4472C4", connectorShape.LineColor);
            Assert.Equal(28575, connectorShape.Style.LineWidthEmus);
            Assert.Equal(8U, connectorShape.Style.LineDashing);
            Assert.Equal(4U, connectorShape.Style.LineStartArrowhead);
            Assert.Equal(1U, connectorShape.Style.LineEndArrowhead);
            Assert.Equal(2U, connectorShape.Style.LineStartArrowWidth);
            Assert.Equal(0U, connectorShape.Style.LineStartArrowLength);
            Assert.Equal(15D,
                connectorShape.Transform.RotationDegrees);
            Assert.True(connectorShape.Transform.FlipHorizontal);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation
                .Load(input);
            PowerPointConnectionShape projectedConnector = Assert.Single(
                projected.Slides[0].Shapes
                    .OfType<PowerPointConnectionShape>());
            Assert.Equal(A.ShapeTypeValues.BentConnector3,
                projectedConnector.ShapeType);
            Assert.Empty(projected.ValidateDocument());
            Assert.Equal(bytes,
                projected.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void NativeWriter_BlocksFreshConnectorAttachmentsUntilSolverRulesCanBeWritten() {
            using PowerPointPresentation source = PowerPointPresentation
                .Create();
            PowerPointSlide slide = source.AddSlide(
                P.SlideLayoutValues.Blank);
            PowerPointConnectionShape connector = slide
                .AddConnectionShape(A.ShapeTypeValues.StraightConnector1,
                    100, 200, 300, 400);
            P.ConnectionShape element = Assert.IsType<P.ConnectionShape>(
                connector.Element);
            element.NonVisualConnectionShapeProperties!
                .NonVisualConnectorShapeDrawingProperties!
                .Append(new A.StartConnection { Id = 2U, Index = 0U });

            LegacyPptWritePreflightReport preflight = source
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-SHAPE");
        }

        [Fact]
        public void NativeWriter_AuthorsExactRoundRectangleAndDonutAdjustments() {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                PowerPointAutoShape round = slide.AddShapePoints(
                    A.ShapeTypeValues.RoundRectangle, 20, 20, 120, 70);
                SetAdjustment(round, "val 30000");
                PowerPointAutoShape donut = slide.AddShapePoints(
                    A.ShapeTypeValues.Donut, 170, 20, 90, 90);
                SetAdjustment(donut, "val 40000");

                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptShape[] shapes = Assert.Single(
                LegacyPptPresentation.Load(bytes).Slides).Shapes.ToArray();
            Assert.Equal(6480, shapes.Single(shape =>
                shape.OfficeArtShapeType == 2).Geometry.AdjustmentValues[0]);
            Assert.Equal(8640, shapes.Single(shape =>
                shape.OfficeArtShapeType == 23).Geometry.AdjustmentValues[0]);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation
                .Load(input);
            Assert.Contains(projected.Slides[0].Shapes
                .OfType<PowerPointAutoShape>(), shape =>
                    GetAdjustment(shape) == "val 30000");
            Assert.Contains(projected.Slides[0].Shapes
                .OfType<PowerPointAutoShape>(), shape =>
                    GetAdjustment(shape) == "val 40000");
            Assert.Equal(bytes,
                projected.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void NativeWriter_BlocksInexactOfficeArtAdjustmentConversion() {
            using PowerPointPresentation source = PowerPointPresentation
                .Create();
            PowerPointAutoShape round = source.AddSlide(
                    P.SlideLayoutValues.Blank)
                .AddShapePoints(A.ShapeTypeValues.RoundRectangle,
                    20, 20, 120, 70);
            SetAdjustment(round, "val 33334");

            LegacyPptWritePreflightReport preflight = source
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-SHAPE");
        }

        [Fact]
        public void NativeWriter_AuthorsNestedGroupsAndChildAnchors() {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                PowerPointAutoShape rectangle = slide.AddShapePoints(
                    A.ShapeTypeValues.Rectangle, 30, 40, 80, 50);
                rectangle.Fill("4472C4");
                PowerPointAutoShape ellipse = slide.AddShapePoints(
                    A.ShapeTypeValues.Ellipse, 125, 50, 65, 45);
                ellipse.Fill("ED7D31");
                PowerPointGroupShape inner = slide.GroupShapes(
                    new PowerPointShape[] { rectangle, ellipse },
                    "Inner group");
                PowerPointTextBox label = slide.AddTextBoxPoints(
                    "Grouped label", 55, 115, 120, 35);
                PowerPointGroupShape outer = slide.GroupShapes(
                    new PowerPointShape[] { inner, label },
                    "Outer group");
                outer.Rotation = 12D;
                outer.HorizontalFlip = true;

                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptSlide binarySlide = Assert.Single(
                LegacyPptPresentation.Load(bytes).Slides);
            LegacyPptShape outerGroup = Assert.Single(binarySlide.Shapes);
            Assert.Equal(LegacyPptShapeKind.Group, outerGroup.Kind);
            Assert.Equal(12D, outerGroup.Transform.RotationDegrees);
            Assert.True(outerGroup.Transform.FlipHorizontal);
            Assert.NotNull(outerGroup.GroupCoordinateBounds);
            Assert.Equal(2, outerGroup.Children.Count);
            LegacyPptShape innerGroup = Assert.Single(outerGroup.Children,
                child => child.Kind == LegacyPptShapeKind.Group);
            Assert.Equal(2, innerGroup.Children.Count);
            Assert.All(innerGroup.Children, child =>
                Assert.True(child.Bounds.Width > 0));
            Assert.Contains(outerGroup.Children,
                child => child.Text == "Grouped label");

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation
                .Load(input);
            PowerPointGroupShape projectedOuter = Assert.Single(
                projected.Slides[0].Shapes
                    .OfType<PowerPointGroupShape>());
            IReadOnlyList<PowerPointShape> projectedChildren = projected
                .Slides[0].GetGroupChildren(projectedOuter);
            Assert.Contains(projectedChildren,
                child => child is PowerPointGroupShape);
            Assert.Contains(projectedChildren,
                child => child is PowerPointTextBox text
                    && text.Text == "Grouped label");
            Assert.Empty(projected.ValidateDocument());
            Assert.Equal(bytes,
                projected.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void NativeWriter_BlocksShapeEffectWithoutOfficeArtEquivalent() {
            using PowerPointPresentation source = PowerPointPresentation
                .Create();
            PowerPointAutoShape shape = source.AddSlide(
                    P.SlideLayoutValues.Blank)
                .AddShapePoints(A.ShapeTypeValues.Rectangle,
                    20, 20, 100, 60);
            shape.SetGlow("4472C4", radiusPoints: 5D);

            LegacyPptWritePreflightReport preflight = source
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-SHAPE-STYLE");
        }

        [Fact]
        public void ImportedShapeStyleAndTransformEdit_UsesIncrementalFoptRewrite() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                PowerPointAutoShape shape = source.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddShapePoints(A.ShapeTypeValues.Rectangle,
                        20, 20, 120, 70);
                shape.Fill("4472C4").Stroke("203864", 2D);
                shape.SetShadow("222222", blurPoints: 3D,
                    distancePoints: 2D, angleDegrees: 90D,
                    transparencyPercent: 40);
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);
            using var input = new MemoryStream(sourceBytes, writable: false);
            using PowerPointPresentation imported = PowerPointPresentation
                .Load(input);
            PowerPointAutoShape edited = Assert.Single(imported.Slides[0]
                .Shapes.OfType<PowerPointAutoShape>());
            edited.FillColor = "ED7D31";
            edited.FillTransparency = 10;
            edited.OutlineColor = "A5A5A5";
            edited.OutlineWidthPoints = 3D;
            edited.OutlineDash = A.PresetLineDashValues.LargeDash;
            edited.Rotation = 42D;
            edited.HorizontalFlip = true;
            edited.VerticalFlip = true;
            edited.ClearShadow();

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                imported.ToBytes(PowerPointFileFormat.Ppt));

            LegacyPptShape savedShape = Assert.Single(Assert.Single(saved.Slides)
                .Shapes);
            Assert.Equal("ED7D31", savedShape.FillColor);
            Assert.Equal(0.9D,
                savedShape.Style.FillOpacity!.Value, 4);
            Assert.Equal("A5A5A5", savedShape.LineColor);
            Assert.Equal(38100, savedShape.Style.LineWidthEmus);
            Assert.Equal(7U, savedShape.Style.LineDashing);
            Assert.Equal(42D, savedShape.Transform.RotationDegrees);
            Assert.True(savedShape.Transform.FlipHorizontal);
            Assert.True(savedShape.Transform.FlipVertical);
            Assert.Null(savedShape.Style.ShadowEnabled);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void NativeWriter_AuthorsAndProjectsHiddenShapesAndGroups() {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                PowerPointAutoShape hiddenChild = slide.AddShapePoints(
                    A.ShapeTypeValues.Rectangle, 20, 20, 90, 50);
                hiddenChild.Name = "Hidden child";
                hiddenChild.Hidden = true;
                PowerPointAutoShape visibleChild = slide.AddShapePoints(
                    A.ShapeTypeValues.Ellipse, 130, 20, 70, 50);
                visibleChild.Name = "Visible child";
                PowerPointGroupShape group = slide.GroupShapes(
                    new PowerPointShape[] { hiddenChild, visibleChild },
                    "Hidden group");
                group.Hidden = true;

                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptShape binaryGroup = Assert.Single(Assert.Single(
                LegacyPptPresentation.Load(bytes).Slides).Shapes);
            Assert.True(binaryGroup.Style.Hidden);
            Assert.True(Assert.Single(binaryGroup.Children,
                child => child.Metadata.Name == "Hidden child").Style.Hidden);
            Assert.False(Assert.Single(binaryGroup.Children,
                child => child.Metadata.Name == "Visible child").Style.Hidden
                ?? false);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation
                .Load(input);
            PowerPointGroupShape projectedGroup = Assert.Single(projected
                .Slides[0].Shapes.OfType<PowerPointGroupShape>());
            Assert.True(projectedGroup.Hidden);
            Assert.True(Assert.Single(projected.Slides[0]
                .GetGroupChildren(projectedGroup), child =>
                    child.Name == "Hidden child").Hidden);
            Assert.False(Assert.Single(projected.Slides[0]
                .GetGroupChildren(projectedGroup), child =>
                    child.Name == "Visible child").Hidden);
            Assert.Empty(projected.ValidateDocument());
            Assert.Equal(bytes,
                projected.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ImportedShapeVisibilityEdit_UsesIncrementalFoptRewrite() {
            byte[] sourceBytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                PowerPointTextBox visible = slide.AddTextBoxPoints(
                    "Visible", 20, 20, 100, 35);
                visible.Name = "Initially visible";
                PowerPointTextBox hidden = slide.AddTextBoxPoints(
                    "Hidden", 20, 70, 100, 35);
                hidden.Name = "Initially hidden";
                hidden.Hidden = true;
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);
            byte[] editedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation
                       .Load(input)) {
                PowerPointTextBox visible = Assert.Single(imported.Slides[0]
                    .TextBoxes, shape => shape.Name == "Initially visible");
                PowerPointTextBox hidden = Assert.Single(imported.Slides[0]
                    .TextBoxes, shape => shape.Name == "Initially hidden");
                visible.Hidden = true;
                hidden.Hidden = false;

                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                editedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                editedBytes);
            Assert.True(Assert.Single(saved.Slides[0].Shapes,
                shape => shape.Metadata.Name == "Initially visible")
                .Style.Hidden);
            Assert.False(Assert.Single(saved.Slides[0].Shapes,
                shape => shape.Metadata.Name == "Initially hidden")
                .Style.Hidden ?? false);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            using var editedInput = new MemoryStream(editedBytes,
                writable: false);
            using PowerPointPresentation reopened = PowerPointPresentation
                .Load(editedInput);
            Assert.True(Assert.Single(reopened.Slides[0].TextBoxes,
                shape => shape.Name == "Initially visible").Hidden);
            Assert.False(Assert.Single(reopened.Slides[0].TextBoxes,
                shape => shape.Name == "Initially hidden").Hidden);
            Assert.Equal(editedBytes,
                reopened.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ImportedShapeVisibilityEdit_RewritesAuthoritativeTertiaryFopt() {
            byte[] image = OfficePngWriter.Encode(new OfficeRasterImage(
                4, 3, OfficeColor.FromRgb(37, 99, 235)));
            byte[] generatedBytes;
            using (PowerPointPresentation source = PowerPointPresentation
                       .Create()) {
                PowerPointSlide slide = source.AddSlide(
                    P.SlideLayoutValues.Blank);
                using var stream = new MemoryStream(image, writable: false);
                PowerPointPicture picture = slide.AddPicture(stream,
                    ImagePartType.Png, 158750, 317500, 635000, 476250);
                picture.Name = "Tertiary visibility picture";
                picture.Hidden = true;
                picture.RecolorColor = OfficeColor.Orange;
                generatedBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation generated = LegacyPptPresentation.Load(
                generatedBytes);
            LegacyPptSlide generatedSlide = Assert.Single(generated.Slides);
            LegacyPptShape generatedPicture = Assert.Single(
                generatedSlide.Shapes);
            Assert.True(generatedPicture.Style.Hidden);
            LegacyPptRecord shapeContainer = LegacyPptRecordReader.ReadSingle(
                generated.Package.DocumentStream,
                checked((int)generatedPicture.RecordOffset),
                new LegacyPptImportOptions());
            LegacyPptRecord tertiaryFopt = Assert.Single(
                shapeContainer.Children,
                child => child.Type == 0xF122);
            Assert.Equal(1, tertiaryFopt.Instance);

            byte[] documentStream = (byte[])generated.Package.DocumentStream
                .Clone();
            int propertyOffset = tertiaryFopt.PayloadOffset;
            BinaryPrimitives.WriteUInt16LittleEndian(
                documentStream.AsSpan(propertyOffset, 2), 0x03BF);
            BinaryPrimitives.WriteUInt32LittleEndian(
                documentStream.AsSpan(propertyOffset + 2, 4), 1U << 14);
            byte[] sourceBytes = generated.Package.RewriteCompoundStreams(
                new Dictionary<string, byte[]> {
                    ["PowerPoint Document"] = documentStream
                });

            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourceBytes);
            Assert.False(Assert.Single(Assert.Single(original.Slides).Shapes)
                .Style.Hidden ?? true);
            byte[] editedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported = PowerPointPresentation
                       .Load(input)) {
                PowerPointPicture picture = Assert.Single(imported.Slides[0]
                    .Pictures);
                Assert.False(picture.Hidden);
                picture.Hidden = true;

                LegacyPptWritePreflightReport preflight = imported
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                editedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                editedBytes);
            Assert.True(Assert.Single(Assert.Single(saved.Slides).Shapes)
                .Style.Hidden);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            using var editedInput = new MemoryStream(editedBytes,
                writable: false);
            using PowerPointPresentation reopened = PowerPointPresentation
                .Load(editedInput);
            Assert.True(Assert.Single(reopened.Slides[0].Pictures).Hidden);
            Assert.Equal(editedBytes,
                reopened.ToBytes(PowerPointFileFormat.Ppt));
        }

        private static void SetAdjustment(PowerPointAutoShape shape,
            string formula) {
            P.Shape element = Assert.IsType<P.Shape>(shape.Element);
            A.PresetGeometry geometry = element.ShapeProperties!
                .GetFirstChild<A.PresetGeometry>()!;
            A.AdjustValueList values = geometry.AdjustValueList ??=
                new A.AdjustValueList();
            values.RemoveAllChildren<A.ShapeGuide>();
            values.Append(new A.ShapeGuide {
                Name = "adj",
                Formula = formula
            });
        }

        private static string? GetAdjustment(PowerPointAutoShape shape) =>
            Assert.IsType<P.Shape>(shape.Element).ShapeProperties?
                .GetFirstChild<A.PresetGeometry>()?.AdjustValueList?
                .GetFirstChild<A.ShapeGuide>()?.Formula?.Value;
    }
}
