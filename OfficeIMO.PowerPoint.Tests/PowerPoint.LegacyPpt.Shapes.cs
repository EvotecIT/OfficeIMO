using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.Drawing.Binary;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        private static string ShapeFixturePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "ShapePowerPoint.ppt");

        private static string TransformFixturePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "TransformPowerPoint.ppt");

        private static string ConnectedFixturePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "ConnectedPowerPoint.ppt");

        private static string AdjustedShapesFixturePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "AdjustedShapesPowerPoint.ppt");

        private static string ShadowFixturePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "ShadowPowerPoint.ppt");

        public static IEnumerable<object[]> RepresentativeOfficeArtPresets => new[] {
            new object[] { (ushort)4, A.ShapeTypeValues.Diamond },
            new object[] { (ushort)9, A.ShapeTypeValues.Hexagon },
            new object[] { (ushort)12, A.ShapeTypeValues.Star5 },
            new object[] { (ushort)55, A.ShapeTypeValues.Chevron },
            new object[] { (ushort)110, A.ShapeTypeValues.FlowChartDecision },
            new object[] { (ushort)32, A.ShapeTypeValues.StraightConnector1 },
            new object[] { (ushort)34, A.ShapeTypeValues.BentConnector3 },
            new object[] { (ushort)38, A.ShapeTypeValues.CurvedConnector3 }
        };

        [Theory]
        [MemberData(nameof(RepresentativeOfficeArtPresets))]
        public void OfficeArtGeometryMapper_MapsRepresentativePresetFamilies(ushort officeArtType,
            A.ShapeTypeValues expected) {
            Assert.True(LegacyPptShapeGeometryMapper.TryGetPreset(officeArtType,
                out A.ShapeTypeValues actual));
            Assert.Equal(expected, actual);
            Assert.True(LegacyPptShapeGeometryMapper.TryGetShapeType(expected,
                out ushort roundTripType));
            Assert.Equal(officeArtType, roundTripType);
        }

        [Theory]
        [InlineData(24)]
        [InlineData(136)]
        [InlineData(201)]
        [InlineData(202)]
        public void OfficeArtGeometryMapper_DoesNotClaimUnsupportedGeometry(ushort officeArtType) {
            Assert.False(LegacyPptShapeGeometryMapper.TryGetPreset(officeArtType, out _));
        }

        [Theory]
        [InlineData(14)]
        [InlineData(17)]
        [InlineData(18)]
        [InlineData(100)]
        [InlineData(178)]
        [InlineData(181)]
        public void OfficeArtGeometryMapper_IdentifiesDocumentedApproximations(ushort officeArtType) {
            Assert.True(LegacyPptShapeGeometryMapper.IsApproximation(officeArtType));
        }

        [Fact]
        public void NeutralReader_DecodesPresetShapesConnectorsAndGroupHierarchy() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(ShapeFixturePath);

            LegacyPptSlide slide = Assert.Single(legacy.Slides);
            Assert.Equal(9, slide.Shapes.Count);
            Assert.DoesNotContain(slide.Shapes, shape => shape.Kind == LegacyPptShapeKind.Unsupported);
            Assert.All(new ushort[] { 4, 9, 12, 55, 110 }, officeArtType =>
                Assert.Contains(slide.Shapes, shape => shape.OfficeArtShapeType == officeArtType));
            Assert.Equal(3, slide.Shapes.Count(shape => shape.Kind == LegacyPptShapeKind.Connector));

            LegacyPptShape group = Assert.Single(slide.Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Group);
            Assert.Equal(new LegacyPptBounds(576, 2880, 2245, 575), group.Bounds);
            Assert.Equal(new LegacyPptBounds(576, 2880, 2245, 575), group.GroupCoordinateBounds);
            Assert.Equal(2, group.Children.Count);
            Assert.Contains(group.Children, child => child.OfficeArtShapeType == 2
                && child.Bounds.Equals(new LegacyPptBounds(576, 2880, 1151, 575)));
            Assert.Contains(group.Children, child => child.OfficeArtShapeType == 3
                && child.Bounds.Equals(new LegacyPptBounds(1958, 2880, 863, 575)));
            Assert.DoesNotContain(legacy.Diagnostics, diagnostic =>
                diagnostic.Code == "PPT-GROUP-UNSUPPORTED"
                || diagnostic.Code == "PPT-GROUP-TRUNCATED");
        }

        [Fact]
        public void NormalLoad_ProjectsPresetShapesConnectorsAndGroupsIntoValidPptxModel() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(ShapeFixturePath);

            PowerPointSlide slide = Assert.Single(presentation.Slides);
            Assert.Equal(9, slide.Shapes.Count);
            PowerPointTextBox diamond = Assert.Single(slide.TextBoxes,
                shape => shape.Text == "Diamond");
            Assert.Equal(A.ShapeTypeValues.Diamond, diamond.ShapeType);
            Assert.Contains(slide.Shapes.OfType<PowerPointAutoShape>(),
                shape => shape.ShapeType == A.ShapeTypeValues.Hexagon);

            PowerPointConnectionShape[] connectors = slide.Shapes
                .OfType<PowerPointConnectionShape>()
                .ToArray();
            Assert.Equal(3, connectors.Length);
            Assert.Contains(connectors, shape => shape.ShapeType == A.ShapeTypeValues.StraightConnector1);
            Assert.Contains(connectors, shape => shape.ShapeType == A.ShapeTypeValues.BentConnector3);
            Assert.Contains(connectors, shape => shape.ShapeType == A.ShapeTypeValues.CurvedConnector3);

            PowerPointGroupShape group = Assert.Single(slide.Shapes.OfType<PowerPointGroupShape>());
            IReadOnlyList<PowerPointShape> children = slide.GetGroupChildren(group);
            Assert.Equal(2, children.Count);
            Assert.Contains(children.OfType<PowerPointTextBox>(), child =>
                child.Text == "Grouped" && child.ShapeType == A.ShapeTypeValues.RoundRectangle);
            Assert.Contains(children.OfType<PowerPointAutoShape>(), child =>
                child.ShapeType == A.ShapeTypeValues.Ellipse);

            for (int index = 0; index < slide.Shapes.Count; index++) {
                Assert.Equal(index, slide.Shapes[index].DrawingOrder);
            }
            Assert.Empty(presentation.ValidateDocument());

            using MemoryStream pptx = presentation.ToStream();
            using PowerPointPresentation reopened = PowerPointPresentation.Load(pptx);
            Assert.Equal(9, reopened.Slides[0].Shapes.Count);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void NormalLoad_ProjectsScaledShapeGradientsWithResolvedEndpointColors() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(ShapeFixturePath);
            LegacyPptShape binaryDiamond = legacy.Slides[0].Shapes.Single(shape =>
                shape.Text == "Diamond");
            Assert.Equal(7U, binaryDiamond.Style.FillType);
            Assert.Equal(-180D, binaryDiamond.Style.FillAngleDegrees);
            Assert.Equal("3E7FCC", binaryDiamond.FillColor);
            Assert.Equal("A4C1FF", binaryDiamond.FillBackColor);

            using PowerPointPresentation presentation = PowerPointPresentation.Load(ShapeFixturePath);
            PowerPointTextBox diamond = presentation.Slides[0].TextBoxes.Single(shape =>
                shape.Text == "Diamond");
            P.Shape element = Assert.IsType<P.Shape>(diamond.Element);
            A.GradientFill gradient = Assert.IsType<A.GradientFill>(element.ShapeProperties!
                .GetFirstChild<A.GradientFill>());
            A.GradientStop[] stops = gradient.GetFirstChild<A.GradientStopList>()!
                .Elements<A.GradientStop>()
                .ToArray();
            Assert.Collection(stops,
                stop => Assert.Equal("3E7FCC", stop.RgbColorModelHex!.Val!.Value),
                stop => Assert.Equal("A4C1FF", stop.RgbColorModelHex!.Val!.Value));
            A.LinearGradientFill linear = Assert.IsType<A.LinearGradientFill>(
                gradient.GetFirstChild<A.LinearGradientFill>());
            Assert.Equal(270 * 60000, linear.Angle!.Value);
            Assert.True(linear.Scaled!.Value);
            Assert.False(gradient.RotateWithShape!.Value);
        }

        [Fact]
        public void ShapeGradientProjection_UsesCompleteCustomStopsWithoutEndpointProperties() {
            OfficeArtShapeStyle style = OfficeArtShapeStyle.Decode(new[] {
                new OfficeArtProperty(0, 0x0180, 4U)
            });
            var source = new LegacyPptShape(LegacyPptShapeKind.Rectangle, 1, 1, 0,
                new LegacyPptBounds(0, 0, 100, 100), string.Empty,
                placeholder: null, style, fillColor: null, lineColor: null,
                fillGradientStops: new[] {
                    new LegacyPptGradientStop("112233", 0D),
                    new LegacyPptGradientStop("445566", 1D)
                });
            var properties = new P.ShapeProperties(
                new A.Transform2D(),
                new A.PresetGeometry(new A.AdjustValueList()) {
                    Preset = A.ShapeTypeValues.Rectangle
                });

            PowerPointPresentation.ApplyLegacyShapeStyle(properties, source);

            A.GradientStop[] stops = properties.GetFirstChild<A.GradientFill>()!
                .GetFirstChild<A.GradientStopList>()!
                .Elements<A.GradientStop>()
                .ToArray();
            Assert.Collection(stops,
                stop => Assert.Equal("112233", stop.RgbColorModelHex!.Val!.Value),
                stop => Assert.Equal("445566", stop.RgbColorModelHex!.Val!.Value));
        }

        [Fact]
        public void ShapeGradientProjection_UsesOfficeArtWhiteDefaultsForOmittedEndpoints() {
            OfficeArtShapeStyle style = OfficeArtShapeStyle.Decode(new[] {
                new OfficeArtProperty(0, 0x0180, 4U)
            });
            var source = new LegacyPptShape(LegacyPptShapeKind.Rectangle, 1, 1, 0,
                new LegacyPptBounds(0, 0, 100, 100), string.Empty,
                placeholder: null, style, fillColor: null, lineColor: null);
            var properties = new P.ShapeProperties(
                new A.Transform2D(),
                new A.PresetGeometry(new A.AdjustValueList()) {
                    Preset = A.ShapeTypeValues.Rectangle
                });

            PowerPointPresentation.ApplyLegacyShapeStyle(properties, source);

            A.GradientStop[] stops = properties.GetFirstChild<A.GradientFill>()!
                .GetFirstChild<A.GradientStopList>()!
                .Elements<A.GradientStop>()
                .ToArray();
            Assert.Collection(stops,
                stop => Assert.Equal("FFFFFF", stop.RgbColorModelHex!.Val!.Value),
                stop => Assert.Equal("FFFFFF", stop.RgbColorModelHex!.Val!.Value));
        }

        [Fact]
        public void UnmodifiedShapeBinarySave_PreservesOriginalPackageExactly() {
            byte[] source = File.ReadAllBytes(ShapeFixturePath);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(ShapeFixturePath);

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            Assert.Equal(source, presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ImportedConnectorAndGroupGeometryEdits_UseIncrementalAnchorRewrite() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(ShapeFixturePath);
            LegacyPptShape originalConnector = original.Slides[0].Shapes.Single(shape =>
                shape.Kind == LegacyPptShapeKind.Connector && shape.OfficeArtShapeType == 32);
            LegacyPptShape originalGroup = original.Slides[0].Shapes.Single(shape =>
                shape.Kind == LegacyPptShapeKind.Group);

            using PowerPointPresentation presentation = PowerPointPresentation.Load(ShapeFixturePath);
            PowerPointSlide slide = presentation.Slides[0];
            PowerPointConnectionShape connector = slide.Shapes.OfType<PowerPointConnectionShape>()
                .Single(shape => shape.ShapeType == A.ShapeTypeValues.StraightConnector1);
            PowerPointGroupShape group = Assert.Single(slide.Shapes.OfType<PowerPointGroupShape>());
            connector.Left += 15875;
            group.Top += 15875;
            P.GroupShape groupElement = Assert.IsType<P.GroupShape>(
                group.Element);
            A.TransformGroup groupTransform = groupElement
                .GroupShapeProperties!.TransformGroup!;
            groupTransform.ChildOffset!.X = checked(
                groupTransform.ChildOffset.X!.Value + 15875L);

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptShape savedConnector = saved.Slides[0].Shapes.Single(shape =>
                shape.Kind == LegacyPptShapeKind.Connector && shape.OfficeArtShapeType == 32);
            LegacyPptShape savedGroup = saved.Slides[0].Shapes.Single(shape =>
                shape.Kind == LegacyPptShapeKind.Group);
            Assert.Equal(originalConnector.Bounds.Left + 10, savedConnector.Bounds.Left);
            Assert.Equal(originalGroup.Bounds.Top + 10, savedGroup.Bounds.Top);
            Assert.Equal(originalGroup.GroupCoordinateBounds!.Value.Left + 10,
                savedGroup.GroupCoordinateBounds!.Value.Left);
            Assert.Equal(originalGroup.Children.Select(child => child.Bounds),
                savedGroup.Children.Select(child => child.Bounds));
            Assert.Equal(original.Package.UserEdits.Count + 1, saved.Package.UserEdits.Count);
        }

        [Fact]
        public void ImportedGroupChildEdit_UsesIncrementalChildAnchorRewrite() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                ShapeFixturePath);
            LegacyPptShape originalGroup = Assert.Single(original.Slides)
                .Shapes.Single(shape => shape.Kind
                    == LegacyPptShapeKind.Group);
            LegacyPptShape originalChild = originalGroup.Children[0];
            using PowerPointPresentation presentation = PowerPointPresentation.Load(ShapeFixturePath);
            PowerPointSlide slide = presentation.Slides[0];
            PowerPointGroupShape group = Assert.Single(slide.Shapes.OfType<PowerPointGroupShape>());
            PowerPointShape child = slide.GetGroupChildren(group)[0];

            child.Left += 15875;
            child.Rotation = 18D;
            child.VerticalFlip = true;

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));

            LegacyPptShape savedGroup = Assert.Single(saved.Slides).Shapes
                .Single(shape => shape.Kind == LegacyPptShapeKind.Group);
            Assert.Equal(originalChild.Bounds.Left + 10,
                savedGroup.Children[0].Bounds.Left);
            Assert.Equal(18D,
                savedGroup.Children[0].Transform.RotationDegrees);
            Assert.True(savedGroup.Children[0].Transform.FlipVertical);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void NeutralReader_DecodesRotationAndFlips() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(TransformFixturePath);

            LegacyPptSlide slide = Assert.Single(legacy.Slides);
            Assert.Equal(5, slide.Shapes.Count);
            Assert.Equal(30D, slide.Shapes.Single(shape => shape.Text == "Rotate 30")
                .Transform.RotationDegrees);
            Assert.Equal(315D, slide.Shapes.Single(shape => shape.Text == "Rotate -45")
                .Transform.RotationDegrees);
            Assert.True(slide.Shapes.Single(shape => shape.Text == "Flip H")
                .Transform.FlipHorizontal);
            Assert.True(slide.Shapes.Single(shape => shape.Text == "Flip V")
                .Transform.FlipVertical);
            Assert.Contains(slide.Shapes, shape => shape.Kind == LegacyPptShapeKind.Group);
        }

        [Fact]
        public void NormalLoad_ProjectsRotationAndFlipsIntoValidPptxModel() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(TransformFixturePath);

            PowerPointSlide slide = Assert.Single(presentation.Slides);
            PowerPointTextBox rotated = slide.TextBoxes.Single(shape => shape.Text == "Rotate 30");
            Assert.Equal(30D, rotated.Rotation);
            Assert.Equal(315D, slide.TextBoxes.Single(shape => shape.Text == "Rotate -45").Rotation);
            Assert.True(slide.TextBoxes.Single(shape => shape.Text == "Flip H").HorizontalFlip);
            Assert.True(slide.TextBoxes.Single(shape => shape.Text == "Flip V").VerticalFlip);
            Assert.Single(slide.Shapes.OfType<PowerPointGroupShape>());
            A.GradientFill gradient = Assert.IsType<A.GradientFill>(
                Assert.IsType<P.Shape>(rotated.Element).ShapeProperties!
                    .GetFirstChild<A.GradientFill>());
            Assert.False(gradient.RotateWithShape!.Value);
            Assert.Equal(300 * 60000, gradient.GetFirstChild<A.LinearGradientFill>()!
                .Angle!.Value);
            Assert.Empty(presentation.ValidateDocument());

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            Assert.Equal(File.ReadAllBytes(TransformFixturePath),
                presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void LegacyGroupProjection_UsesNativeRotationAndFlipAttributes() {
            OfficeArtShapeStyle style = OfficeArtShapeStyle.Decode(Array.Empty<OfficeArtProperty>());
            var child = new LegacyPptShape(LegacyPptShapeKind.Rectangle, 1, 20, 0,
                new LegacyPptBounds(0, 0, 100, 100), string.Empty,
                placeholder: null, style, null, null);
            OfficeArtShapeTransform transform = OfficeArtShapeTransform.Decode(1U << 6,
                new[] { new OfficeArtProperty(0, 0x0004, 15U * 65536U) });
            var source = new LegacyPptShape(LegacyPptShapeKind.Group, 0, 10, 0,
                new LegacyPptBounds(100, 200, 400, 300), string.Empty,
                placeholder: null, style, null, null, transform: transform,
                groupCoordinateBounds: new LegacyPptBounds(0, 0, 400, 300),
                children: new[] { child });
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            uint nextShapeId = 2;

            P.GroupShape group = Assert.IsType<P.GroupShape>(
                PowerPointPresentation.CreateLegacyOpenXmlShape(slide.SlidePart, source,
                    ref nextShapeId));
            A.TransformGroup projected = group.GroupShapeProperties!.TransformGroup!;

            Assert.Equal(15 * 60000, projected.Rotation!.Value);
            Assert.True(projected.HorizontalFlip!.Value);
            Assert.Null(projected.VerticalFlip);
            var wrapper = new PowerPointGroupShape(group, slide.SlidePart);
            Assert.Equal(15D, wrapper.Rotation);
            wrapper.Rotation = 20D;
            wrapper.VerticalFlip = true;
            Assert.Equal(20D, wrapper.Rotation);
            Assert.True(wrapper.VerticalFlip);
        }

        [Fact]
        public void ImportedTransformEdit_PreservesUnrelatedRichStyle() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                TransformFixturePath);
            LegacyPptShape originalShape = Assert.Single(original.Slides)
                .Shapes.Single(item => item.Text == "Rotate 30");
            using PowerPointPresentation presentation = PowerPointPresentation.Load(TransformFixturePath);
            PowerPointTextBox shape = presentation.Slides[0].TextBoxes.Single(item =>
                item.Text == "Rotate 30");

            shape.Rotation = 42D;

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptShape savedShape = Assert.Single(saved.Slides).Shapes
                .Single(item => item.Text == "Rotate 30");
            Assert.Equal(42D, savedShape.Transform.RotationDegrees);
            Assert.Equal(originalShape.Style.FillType,
                savedShape.Style.FillType);
            Assert.Equal(originalShape.Style.FillGradientStops,
                savedShape.Style.FillGradientStops);
            Assert.Equal(originalShape.FillColor, savedShape.FillColor);
            Assert.Equal(originalShape.LineColor, savedShape.LineColor);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        [Fact]
        public void NeutralReader_DecodesConnectorSolverRuleAndConnectionSites() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(ConnectedFixturePath);

            LegacyPptSlide slide = Assert.Single(legacy.Slides);
            LegacyPptShape start = slide.Shapes.Single(shape => shape.Text == "Start");
            LegacyPptShape end = slide.Shapes.Single(shape => shape.Text == "End");
            LegacyPptShape connector = Assert.Single(slide.Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Connector);
            LegacyPptConnectorRule rule = Assert.Single(slide.ConnectorRules);
            Assert.Equal(start.ShapeId, rule.StartShapeId);
            Assert.Equal(end.ShapeId, rule.EndShapeId);
            Assert.Equal(connector.ShapeId, rule.ConnectorShapeId);
            Assert.Equal(3U, rule.StartConnectionSiteIndex);
            Assert.Equal(1U, rule.EndConnectionSiteIndex);
        }

        [Fact]
        public void NormalLoad_ProjectsNativeConnectorAttachmentsAndPreservesBinaryExactly() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(ConnectedFixturePath);

            PowerPointSlide slide = Assert.Single(presentation.Slides);
            PowerPointTextBox start = slide.TextBoxes.Single(shape => shape.Text == "Start");
            PowerPointTextBox end = slide.TextBoxes.Single(shape => shape.Text == "End");
            PowerPointConnectionShape connector = Assert.Single(
                slide.Shapes.OfType<PowerPointConnectionShape>());
            P.ConnectionShape element = Assert.IsType<P.ConnectionShape>(connector.Element);
            A.StartConnection startConnection = Assert.IsType<A.StartConnection>(element
                .NonVisualConnectionShapeProperties!.NonVisualConnectorShapeDrawingProperties!
                .GetFirstChild<A.StartConnection>());
            A.EndConnection endConnection = Assert.IsType<A.EndConnection>(element
                .NonVisualConnectionShapeProperties.NonVisualConnectorShapeDrawingProperties!
                .GetFirstChild<A.EndConnection>());
            Assert.Equal(start.Id, startConnection.Id!.Value);
            Assert.Equal(end.Id, endConnection.Id!.Value);
            Assert.Equal(3U, startConnection.Index!.Value);
            Assert.Equal(1U, endConnection.Index!.Value);
            Assert.Empty(presentation.ValidateDocument());
            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            Assert.Equal(File.ReadAllBytes(ConnectedFixturePath),
                presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ImportedConnectorAttachmentEdit_RemainsLossBlocked() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(ConnectedFixturePath);
            PowerPointConnectionShape connector = Assert.Single(presentation.Slides[0].Shapes
                .OfType<PowerPointConnectionShape>());
            P.ConnectionShape element = Assert.IsType<P.ConnectionShape>(connector.Element);
            A.StartConnection start = element.NonVisualConnectionShapeProperties!
                .NonVisualConnectorShapeDrawingProperties!.GetFirstChild<A.StartConnection>()!;

            start.Index = 0U;

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding => finding.Code == "PPT-WRITE-IMPORT-LOSS");
        }

        [Fact]
        public void NeutralReader_DecodesShapeSpecificAdjustmentSlots() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(AdjustedShapesFixturePath);
            LegacyPptSlide slide = Assert.Single(legacy.Slides);

            Assert.Equal(6480, slide.Shapes.Single(shape => shape.Text == "Round")
                .Geometry.AdjustmentValues[0]);
            Assert.Equal(5400, slide.Shapes.Single(shape => shape.Text == "Chevron")
                .Geometry.AdjustmentValues[0]);
            LegacyPptShape arrow = slide.Shapes.Single(shape => shape.Text == "Arrow");
            Assert.Equal(7559, arrow.Geometry.AdjustmentValues[0]);
            Assert.Equal(12960, arrow.Geometry.AdjustmentValues[1]);
            Assert.Equal(8640, slide.Shapes.Single(shape => shape.Text == "Donut")
                .Geometry.AdjustmentValues[0]);
            Assert.Equal(6480, slide.Shapes.Single(shape => shape.Text == "Trapezoid")
                .Geometry.AdjustmentValues[0]);
            LegacyPptShape arc = slide.Shapes.Single(shape => shape.Text == "Arc");
            Assert.Equal(163074539, arc.Geometry.AdjustmentValues[0]);
            Assert.Equal(32614907, arc.Geometry.AdjustmentValues[1]);
        }

        [Fact]
        public void NeutralReader_DecodesOffsetShadowProperties() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(ShadowFixturePath);
            LegacyPptShape[] shapes = Assert.Single(legacy.Slides).Shapes
                .OrderBy(shape => shape.Bounds.Left)
                .ToArray();

            Assert.Equal(2, shapes.Length);
            Assert.All(shapes, shape => Assert.True(shape.Style.HasProjectableShadow));
            Assert.Equal("445566", shapes[0].ShadowColor);
            Assert.Equal(39336D / 65536D, shapes[0].Style.ShadowOpacity);
            Assert.Equal(26640, shapes[0].Style.ShadowOffsetXEmus);
            Assert.Equal(26640, shapes[0].Style.ShadowOffsetYEmus);
            Assert.Equal("000000", shapes[1].ShadowColor);
            Assert.Equal(-18000, shapes[1].Style.ShadowOffsetXEmus);
            Assert.Equal(18000, shapes[1].Style.ShadowOffsetYEmus);
            Assert.DoesNotContain(legacy.Diagnostics, diagnostic =>
                diagnostic.Code == "PPT-SHAPE-STYLE-PARTIAL");
        }

        [Fact]
        public void NormalLoad_ProjectsOffsetShadowsAndPreservesBinaryExactly() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(ShadowFixturePath);
            PowerPointSlide slide = Assert.Single(presentation.Slides);

            A.OuterShadow first = GetOuterShadow(slide, "Shadow 45");
            Assert.Equal(37675L, first.Distance!.Value);
            Assert.Equal(2700000, first.Direction!.Value);
            Assert.Equal(0L, first.BlurRadius!.Value);
            A.RgbColorModelHex firstColor = Assert.IsType<A.RgbColorModelHex>(first.FirstChild);
            Assert.Equal("445566", firstColor.Val!.Value);
            Assert.Equal(60022, firstColor.GetFirstChild<A.Alpha>()!.Val!.Value);

            A.OuterShadow second = GetOuterShadow(slide, "Shadow 135");
            Assert.Equal(25456L, second.Distance!.Value);
            Assert.Equal(8100000, second.Direction!.Value);
            Assert.Equal("000000", Assert.IsType<A.RgbColorModelHex>(second.FirstChild).Val!.Value);
            Assert.Empty(presentation.ValidateDocument());
            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            Assert.Equal(File.ReadAllBytes(ShadowFixturePath),
                presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ImportedRotatingShadowEdit_RemainsLossBlocked() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(ShadowFixturePath);
            A.OuterShadow shadow = GetOuterShadow(presentation.Slides[0], "Shadow 45");

            shadow.RotateWithShape = true;

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding => finding.Code == "PPT-WRITE-IMPORT-LOSS");
        }

        [Fact]
        public void NormalLoad_ProjectsOnlyProvenExactPresetAdjustments() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(
                AdjustedShapesFixturePath);
            PowerPointSlide slide = Assert.Single(presentation.Slides);

            Assert.Equal("val 30000", GetAdjustmentFormula(slide, "Round", "adj"));
            Assert.Equal("val 40000", GetAdjustmentFormula(slide, "Donut", "adj"));
            Assert.Null(GetAdjustmentFormula(slide, "Chevron", "adj"));
            Assert.Null(GetAdjustmentFormula(slide, "Arrow", "adj1"));
            Assert.Null(GetAdjustmentFormula(slide, "Trapezoid", "adj"));
            Assert.Null(GetAdjustmentFormula(slide, "Arc", "adj1"));
            Assert.Empty(presentation.ValidateDocument());
            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            Assert.Equal(File.ReadAllBytes(AdjustedShapesFixturePath),
                presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ImportedExactPresetAdjustmentEdit_UsesIncrementalFoptRewrite() {
            byte[] originalBytes = File.ReadAllBytes(
                AdjustedShapesFixturePath);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(
                AdjustedShapesFixturePath);
            PowerPointTextBox shape = presentation.Slides[0].TextBoxes.Single(item =>
                item.Text == "Round");
            P.Shape element = Assert.IsType<P.Shape>(shape.Element);
            A.ShapeGuide adjustment = Assert.Single(element.ShapeProperties!
                .GetFirstChild<A.PresetGeometry>()!.AdjustValueList!
                .Elements<A.ShapeGuide>());

            adjustment.Formula = "val 25000";

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                originalBytes);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));

            LegacyPptShape originalRound = Assert.Single(original.Slides)
                .Shapes.Single(item => item.Text == "Round");
            LegacyPptShape savedRound = Assert.Single(saved.Slides).Shapes
                .Single(item => item.Text == "Round");
            Assert.Equal(5400,
                savedRound.Geometry.AdjustmentValues[0]);
            Assert.Equal(originalRound.Style.Properties
                    .Where(property => property.PropertyId != 0x0147)
                    .Select(property => (property.RawOperationId,
                        property.Value)),
                savedRound.Style.Properties
                    .Where(property => property.PropertyId != 0x0147)
                    .Select(property => (property.RawOperationId,
                        property.Value)));
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
        }

        private static string? GetAdjustmentFormula(PowerPointSlide slide, string text,
            string guideName) {
            PowerPointTextBox shape = slide.TextBoxes.Single(item => item.Text == text);
            P.Shape element = Assert.IsType<P.Shape>(shape.Element);
            return element.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.AdjustValueList?
                .Elements<A.ShapeGuide>()
                .FirstOrDefault(guide => guide.Name?.Value == guideName)?.Formula?.Value;
        }

        private static A.OuterShadow GetOuterShadow(PowerPointSlide slide, string text) {
            PowerPointTextBox shape = slide.TextBoxes.Single(item => item.Text == text);
            P.Shape element = Assert.IsType<P.Shape>(shape.Element);
            return Assert.IsType<A.OuterShadow>(element.ShapeProperties?
                .GetFirstChild<A.EffectList>()?.GetFirstChild<A.OuterShadow>());
        }
    }
}
