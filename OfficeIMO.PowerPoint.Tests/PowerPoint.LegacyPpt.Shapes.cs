using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        private static string ShapeFixturePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "ShapePowerPoint.ppt");

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

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptShape savedConnector = saved.Slides[0].Shapes.Single(shape =>
                shape.Kind == LegacyPptShapeKind.Connector && shape.OfficeArtShapeType == 32);
            LegacyPptShape savedGroup = saved.Slides[0].Shapes.Single(shape =>
                shape.Kind == LegacyPptShapeKind.Group);
            Assert.Equal(originalConnector.Bounds.Left + 10, savedConnector.Bounds.Left);
            Assert.Equal(originalGroup.Bounds.Top + 10, savedGroup.Bounds.Top);
            Assert.Equal(originalGroup.Children.Select(child => child.Bounds),
                savedGroup.Children.Select(child => child.Bounds));
            Assert.Equal(original.Package.UserEdits.Count + 1, saved.Package.UserEdits.Count);
        }

        [Fact]
        public void ImportedGroupChildEdit_RemainsLossBlocked() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(ShapeFixturePath);
            PowerPointSlide slide = presentation.Slides[0];
            PowerPointGroupShape group = Assert.Single(slide.Shapes.OfType<PowerPointGroupShape>());
            PowerPointShape child = slide.GetGroupChildren(group)[0];

            child.Left += 15875;

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding => finding.Code == "PPT-WRITE-IMPORT-LOSS");
        }
    }
}
