using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        [Fact]
        public void PowerPointSlide_AuthorsSharedCustomGeometryThroughPublicApi() {
            OfficeShape sharedPath = OfficeShape.Path(
                OfficePathCommand.MoveTo(0, 50),
                OfficePathCommand.QuadraticBezierTo(25, 0, 50, 25),
                OfficePathCommand.CubicBezierTo(70, 45, 82, 100, 100, 50),
                OfficePathCommand.LineTo(76, 100),
                OfficePathCommand.LineTo(24, 100),
                OfficePathCommand.Close());
            sharedPath.FillColor = OfficeColor.FromRgb(14, 165, 233);
            sharedPath.StrokeColor = OfficeColor.FromRgb(12, 74, 110);
            sharedPath.StrokeWidth = 2.25D;

            using var stream = new System.IO.MemoryStream();
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
                presentation.SlideSize.SetSizePoints(180, 140);
                PowerPointAutoShape shape = presentation.AddSlide().AddCustomGeometryPoints(
                    sharedPath, 30, 20, 120, 100, "Shared freeform");
                Assert.Equal("Shared freeform", shape.Name);
                Assert.Null(shape.ShapeType);
                Assert.Equal("0EA5E9", shape.FillColor);
                Assert.Equal("0C4A6E", shape.OutlineColor);
                Assert.Equal(2.25D, shape.OutlineWidthPoints);
                presentation.Save();
            }

            stream.Position = 0;
            using PowerPointPresentation reopened = PowerPointPresentation.Load(stream);
            PowerPointAutoShape authored = Assert.IsType<PowerPointAutoShape>(
                Assert.Single(reopened.Slides[0].Shapes));
            Assert.Null(authored.ShapeType);
            OfficeImageExportResult svg = reopened.Slides[0].ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = reopened.Slides[0].ExportImage(OfficeImageExportFormat.Png);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("Q", svgText, StringComparison.Ordinal);
            Assert.Contains("C", svgText, StringComparison.Ordinal);
            Assert.Contains("#0EA5E9", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(14, 165, 233)) > 100);
        }

        [Fact]
        public void PowerPointSlide_RejectsFilledMultiContourEvenOddGeometry() {
            OfficeShape geometry = OfficeShape.Path(
                OfficePathCommand.MoveTo(0, 0),
                OfficePathCommand.LineTo(100, 0),
                OfficePathCommand.LineTo(100, 100),
                OfficePathCommand.Close(),
                OfficePathCommand.MoveTo(25, 25),
                OfficePathCommand.LineTo(75, 25),
                OfficePathCommand.LineTo(75, 75),
                OfficePathCommand.Close());
            geometry.FillColor = OfficeColor.FromRgb(14, 165, 233);
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();

            NotSupportedException error = Assert.Throws<NotSupportedException>(() =>
                slide.AddCustomGeometryPoints(geometry, 10, 10, 100, 100));
            Assert.Contains("even-odd", error.Message,
                StringComparison.OrdinalIgnoreCase);

            geometry.FillRule = OfficeFillRule.NonZero;
            Assert.NotNull(slide.AddCustomGeometryPoints(geometry,
                10, 10, 100, 100));
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void PowerPointSlide_ProjectsCustomGeometryWithNonZeroFillRule() {
            OfficeShape geometry = OfficeShape.Path(
                OfficePathCommand.MoveTo(0, 0),
                OfficePathCommand.LineTo(100, 0),
                OfficePathCommand.LineTo(100, 100),
                OfficePathCommand.LineTo(0, 100),
                OfficePathCommand.Close(),
                OfficePathCommand.MoveTo(25, 25),
                OfficePathCommand.LineTo(75, 25),
                OfficePathCommand.LineTo(75, 75),
                OfficePathCommand.LineTo(25, 75),
                OfficePathCommand.Close());
            geometry.FillRule = OfficeFillRule.NonZero;
            geometry.FillColor = OfficeColor.FromRgb(14, 165, 233);
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddCustomGeometryPoints(geometry, 10, 10, 100, 100);

            PowerPointSlideVisualSnapshot snapshot =
                slide.CreateVisualSnapshot(new PowerPointImageExportOptions {
                    IncludeSlideBackground = false
                });

            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>(), element => element.X == 10D
                    && element.Y == 10D);
            Assert.Equal(OfficeFillRule.NonZero, rendered.Shape.FillRule);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void PowerPointSlide_AuthorsCombinedCustomGeometryStrokeOpacity() {
            OfficeShape polygon = OfficeShape.Polygon(
                new OfficePoint(0, 0),
                new OfficePoint(100, 0),
                new OfficePoint(50, 100));
            polygon.StrokeColor = OfficeColor.FromRgba(37, 99, 235, 128);
            polygon.StrokeOpacity = 0.5D;
            polygon.StrokeWidth = 3D;

            using var stream = new System.IO.MemoryStream();
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
                PowerPointAutoShape shape = presentation.AddSlide().AddCustomGeometryPoints(
                    polygon, 20, 20, 120, 80);
                Assert.Equal(75, shape.OutlineTransparency);
                A.Alpha alpha = shape.Element.Descendants<A.Outline>().Single()
                    .Descendants<A.Alpha>().Single();
                Assert.InRange(alpha.Val!.Value, 25097, 25099);
                presentation.Save();
            }

            stream.Position = 0;
            using PowerPointPresentation reopened = PowerPointPresentation.Load(stream);
            PowerPointAutoShape authored = Assert.IsType<PowerPointAutoShape>(
                Assert.Single(reopened.Slides[0].Shapes));
            Assert.Equal(75, authored.OutlineTransparency);
        }

        [Fact]
        public void PowerPointSlide_AuthorsCombinedCustomGeometryFillOpacity() {
            OfficeShape polygon = OfficeShape.Polygon(
                new OfficePoint(0, 0),
                new OfficePoint(100, 0),
                new OfficePoint(50, 100));
            polygon.FillColor = OfficeColor.FromRgba(14, 165, 233, 128);
            polygon.FillOpacity = 0.5D;

            using var stream = new System.IO.MemoryStream();
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
                PowerPointAutoShape shape = presentation.AddSlide().AddCustomGeometryPoints(
                    polygon, 20, 20, 120, 80);
                Assert.Equal(75, shape.FillTransparency);
                A.Alpha alpha = shape.Element.Descendants<A.SolidFill>().First()
                    .Descendants<A.Alpha>().Single();
                Assert.InRange(alpha.Val!.Value, 25097, 25099);
                presentation.Save();
            }

            stream.Position = 0;
            using PowerPointPresentation reopened = PowerPointPresentation.Load(stream);
            PowerPointAutoShape authored = Assert.IsType<PowerPointAutoShape>(
                Assert.Single(reopened.Slides[0].Shapes));
            Assert.Equal(75, authored.FillTransparency);
        }

        [Fact]
        public void PowerPointShape_OutlineOpacityPreservesSchemeColorChoice() {
            using var stream = new System.IO.MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointAutoShape shape = presentation.AddSlide().AddRectanglePoints(
                20, 20, 120, 80);
            shape.OutlineColor = "112233";
            A.SolidFill solid = shape.Element.Descendants<A.Outline>().Single()
                .GetFirstChild<A.SolidFill>()!;
            solid.RemoveAllChildren();
            var scheme = new A.SchemeColor { Val = A.SchemeColorValues.Accent1 };
            solid.Append(scheme);

            shape.SetOutlineOpacity(0.4D);

            Assert.Same(scheme, Assert.Single(solid.ChildElements));
            Assert.Equal(40000, scheme.GetFirstChild<A.Alpha>()!.Val!.Value);
            Assert.Equal(60, shape.OutlineTransparency);
            Assert.Null(solid.RgbColorModelHex);

            shape.SetOutlineOpacity(null);
            Assert.Same(scheme, Assert.Single(solid.ChildElements));
            Assert.Null(scheme.GetFirstChild<A.Alpha>());
        }

        [Fact]
        public void PowerPointShape_OutlineTransparencyPreservesThemeLineReference() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            presentation.SetThemeColorForAllMasters(
                PowerPointThemeColor.Accent2, "123456");
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectanglePoints(
                20, 20, 120, 80);
            Shape openXmlShape = Assert.IsType<Shape>(shape.Element);
            A.Outline themeOutline = slide.SlidePart.SlideLayoutPart!
                .SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .FormatScheme!.LineStyleList!.Elements<A.Outline>().First();
            themeOutline.Width = 38100;
            themeOutline.CapType = A.LineCapValues.Round;
            themeOutline.RemoveAllChildren<A.PresetDash>();
            themeOutline.RemoveAllChildren<A.Round>();
            themeOutline.RemoveAllChildren<A.Bevel>();
            themeOutline.RemoveAllChildren<A.Miter>();
            A.SolidFill themeFill = themeOutline.GetFirstChild<A.SolidFill>()!;
            A.PresetDash themeDash = themeOutline.InsertAfter(
                new A.PresetDash { Val = A.PresetLineDashValues.Dash },
                themeFill)!;
            themeOutline.InsertAfter(new A.Round(), themeDash);
            var scheme = new A.SchemeColor {
                Val = A.SchemeColorValues.Accent2
            };
            openXmlShape.ShapeStyle = new ShapeStyle(
                new A.LineReference(scheme) { Index = 1U },
                new A.FillReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent1
                }) { Index = 1U },
                new A.EffectReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent1
                }) { Index = 0U },
                new A.FontReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Dark1
                }) { Index = A.FontCollectionIndexValues.Minor });

            shape.OutlineTransparency = 40;

            A.Outline outline = openXmlShape.ShapeProperties!
                .GetFirstChild<A.Outline>()!;
            Assert.Null(outline.GetFirstChild<A.SolidFill>());
            Assert.Same(scheme, openXmlShape.ShapeStyle.LineReference!
                .GetFirstChild<A.SchemeColor>());
            Assert.Equal(60000, scheme.GetFirstChild<A.Alpha>()!.Val!.Value);
            Assert.Equal(40, shape.OutlineTransparency);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(
                new PowerPointImageExportOptions {
                    IncludeSlideBackground = false
                });
            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>(), item =>
                    item.X == 20D && item.Y == 20D);
            OfficeColor renderedStroke = Assert.IsType<OfficeColor>(
                rendered.Shape.StrokeColor);
            Assert.Equal((byte)0x12, renderedStroke.R);
            Assert.Equal((byte)0x34, renderedStroke.G);
            Assert.Equal((byte)0x56, renderedStroke.B);
            Assert.Equal((byte)153, renderedStroke.A);
            Assert.Equal(3D, rendered.Shape.StrokeWidth);
            Assert.Equal(OfficeStrokeDashStyle.Dash,
                rendered.Shape.StrokeDashStyle);
            Assert.Equal(OfficeStrokeLineCap.Round,
                rendered.Shape.StrokeLineCap);
            Assert.Equal(OfficeStrokeLineJoin.Round,
                rendered.Shape.StrokeLineJoin);
            OfficeImageExportResult svg = slide.ExportImage(
                OfficeImageExportFormat.Svg);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("stroke-width=\"3\"", svgText,
                StringComparison.Ordinal);
            Assert.Contains("stroke-dasharray=", svgText,
                StringComparison.Ordinal);
            Assert.Contains("stroke-linecap=\"round\"", svgText,
                StringComparison.Ordinal);
            Assert.Contains("stroke-linejoin=\"round\"", svgText,
                StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointShape_FillTransparencyPreservesSchemeColorChoice() {
            using var stream = new System.IO.MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointAutoShape shape = presentation.AddSlide().AddRectanglePoints(
                20, 20, 120, 80);
            shape.FillColor = "112233";
            A.SolidFill solid = shape.Element.Descendants<ShapeProperties>().Single()
                .GetFirstChild<A.SolidFill>()!;
            solid.RemoveAllChildren();
            var scheme = new A.SchemeColor { Val = A.SchemeColorValues.Accent2 };
            solid.Append(scheme);

            shape.FillTransparency = 60;

            Assert.Same(scheme, Assert.Single(solid.ChildElements));
            Assert.Equal(40000, scheme.GetFirstChild<A.Alpha>()!.Val!.Value);
            Assert.Equal(60, shape.FillTransparency);
            Assert.Null(solid.RgbColorModelHex);

            shape.FillTransparency = null;
            Assert.Same(scheme, Assert.Single(solid.ChildElements));
            Assert.Null(scheme.GetFirstChild<A.Alpha>());
        }

        [Fact]
        public void PowerPointShape_FillTransparencyPreservesThemeFillReference() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            presentation.SetThemeColorForAllMasters(
                PowerPointThemeColor.Accent2, "123456");
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectanglePoints(
                20, 20, 120, 80);
            Shape openXmlShape = Assert.IsType<Shape>(shape.Element);
            openXmlShape.ShapeProperties!.RemoveAllChildren<A.SolidFill>();
            var scheme = new A.SchemeColor {
                Val = A.SchemeColorValues.Accent2
            };
            openXmlShape.ShapeStyle = new ShapeStyle(
                new A.LineReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent1
                }) { Index = 1U },
                new A.FillReference(scheme) { Index = 1U },
                new A.EffectReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent1
                }) { Index = 0U },
                new A.FontReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Dark1
                }) { Index = A.FontCollectionIndexValues.Minor });

            shape.FillTransparency = 40;

            Assert.Null(openXmlShape.ShapeProperties!
                .GetFirstChild<A.SolidFill>());
            Assert.Same(scheme, openXmlShape.ShapeStyle.FillReference!
                .GetFirstChild<A.SchemeColor>());
            Assert.Equal(60000, scheme.GetFirstChild<A.Alpha>()!.Val!.Value);
            Assert.Equal(40, shape.FillTransparency);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(
                new PowerPointImageExportOptions {
                    IncludeSlideBackground = false
                });
            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>(), item =>
                    item.X == 20D && item.Y == 20D);
            OfficeColor renderedFill = Assert.IsType<OfficeColor>(
                rendered.Shape.FillColor);
            Assert.Equal((byte)0x12, renderedFill.R);
            Assert.Equal((byte)0x34, renderedFill.G);
            Assert.Equal((byte)0x56, renderedFill.B);
            Assert.Equal((byte)153, renderedFill.A);
        }

        [Fact]
        public void PowerPointSlide_AuthorsSharedPolygonAndRejectsNonFreeformDescriptors() {
            OfficeShape polygon = OfficeShape.Polygon(
                new OfficePoint(0, 0),
                new OfficePoint(100, 0),
                new OfficePoint(75, 100),
                new OfficePoint(25, 100));
            polygon.FillColor = OfficeColor.FromRgb(34, 197, 94);

            using var stream = new System.IO.MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddCustomGeometryCm(
                polygon, 1, 1, 4, 3, "Shared polygon");
            Assert.Null(shape.ShapeType);
            Assert.Throws<ArgumentException>(() => slide.AddCustomGeometryPoints(
                OfficeShape.Rectangle(10, 10), 0, 0, 10, 10));
        }

        [Fact]
        public void PowerPointSlide_ProjectsCustomGeometryThroughSharedDrawingPath() {
            using var stream = new System.IO.MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 120);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape freeform = slide.AddShapePoints(A.ShapeTypeValues.Rectangle, 30, 24, 88, 56);
            freeform.FillColor = "C084FC";
            freeform.OutlineColor = "6B21A8";
            freeform.OutlineWidthPoints = 2D;

            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Shape>().Last();
            ShapeProperties properties = shape.ShapeProperties!;
            A.Transform2D transform = properties.GetFirstChild<A.Transform2D>()!;
            properties.RemoveAllChildren<A.PresetGeometry>();
            properties.InsertAfter(CreateDiamondCustomGeometry(), transform);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), element =>
                Math.Abs(element.X - 30D) < 0.000001D &&
                Math.Abs(element.Y - 24D) < 0.000001D);
            Assert.Equal(OfficeShapeKind.Path, rendered.Shape.Kind);
            Assert.Equal(88D, rendered.Shape.Width, 1);
            Assert.Equal(56D, rendered.Shape.Height, 1);
            Assert.Contains(rendered.Shape.PathCommands, command => command.Kind == OfficePathCommandKind.Close);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("#C084FC", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#6B21A8", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(192, 132, 252)) > 100);
        }

        [Fact]
        public void PowerPointSlide_ProjectsCurvedCustomGeometryThroughSharedDrawingPath() {
            using var stream = new System.IO.MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 140);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape freeform = slide.AddShapePoints(A.ShapeTypeValues.Rectangle, 20, 30, 120, 72);
            freeform.FillColor = "22C55E";
            freeform.OutlineColor = "166534";
            freeform.OutlineWidthPoints = 2D;

            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Shape>().Last();
            ShapeProperties properties = shape.ShapeProperties!;
            A.Transform2D transform = properties.GetFirstChild<A.Transform2D>()!;
            properties.RemoveAllChildren<A.PresetGeometry>();
            properties.InsertAfter(CreateCurvedCustomGeometry(), transform);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), element =>
                Math.Abs(element.X - 20D) < 0.000001D &&
                Math.Abs(element.Y - 30D) < 0.000001D);
            Assert.Equal(OfficeShapeKind.Path, rendered.Shape.Kind);
            Assert.Equal(120D, rendered.Shape.Width, 1);
            Assert.Equal(72D, rendered.Shape.Height, 1);
            Assert.Contains(rendered.Shape.PathCommands, command => command.Kind == OfficePathCommandKind.QuadraticBezierTo);
            Assert.Contains(rendered.Shape.PathCommands, command => command.Kind == OfficePathCommandKind.CubicBezierTo);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("Q", svgText, StringComparison.Ordinal);
            Assert.Contains("C", svgText, StringComparison.Ordinal);
            Assert.Contains("#22C55E", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#166534", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(34, 197, 94)) > 100);
        }

        [Fact]
        public void PowerPointSlide_ProjectsGuidedCustomGeometryThroughSharedDrawingPath() {
            using var stream = new System.IO.MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(190, 150);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape freeform = slide.AddShapePoints(A.ShapeTypeValues.Rectangle, 32, 28, 120, 90);
            freeform.FillColor = "F59E0B";
            freeform.OutlineColor = "92400E";
            freeform.OutlineWidthPoints = 2D;

            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Shape>().Last();
            ShapeProperties properties = shape.ShapeProperties!;
            A.Transform2D transform = properties.GetFirstChild<A.Transform2D>()!;
            properties.RemoveAllChildren<A.PresetGeometry>();
            properties.InsertAfter(CreateGuidedCustomGeometry(), transform);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), element =>
                Math.Abs(element.X - 32D) < 0.000001D &&
                Math.Abs(element.Y - 28D) < 0.000001D);
            Assert.Equal(OfficeShapeKind.Path, rendered.Shape.Kind);
            Assert.Equal(120D, rendered.Shape.Width, 1);
            Assert.Equal(90D, rendered.Shape.Height, 1);
            Assert.Equal(OfficePathCommandKind.MoveTo, rendered.Shape.PathCommands[0].Kind);
            AssertCustomGeometryPointNear(rendered.Shape.PathCommands[0].Point, 30D, 0D);
            Assert.Equal(OfficePathCommandKind.LineTo, rendered.Shape.PathCommands[1].Kind);
            AssertCustomGeometryPointNear(rendered.Shape.PathCommands[1].Point, 90D, 0D);
            Assert.Equal(OfficePathCommandKind.LineTo, rendered.Shape.PathCommands[2].Kind);
            AssertCustomGeometryPointNear(rendered.Shape.PathCommands[2].Point, 120D, 45D);
            Assert.Contains(rendered.Shape.PathCommands, command => command.Kind == OfficePathCommandKind.Close);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("#F59E0B", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#92400E", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(245, 158, 11)) > 100);
        }

        [Fact]
        public void PowerPointSlide_ProjectsArcCustomGeometryThroughSharedDrawingPath() {
            using var stream = new System.IO.MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(190, 150);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape freeform = slide.AddShapePoints(A.ShapeTypeValues.Rectangle, 36, 30, 120, 80);
            freeform.FillColor = "38BDF8";
            freeform.OutlineColor = "075985";
            freeform.OutlineWidthPoints = 2D;

            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Shape>().Last();
            ShapeProperties properties = shape.ShapeProperties!;
            A.Transform2D transform = properties.GetFirstChild<A.Transform2D>()!;
            properties.RemoveAllChildren<A.PresetGeometry>();
            properties.InsertAfter(CreateArcCustomGeometry(), transform);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), element =>
                Math.Abs(element.X - 36D) < 0.000001D &&
                Math.Abs(element.Y - 30D) < 0.000001D);
            Assert.Equal(OfficeShapeKind.Path, rendered.Shape.Kind);
            Assert.Equal(120D, rendered.Shape.Width, 1);
            Assert.Equal(80D, rendered.Shape.Height, 1);
            OfficePathCommand arcCommand = Assert.Single(rendered.Shape.PathCommands, command => command.Kind == OfficePathCommandKind.CubicBezierTo);
            AssertCustomGeometryPointNear(arcCommand.Point, 60D, 80D);
            Assert.Contains(rendered.Shape.PathCommands, command => command.Kind == OfficePathCommandKind.Close);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("C", svgText, StringComparison.Ordinal);
            Assert.Contains("#38BDF8", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#075985", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(56, 189, 248)) > 100);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTrigonometricGuidedCustomGeometryThroughSharedDrawingPath() {
            using var stream = new System.IO.MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(190, 150);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape freeform = slide.AddShapePoints(A.ShapeTypeValues.Rectangle, 34, 26, 100, 100);
            freeform.FillColor = "A3E635";
            freeform.OutlineColor = "3F6212";
            freeform.OutlineWidthPoints = 2D;

            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Shape>().Last();
            ShapeProperties properties = shape.ShapeProperties!;
            A.Transform2D transform = properties.GetFirstChild<A.Transform2D>()!;
            properties.RemoveAllChildren<A.PresetGeometry>();
            properties.InsertAfter(CreateTrigonometricGuidedCustomGeometry(), transform);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), element =>
                Math.Abs(element.X - 34D) < 0.000001D &&
                Math.Abs(element.Y - 26D) < 0.000001D);
            Assert.Equal(OfficeShapeKind.Path, rendered.Shape.Kind);
            Assert.Equal(100D, rendered.Shape.Width, 1);
            Assert.Equal(100D, rendered.Shape.Height, 1);
            AssertCustomGeometryPointNear(rendered.Shape.PathCommands[1].Point, 70.71067811865476D, 25D);
            AssertCustomGeometryPointNear(rendered.Shape.PathCommands[2].Point, 100D, 70.71067811865476D);
            AssertCustomGeometryPointNear(rendered.Shape.PathCommands[3].Point, 35.35533905932738D, 100D);
            Assert.Contains(rendered.Shape.PathCommands, command => command.Kind == OfficePathCommandKind.Close);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("#A3E635", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#3F6212", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(163, 230, 53)) > 100);
        }

        [Fact]
        public void PowerPointSlide_ProjectsAngleDerivedGuidedCustomGeometryThroughSharedDrawingPath() {
            using var stream = new System.IO.MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(190, 150);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape freeform = slide.AddShapePoints(A.ShapeTypeValues.Rectangle, 30, 22, 100, 100);
            freeform.FillColor = "F472B6";
            freeform.OutlineColor = "9D174D";
            freeform.OutlineWidthPoints = 2D;

            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Shape>().Last();
            ShapeProperties properties = shape.ShapeProperties!;
            A.Transform2D transform = properties.GetFirstChild<A.Transform2D>()!;
            properties.RemoveAllChildren<A.PresetGeometry>();
            properties.InsertAfter(CreateAngleDerivedGuidedCustomGeometry(), transform);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), element =>
                Math.Abs(element.X - 30D) < 0.000001D &&
                Math.Abs(element.Y - 22D) < 0.000001D);
            Assert.Equal(OfficeShapeKind.Path, rendered.Shape.Kind);
            Assert.Equal(100D, rendered.Shape.Width, 1);
            Assert.Equal(100D, rendered.Shape.Height, 1);
            AssertCustomGeometryPointNear(rendered.Shape.PathCommands[1].Point, 60D, 80D);
            AssertCustomGeometryPointNear(rendered.Shape.PathCommands[2].Point, 100D, 80D);
            Assert.Contains(rendered.Shape.PathCommands, command => command.Kind == OfficePathCommandKind.Close);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("#F472B6", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#9D174D", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(244, 114, 182)) > 100);
        }

        [Fact]
        public void PowerPointSlide_CustomGeometryPreservesDeclaredPathCanvas() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            presentation.SlideSize.SetSizePoints(180, 140);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape freeform = slide.AddShapePoints(
                A.ShapeTypeValues.Rectangle, 30, 24, 80, 60);
            freeform.FillColor = "0EA5E9";
            freeform.Rotation = 25D;

            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<Shape>().Last();
            ShapeProperties properties = shape.ShapeProperties!;
            A.Transform2D transform = properties.GetFirstChild<A.Transform2D>()!;
            properties.RemoveAllChildren<A.PresetGeometry>();
            properties.InsertAfter(CreateInsetCustomGeometry(), transform);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>(), element =>
                    Math.Abs(element.X - 30D) < 0.000001D
                    && Math.Abs(element.Y - 24D) < 0.000001D);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            Assert.Equal(80D, rendered.Shape.Width);
            Assert.Equal(60D, rendered.Shape.Height);
            AssertCustomGeometryPointNear(rendered.Shape.PathCommands[0].Point,
                20D, 15D);
            OfficePoint center = rendered.Shape.Transform!.Value.TransformPoint(
                new OfficePoint(40D, 30D));
            AssertCustomGeometryPointNear(center, 40D, 30D);
        }

        [Fact]
        public void PowerPointSlide_CustomGeometryHonorsPerPathFillAndStroke() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            presentation.SlideSize.SetSizePoints(180, 140);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape freeform = slide.AddShapePoints(
                A.ShapeTypeValues.Rectangle, 20, 20, 120, 80);
            freeform.FillColor = "22C55E";
            freeform.OutlineColor = "1E3A8A";
            freeform.OutlineWidthPoints = 3D;

            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<Shape>().Last();
            ShapeProperties properties = shape.ShapeProperties!;
            A.Transform2D transform = properties.GetFirstChild<A.Transform2D>()!;
            properties.RemoveAllChildren<A.PresetGeometry>();
            properties.InsertAfter(CreatePerPathStyledCustomGeometry(), transform);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeDrawingShape[] rendered = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(element => element.X == 20D && element.Y == 20D)
                .ToArray();
            OfficeImageExportResult svg = slide.ExportImage(
                OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.Equal(2, rendered.Length);
            Assert.Null(rendered[0].Shape.FillColor);
            Assert.Equal(OfficeColor.FromRgb(30, 58, 138),
                rendered[0].Shape.StrokeColor);
            Assert.Equal(OfficeColor.FromRgb(34, 197, 94),
                rendered[1].Shape.FillColor);
            Assert.Null(rendered[1].Shape.StrokeColor);
            Assert.Equal(0D, rendered[1].Shape.StrokeWidth);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Equal(2, svgText.Split(new[] { "<path" },
                StringSplitOptions.None).Length - 1);
            Assert.Contains("fill=\"none\"", svgText, StringComparison.Ordinal);
            Assert.Contains("stroke=\"none\"", svgText, StringComparison.Ordinal);
        }

        private static A.CustomGeometry CreateDiamondCustomGeometry() {
            return new A.CustomGeometry(
                new A.PathList(
                    new A.Path(
                        new A.MoveTo(new A.Point { X = "50000", Y = "0" }),
                        new A.LineTo(new A.Point { X = "100000", Y = "50000" }),
                        new A.LineTo(new A.Point { X = "50000", Y = "100000" }),
                        new A.LineTo(new A.Point { X = "0", Y = "50000" }),
                        new A.CloseShapePath()) {
                        Width = 100000L,
                        Height = 100000L
                    }));
        }

        private static A.CustomGeometry CreateInsetCustomGeometry() {
            return new A.CustomGeometry(
                new A.PathList(
                    new A.Path(
                        new A.MoveTo(new A.Point { X = "25000", Y = "25000" }),
                        new A.LineTo(new A.Point { X = "75000", Y = "25000" }),
                        new A.LineTo(new A.Point { X = "75000", Y = "75000" }),
                        new A.LineTo(new A.Point { X = "25000", Y = "75000" }),
                        new A.CloseShapePath()) {
                        Width = 100000L,
                        Height = 100000L
                    }));
        }

        private static A.CustomGeometry CreatePerPathStyledCustomGeometry() {
            return new A.CustomGeometry(
                new A.PathList(
                    new A.Path(
                        new A.MoveTo(new A.Point { X = "10000", Y = "10000" }),
                        new A.LineTo(new A.Point { X = "90000", Y = "10000" }),
                        new A.LineTo(new A.Point { X = "90000", Y = "40000" }),
                        new A.LineTo(new A.Point { X = "10000", Y = "40000" }),
                        new A.CloseShapePath()) {
                        Width = 100000L,
                        Height = 100000L,
                        Fill = A.PathFillModeValues.None,
                        Stroke = true
                    },
                    new A.Path(
                        new A.MoveTo(new A.Point { X = "10000", Y = "60000" }),
                        new A.LineTo(new A.Point { X = "90000", Y = "60000" }),
                        new A.LineTo(new A.Point { X = "90000", Y = "90000" }),
                        new A.LineTo(new A.Point { X = "10000", Y = "90000" }),
                        new A.CloseShapePath()) {
                        Width = 100000L,
                        Height = 100000L,
                        Fill = A.PathFillModeValues.Norm,
                        Stroke = false
                    }));
        }

        private static A.CustomGeometry CreateCurvedCustomGeometry() {
            return new A.CustomGeometry(
                new A.PathList(
                    new A.Path(
                        new A.MoveTo(new A.Point { X = "0", Y = "50000" }),
                        new A.QuadraticBezierCurveTo(
                            new A.Point { X = "25000", Y = "0" },
                            new A.Point { X = "50000", Y = "50000" }),
                        new A.CubicBezierCurveTo(
                            new A.Point { X = "65000", Y = "100000" },
                            new A.Point { X = "85000", Y = "100000" },
                            new A.Point { X = "100000", Y = "50000" }),
                        new A.LineTo(new A.Point { X = "100000", Y = "100000" }),
                        new A.LineTo(new A.Point { X = "0", Y = "100000" }),
                        new A.CloseShapePath()) {
                        Width = 100000L,
                        Height = 100000L
                    }));
        }

        private static A.CustomGeometry CreateGuidedCustomGeometry() {
            return new A.CustomGeometry(
                new A.AdjustValueList(
                    new A.ShapeGuide { Name = "inset", Formula = "val 25000" }),
                new A.ShapeGuideList(
                    new A.ShapeGuide { Name = "leftGuide", Formula = "pin l inset r" },
                    new A.ShapeGuide { Name = "rightGuide", Formula = "*/ w 3 4" },
                    new A.ShapeGuide { Name = "centerGuide", Formula = "+/ t b 2" }),
                new A.PathList(
                    new A.Path(
                        new A.MoveTo(new A.Point { X = "leftGuide", Y = "t" }),
                        new A.LineTo(new A.Point { X = "rightGuide", Y = "t" }),
                        new A.LineTo(new A.Point { X = "r", Y = "centerGuide" }),
                        new A.LineTo(new A.Point { X = "rightGuide", Y = "b" }),
                        new A.LineTo(new A.Point { X = "leftGuide", Y = "b" }),
                        new A.LineTo(new A.Point { X = "l", Y = "centerGuide" }),
                        new A.CloseShapePath()) {
                        Width = 100000L,
                        Height = 100000L
                    }));
        }

        private static A.CustomGeometry CreateArcCustomGeometry() {
            return new A.CustomGeometry(
                new A.PathList(
                    new A.Path(
                        new A.MoveTo(new A.Point { X = "100000", Y = "0" }),
                        new A.ArcTo {
                            WidthRadius = "50000",
                            HeightRadius = "100000",
                            StartAngle = "0",
                            SwingAngle = "5400000"
                        },
                        new A.LineTo(new A.Point { X = "0", Y = "100000" }),
                        new A.LineTo(new A.Point { X = "0", Y = "0" }),
                        new A.CloseShapePath()) {
                        Width = 100000L,
                        Height = 100000L
                    }));
        }

        private static A.CustomGeometry CreateTrigonometricGuidedCustomGeometry() {
            return new A.CustomGeometry(
                new A.AdjustValueList(
                    new A.ShapeGuide { Name = "angle45", Formula = "val 2700000" }),
                new A.ShapeGuideList(
                    new A.ShapeGuide { Name = "sinX", Formula = "sin w angle45" },
                    new A.ShapeGuide { Name = "cosY", Formula = "cos h angle45" },
                    new A.ShapeGuide { Name = "tanY", Formula = "tan wd4 angle45" },
                    new A.ShapeGuide { Name = "diagX", Formula = "mod wd4 hd4 0" }),
                new A.PathList(
                    new A.Path(
                        new A.MoveTo(new A.Point { X = "l", Y = "t" }),
                        new A.LineTo(new A.Point { X = "sinX", Y = "tanY" }),
                        new A.LineTo(new A.Point { X = "r", Y = "cosY" }),
                        new A.LineTo(new A.Point { X = "diagX", Y = "b" }),
                        new A.CloseShapePath()) {
                        Width = 100000L,
                        Height = 100000L
                    }));
        }

        private static A.CustomGeometry CreateAngleDerivedGuidedCustomGeometry() {
            return new A.CustomGeometry(
                new A.ShapeGuideList(
                    new A.ShapeGuide { Name = "vectorX", Formula = "val 30000" },
                    new A.ShapeGuide { Name = "vectorY", Formula = "val 40000" },
                    new A.ShapeGuide { Name = "angle", Formula = "at2 vectorX vectorY" },
                    new A.ShapeGuide { Name = "xOffset", Formula = "cat2 w vectorX vectorY" },
                    new A.ShapeGuide { Name = "yOffset", Formula = "sat2 h vectorX vectorY" },
                    new A.ShapeGuide { Name = "rightY", Formula = "sin h angle" }),
                new A.PathList(
                    new A.Path(
                        new A.MoveTo(new A.Point { X = "l", Y = "t" }),
                        new A.LineTo(new A.Point { X = "xOffset", Y = "yOffset" }),
                        new A.LineTo(new A.Point { X = "r", Y = "rightY" }),
                        new A.LineTo(new A.Point { X = "l", Y = "b" }),
                        new A.CloseShapePath()) {
                        Width = 100000L,
                        Height = 100000L
                    }));
        }

        private static void AssertCustomGeometryPointNear(OfficePoint actual, double expectedX, double expectedY) {
            Assert.InRange(Math.Abs(actual.X - expectedX), 0D, 0.000001D);
            Assert.InRange(Math.Abs(actual.Y - expectedY), 0D, 0.000001D);
        }
    }
}
