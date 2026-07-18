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
                    new A.ShapeGuide { Name = "angle", Formula = "at2 vectorY vectorX" },
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
