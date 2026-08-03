using System;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        [Fact]
        public void PowerPointSlide_CustomGeometryArcAfterCloseStartsAtSubpathOrigin() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape source = slide.AddShapePoints(
                A.ShapeTypeValues.Rectangle, 20, 20, 120, 80);
            Shape shape = Assert.IsType<Shape>(source.Element);
            ShapeProperties properties = shape.ShapeProperties!;
            A.Transform2D transform = properties.GetFirstChild<A.Transform2D>()!;
            properties.RemoveAllChildren<A.PresetGeometry>();
            properties.InsertAfter(new A.CustomGeometry(
                new A.PathList(new A.Path(
                    new A.MoveTo(new A.Point { X = "20000", Y = "20000" }),
                    new A.LineTo(new A.Point { X = "80000", Y = "20000" }),
                    new A.CloseShapePath(),
                    new A.ArcTo {
                        WidthRadius = "20000",
                        HeightRadius = "20000",
                        StartAngle = "0",
                        SwingAngle = "5400000"
                    }) {
                    Width = 100000L,
                    Height = 100000L
                })), transform);

            PowerPointSlideVisualSnapshot snapshot =
                slide.CreateVisualSnapshot();

            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>(), element =>
                    Math.Abs(element.X - 20D) < 0.000001D
                    && Math.Abs(element.Y - 20D) < 0.000001D);
            OfficePathCommand arc = Assert.Single(rendered.Shape.PathCommands,
                command => command.Kind
                    == OfficePathCommandKind.CubicBezierTo);
            AssertCustomGeometryPointNear(arc.Point, 0D, 32D);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
        }

        [Fact]
        public void PowerPointSlide_ResolvesFixedThemeReferenceColors() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape source = slide.AddRectanglePoints(
                20, 20, 80, 40);
            Shape shape = Assert.IsType<Shape>(source.Element);
            ShapeProperties properties = shape.ShapeProperties!;
            properties.RemoveAllChildren<A.SolidFill>();
            A.FormatScheme format = slide.SlidePart.SlideLayoutPart!
                .SlideMasterPart!.ThemePart!.Theme!.ThemeElements!
                .FormatScheme!;
            format.FillStyleList!.ReplaceChild(new A.SolidFill(
                    new A.SchemeColor { Val = A.SchemeColorValues.PhColor }),
                format.FillStyleList.ChildElements[0]);
            format.LineStyleList!.ReplaceChild(new A.Outline(
                    new A.SolidFill(new A.SchemeColor {
                        Val = A.SchemeColorValues.PhColor
                    })) { Width = 25400 },
                format.LineStyleList.ChildElements[0]);
            shape.ShapeStyle = new ShapeStyle(
                new A.LineReference(new A.RgbColorModelHex {
                    Val = "445566"
                }) { Index = 1U },
                new A.FillReference(new A.RgbColorModelHex {
                    Val = "112233"
                }) { Index = 1U },
                new A.EffectReference { Index = 0U },
                new A.FontReference {
                    Index = A.FontCollectionIndexValues.Minor
                });

            PowerPointSlideVisualSnapshot snapshot =
                slide.CreateVisualSnapshot();

            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>(), element =>
                    Math.Abs(element.X - 20D) < 0.000001D
                    && Math.Abs(element.Y - 20D) < 0.000001D);
            Assert.Equal(OfficeColor.FromRgb(17, 34, 51),
                rendered.Shape.FillColor);
            Assert.Equal(OfficeColor.FromRgb(68, 85, 102),
                rendered.Shape.StrokeColor);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
        }

        [Fact]
        public void PowerPointSlide_RejectsCustomGeometryTransformOverflowBeforeMutation() {
            OfficeShape geometry = OfficeShape.Path(
                OfficePathCommand.MoveTo(0, 0),
                OfficePathCommand.LineTo(100, 0),
                OfficePathCommand.LineTo(100, 100),
                OfficePathCommand.Close());
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            const long MinimumCoordinate = -27273042329600L;
            const long MaximumCoordinate = 27273042316900L;

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                slide.AddCustomGeometry(geometry,
                    MinimumCoordinate - 1L, 0L, 100L, 100L));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                slide.AddCustomGeometry(geometry,
                    0L, MaximumCoordinate + 1L, 100L, 100L));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                slide.AddCustomGeometry(geometry,
                    0L, 0L, MaximumCoordinate + 1L, 100L));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                slide.AddCustomGeometry(geometry,
                    0L, 0L, 100L, MaximumCoordinate + 1L));
            Assert.Empty(slide.Shapes);
        }
    }
}
