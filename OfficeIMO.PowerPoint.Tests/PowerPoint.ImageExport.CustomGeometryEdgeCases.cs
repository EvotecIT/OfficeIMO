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

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.FromRgb(17, 34, 51), image!.GetPixel(40, 40));
            string svg = System.Text.Encoding.UTF8.GetString(slide.ExportImage(OfficeImageExportFormat.Svg).Bytes);
            Assert.Contains("#445566", svg, StringComparison.OrdinalIgnoreCase);

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
