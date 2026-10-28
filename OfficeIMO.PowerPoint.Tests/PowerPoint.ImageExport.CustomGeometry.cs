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
            sharedPath.FillRule = OfficeFillRule.NonZero;
            sharedPath.StrokeColor = OfficeColor.FromRgb(12, 74, 110);
            sharedPath.StrokeWidth = 2.25D;

            using var stream = new System.IO.MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(180),OfficeIMO.OfficeLength.Points(140));
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
            using PowerPointDocument reopened = PowerPointDocument.Load(stream);
            PowerPointAutoShape authored = Assert.IsType<PowerPointAutoShape>(
                Assert.Single(reopened.Slides[0].Shapes));
            Assert.Null(authored.ShapeType);
            OfficeImageExportResult svg = reopened.Slides[0].ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = reopened.Slides[0].ExportImage(OfficeImageExportFormat.Png);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("#0EA5E9", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(14, 165, 233)) > 100);
        }

        [Fact]
        public void PowerPointSlide_RejectsUnrepresentableStrokeWidthsBeforeMutation() {
            double[] invalidWidths = {
                -1D, double.NaN, double.PositiveInfinity, 1585D
            };
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            foreach (double width in invalidWidths) {
                OfficeShape path = OfficeShape.Path(
                    OfficePathCommand.MoveTo(0, 0),
                    OfficePathCommand.LineTo(100, 100));
                path.StrokeColor = OfficeColor.FromRgb(37, 99, 235);
                path.StrokeWidth = width;

                Assert.Throws<ArgumentOutOfRangeException>(() =>
                    slide.AddCustomGeometryPoints(path,
                        10, 10, 100, 100));
                Assert.Empty(slide.Shapes);
            }

            PowerPointAutoShape shape = slide.AddRectangle(
                OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(100), OfficeIMO.OfficeLength.Points(100));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                shape.OutlineWidthPoints = double.NaN);
            Assert.Null(shape.OutlineWidthPoints);
        }

        [Fact]
        public void PowerPointSlide_RejectsCustomGeometryGradientsBeforeMutation() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            OfficeShape fillGradient = OfficeShape.Polygon(
                new OfficePoint(0, 0), new OfficePoint(100, 0),
                new OfficePoint(50, 100));
            fillGradient.FillColor = OfficeColor.FromRgb(37, 99, 235);
            fillGradient.FillGradient = OfficeLinearGradient.Horizontal(
                OfficeColor.FromRgb(37, 99, 235),
                OfficeColor.FromRgb(14, 165, 233));
            Assert.Throws<NotSupportedException>(() =>
                slide.AddCustomGeometryPoints(fillGradient,
                    10, 10, 100, 100));

            OfficeShape strokeGradient = OfficeShape.Path(
                OfficePathCommand.MoveTo(0, 0),
                OfficePathCommand.LineTo(100, 100));
            strokeGradient.StrokeRadialGradient =
                OfficeRadialGradient.Centered(
                    OfficeColor.FromRgb(37, 99, 235),
                    OfficeColor.FromRgb(14, 165, 233));
            Assert.Throws<NotSupportedException>(() =>
                slide.AddCustomGeometryPoints(strokeGradient,
                    10, 10, 100, 100));
            Assert.Empty(slide.Shapes);
        }

        [Fact]
        public void PowerPointSlide_RejectsEveryFilledEvenOddGeometry() {
            OfficeShape geometry = OfficeShape.Path(
                OfficePathCommand.MoveTo(0, 0),
                OfficePathCommand.LineTo(100, 0),
                OfficePathCommand.LineTo(100, 100),
                OfficePathCommand.LineTo(0, 100),
                OfficePathCommand.LineTo(0, 0),
                OfficePathCommand.LineTo(100, 0),
                OfficePathCommand.LineTo(100, 100),
                OfficePathCommand.LineTo(0, 100),
                OfficePathCommand.Close());
            geometry.FillColor = OfficeColor.FromRgb(14, 165, 233);
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();

            NotSupportedException error = Assert.Throws<NotSupportedException>(() =>
                slide.AddCustomGeometryPoints(geometry, 10, 10, 100, 100));
            Assert.Contains("even-odd", error.Message,
                StringComparison.OrdinalIgnoreCase);

            OfficeShape selfIntersecting = OfficeShape.Polygon(
                new OfficePoint(0, 0),
                new OfficePoint(100, 100),
                new OfficePoint(0, 100),
                new OfficePoint(100, 0));
            selfIntersecting.FillColor = OfficeColor.FromRgb(14, 165, 233);
            Assert.Throws<NotSupportedException>(() =>
                slide.AddCustomGeometryPoints(selfIntersecting,
                    10, 10, 100, 100));

            geometry.FillRule = OfficeFillRule.NonZero;
            Assert.NotNull(slide.AddCustomGeometryPoints(geometry,
                10, 10, 100, 100));
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void PowerPointSlide_BoundsEvenOddIntersectionValidationBeforeMutation() {
            OfficePoint[] points = Enumerable.Range(0, 2049)
                .Select(index => {
                    double angle = index * Math.PI * 2D / 2049D;
                    return new OfficePoint(Math.Cos(angle) * 100D,
                        Math.Sin(angle) * 100D);
                })
                .ToArray();
            OfficeShape geometry = OfficeShape.Polygon(points);
            geometry.FillRule = OfficeFillRule.EvenOdd;
            geometry.FillColor = OfficeColor.FromRgb(14, 165, 233);
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();

            NotSupportedException error = Assert.Throws<NotSupportedException>(() =>
                slide.AddCustomGeometryPoints(geometry, 10, 10, 100, 100));
            Assert.Contains("even-odd", error.Message,
                StringComparison.OrdinalIgnoreCase);
            Assert.Empty(slide.Shapes);
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
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                PowerPointAutoShape shape = presentation.AddSlide().AddCustomGeometryPoints(
                    polygon, 20, 20, 120, 80);
                Assert.Equal(75, shape.OutlineTransparency);
                A.Alpha alpha = shape.Element.Descendants<A.Outline>().Single()
                    .Descendants<A.Alpha>().Single();
                Assert.InRange(alpha.Val!.Value, 25097, 25099);
                presentation.Save();
            }

            stream.Position = 0;
            using PowerPointDocument reopened = PowerPointDocument.Load(stream);
            PowerPointAutoShape authored = Assert.IsType<PowerPointAutoShape>(
                Assert.Single(reopened.Slides[0].Shapes));
            Assert.Equal(75, authored.OutlineTransparency);
        }

        [Fact]
        public void PowerPointSlide_AuthorsCompleteCustomGeometryStrokeStyle() {
            OfficeShape path = OfficeShape.Path(
                OfficePathCommand.MoveTo(0, 100),
                OfficePathCommand.LineTo(50, 0),
                OfficePathCommand.LineTo(100, 100));
            path.StrokeColor = OfficeColor.FromRgb(37, 99, 235);
            path.StrokeWidth = 2D;
            path.StrokeDashStyle = OfficeStrokeDashStyle.DashDot;
            path.StrokeLineCap = OfficeStrokeLineCap.Round;
            path.StrokeLineJoin = OfficeStrokeLineJoin.Bevel;
            path.StrokeStartMarker = new OfficeLineMarker(
                OfficeLineMarkerKind.Triangle, 9D, 12D);
            path.StrokeEndMarker = new OfficeLineMarker(
                OfficeLineMarkerKind.Diamond, 12D, 16D);

            using var stream = new System.IO.MemoryStream();
            using (PowerPointDocument presentation =
                   PowerPointDocument.Create(stream)) {
                presentation.AddSlide().AddCustomGeometryPoints(path,
                    20, 20, 120, 80);
                var validation = presentation.ValidateDocument();
                Assert.True(validation.Count == 0, string.Join(
                    Environment.NewLine, validation.Select(error =>
                        error.Description + Environment.NewLine
                        + error.NodeName)));
                presentation.Save();
            }

            stream.Position = 0;
            using PowerPointDocument reopened =
                PowerPointDocument.Load(stream);
            PowerPointAutoShape authored = Assert.IsType<PowerPointAutoShape>(
                Assert.Single(reopened.Slides[0].Shapes));
            A.Outline outline = authored.Element.Descendants<A.Outline>()
                .Single();
            Assert.Equal(A.PresetLineDashValues.DashDot,
                outline.GetFirstChild<A.PresetDash>()!.Val!.Value);
            Assert.Equal(A.LineCapValues.Round, outline.CapType!.Value);
            Assert.NotNull(outline.GetFirstChild<A.LineJoinBevel>());
            A.HeadEnd head = outline.GetFirstChild<A.HeadEnd>()!;
            Assert.Equal(A.LineEndValues.Triangle, head.Type!.Value);
            Assert.Equal(A.LineEndWidthValues.Medium, head.Width!.Value);
            Assert.Equal(A.LineEndLengthValues.Medium, head.Length!.Value);
            A.TailEnd tail = outline.GetFirstChild<A.TailEnd>()!;
            Assert.Equal(A.LineEndValues.Diamond, tail.Type!.Value);
            Assert.Equal(A.LineEndWidthValues.Large, tail.Width!.Value);
            Assert.Equal(A.LineEndLengthValues.Large, tail.Length!.Value);
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
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                PowerPointAutoShape shape = presentation.AddSlide().AddCustomGeometryPoints(
                    polygon, 20, 20, 120, 80);
                Assert.Equal(75, shape.FillTransparency);
                A.Alpha alpha = shape.Element.Descendants<A.SolidFill>().First()
                    .Descendants<A.Alpha>().Single();
                Assert.InRange(alpha.Val!.Value, 25097, 25099);
                presentation.Save();
            }

            stream.Position = 0;
            using PowerPointDocument reopened = PowerPointDocument.Load(stream);
            PowerPointAutoShape authored = Assert.IsType<PowerPointAutoShape>(
                Assert.Single(reopened.Slides[0].Shapes));
            Assert.Equal(75, authored.FillTransparency);
        }

        [Fact]
        public void PowerPointShape_OutlineOpacityPreservesSchemeColorChoice() {
            using var stream = new System.IO.MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            PowerPointAutoShape shape = presentation.AddSlide().AddRectangle(
                OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(120), OfficeIMO.OfficeLength.Points(80));
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            presentation.SetThemeColorForAllMasters(
                PowerPointThemeColor.Accent2, "123456");
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(120), OfficeIMO.OfficeLength.Points(80));
            Shape openXmlShape = Assert.IsType<Shape>(shape.Element);
            A.Outline themeOutline = slide.SlidePart.SlideLayoutPart!
                .SlideMasterPart!.ThemePart!.Theme.ThemeElements!
                .FormatScheme!.LineStyleList!.Elements<A.Outline>().First();
            themeOutline.Width = 38100;
            themeOutline.CapType = A.LineCapValues.Round;
            themeOutline.RemoveAllChildren<A.PresetDash>();
            themeOutline.RemoveAllChildren<A.Round>();
            themeOutline.RemoveAllChildren<A.LineJoinBevel>();
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
            OfficeImageExportResult svg = slide.ExportImage(
                OfficeImageExportFormat.Svg);
            // The PDF-backed SVG may flatten strokes into paths. Preserve authoring style above,
            // then verify the effective theme stroke color and transparency in the exported pixels.
            OfficeRasterImage image = ReadSvgOutput(svg);
            Assert.True(CountPixelsNear(image, OfficeColor.FromRgb(113, 133, 154)) > 20);
        }

        [Fact]
        public void PowerPointShape_FillTransparencyPreservesSchemeColorChoice() {
            using var stream = new System.IO.MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            PowerPointAutoShape shape = presentation.AddSlide().AddRectangle(
                OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(120), OfficeIMO.OfficeLength.Points(80));
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
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            presentation.SetThemeColorForAllMasters(
                PowerPointThemeColor.Accent2, "123456");
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(120), OfficeIMO.OfficeLength.Points(80));
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
        }

        [Fact]
        public void PowerPointShape_FillTransparencyMaterializesFixedThemeFill() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(120), OfficeIMO.OfficeLength.Points(80));
            Shape openXmlShape = Assert.IsType<Shape>(shape.Element);
            openXmlShape.ShapeProperties!.RemoveAllChildren<A.SolidFill>();
            A.FillStyleList fillStyles = slide.SlidePart.SlideLayoutPart!
                .SlideMasterPart!.ThemePart!.Theme!.ThemeElements!
                .FormatScheme!.FillStyleList!;
            fillStyles.ReplaceChild(new A.SolidFill(
                    new A.RgbColorModelHex { Val = "654321" }),
                fillStyles.ChildElements[0]);
            var referenceColor = new A.SchemeColor {
                Val = A.SchemeColorValues.Accent2
            };
            openXmlShape.ShapeStyle = new ShapeStyle(
                new A.LineReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent1
                }) { Index = 1U },
                new A.FillReference(referenceColor) { Index = 1U },
                new A.EffectReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent1
                }) { Index = 0U },
                new A.FontReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Dark1
                }) { Index = A.FontCollectionIndexValues.Minor });

            shape.FillTransparency = 40;

            A.SolidFill local = openXmlShape.ShapeProperties!
                .GetFirstChild<A.SolidFill>()!;
            A.RgbColorModelHex localColor = local
                .GetFirstChild<A.RgbColorModelHex>()!;
            Assert.Equal("654321", localColor.Val!.Value);
            Assert.Equal(60000, localColor.GetFirstChild<A.Alpha>()!.Val!.Value);
            Assert.Null(referenceColor.GetFirstChild<A.Alpha>());
            Assert.Equal(40, shape.FillTransparency);
        }

        [Fact]
        public void PowerPointShape_FillTransparencyRejectsTransformedPlaceholderThemeFill() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            presentation.SetThemeColorForAllMasters(
                PowerPointThemeColor.Accent2, "123456");
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(120), OfficeIMO.OfficeLength.Points(80));
            Shape openXmlShape = Assert.IsType<Shape>(shape.Element);
            openXmlShape.ShapeProperties!.RemoveAllChildren<A.SolidFill>();
            A.FillStyleList fillStyles = slide.SlidePart.SlideLayoutPart!
                .SlideMasterPart!.ThemePart!.Theme!.ThemeElements!
                .FormatScheme!.FillStyleList!;
            fillStyles.ReplaceChild(new A.SolidFill(
                    new A.SchemeColor(new A.Alpha { Val = 50000 }) {
                        Val = A.SchemeColorValues.PhColor
                    }),
                fillStyles.ChildElements[0]);
            var referenceColor = new A.SchemeColor {
                Val = A.SchemeColorValues.Accent2
            };
            openXmlShape.ShapeStyle = new ShapeStyle(
                new A.LineReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent1
                }) { Index = 1U },
                new A.FillReference(referenceColor) { Index = 1U },
                new A.EffectReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent1
                }) { Index = 0U },
                new A.FontReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Dark1
                }) { Index = A.FontCollectionIndexValues.Minor });

            string propertiesBefore = openXmlShape.ShapeProperties.OuterXml;
            string styleBefore = openXmlShape.ShapeStyle.OuterXml;

            Assert.Throws<NotSupportedException>(() =>
                shape.FillTransparency = 40);

            Assert.Equal(propertiesBefore,
                openXmlShape.ShapeProperties.OuterXml);
            Assert.Equal(styleBefore, openXmlShape.ShapeStyle.OuterXml);
        }

        [Fact]
        public void PowerPointShape_FillTransparencyRejectsInheritedPictureFillWithoutMutation() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(120), OfficeIMO.OfficeLength.Points(80));
            Shape openXmlShape = Assert.IsType<Shape>(shape.Element);
            openXmlShape.ShapeProperties!.RemoveAllChildren<A.SolidFill>();
            A.FillStyleList fillStyles = slide.SlidePart.SlideLayoutPart!
                .SlideMasterPart!.ThemePart!.Theme!.ThemeElements!
                .FormatScheme!.FillStyleList!;
            fillStyles.ReplaceChild(new A.BlipFill(new A.Blip()),
                fillStyles.ChildElements[0]);
            openXmlShape.ShapeStyle = new ShapeStyle(
                new A.LineReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent1
                }) { Index = 1U },
                new A.FillReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent2
                }) { Index = 1U },
                new A.EffectReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent1
                }) { Index = 0U },
                new A.FontReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Dark1
                }) { Index = A.FontCollectionIndexValues.Minor });
            string propertiesBefore = openXmlShape.ShapeProperties.OuterXml;
            string styleBefore = openXmlShape.ShapeStyle.OuterXml;

            Assert.Throws<NotSupportedException>(() =>
                shape.FillTransparency = 40);

            Assert.Equal(propertiesBefore, openXmlShape.ShapeProperties.OuterXml);
            Assert.Equal(styleBefore, openXmlShape.ShapeStyle.OuterXml);
        }

        [Fact]
        public void PowerPointShape_OutlineTransparencyMaterializesFixedThemeLine() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(120), OfficeIMO.OfficeLength.Points(80));
            Shape openXmlShape = Assert.IsType<Shape>(shape.Element);
            openXmlShape.ShapeProperties!.RemoveAllChildren<A.Outline>();
            A.LineStyleList lineStyles = slide.SlidePart.SlideLayoutPart!
                .SlideMasterPart!.ThemePart!.Theme!.ThemeElements!
                .FormatScheme!.LineStyleList!;
            A.Outline fixedThemeLine = new A.Outline(
                new A.SolidFill(
                    new A.RgbColorModelHex { Val = "654321" }),
                new A.PresetDash { Val = A.PresetLineDashValues.Dash }) {
                Width = 38100
            };
            lineStyles.ReplaceChild(fixedThemeLine,
                lineStyles.ChildElements[0]);
            var referenceColor = new A.SchemeColor {
                Val = A.SchemeColorValues.Accent2
            };
            openXmlShape.ShapeStyle = new ShapeStyle(
                new A.LineReference(referenceColor) { Index = 1U },
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

            A.Outline local = openXmlShape.ShapeProperties!
                .GetFirstChild<A.Outline>()!;
            A.RgbColorModelHex localColor = local
                .GetFirstChild<A.SolidFill>()!
                .GetFirstChild<A.RgbColorModelHex>()!;
            Assert.Equal("654321", localColor.Val!.Value);
            Assert.Equal(60000, localColor.GetFirstChild<A.Alpha>()!.Val!.Value);
            Assert.Equal(38100, local.Width!.Value);
            Assert.Equal(A.PresetLineDashValues.Dash,
                local.GetFirstChild<A.PresetDash>()!.Val!.Value);
            Assert.Null(referenceColor.GetFirstChild<A.Alpha>());
            Assert.Equal(40, shape.OutlineTransparency);
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void PowerPointShape_OutlineTransparencyRejectsTransformedPlaceholderThemeLine() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(120), OfficeIMO.OfficeLength.Points(80));
            Shape openXmlShape = Assert.IsType<Shape>(shape.Element);
            openXmlShape.ShapeProperties!.RemoveAllChildren<A.Outline>();
            A.LineStyleList lineStyles = slide.SlidePart.SlideLayoutPart!
                .SlideMasterPart!.ThemePart!.Theme!.ThemeElements!
                .FormatScheme!.LineStyleList!;
            lineStyles.ReplaceChild(new A.Outline(
                    new A.SolidFill(new A.SchemeColor(
                        new A.AlphaModulation { Val = 50000 }) {
                        Val = A.SchemeColorValues.PhColor
                    })),
                lineStyles.ChildElements[0]);
            openXmlShape.ShapeStyle = new ShapeStyle(
                new A.LineReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent2
                }) { Index = 1U },
                new A.FillReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent1
                }) { Index = 1U },
                new A.EffectReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent1
                }) { Index = 0U },
                new A.FontReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Dark1
                }) { Index = A.FontCollectionIndexValues.Minor });
            string propertiesBefore = openXmlShape.ShapeProperties.OuterXml;
            string styleBefore = openXmlShape.ShapeStyle.OuterXml;

            Assert.Throws<NotSupportedException>(() =>
                shape.OutlineTransparency = 40);

            Assert.Equal(propertiesBefore,
                openXmlShape.ShapeProperties.OuterXml);
            Assert.Equal(styleBefore, openXmlShape.ShapeStyle.OuterXml);
        }

        [Fact]
        public void PowerPointShape_TransparencyMaterializesMixedThemeColors() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddRectangle(
                OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(120), OfficeIMO.OfficeLength.Points(80));
            Shape openXmlShape = Assert.IsType<Shape>(shape.Element);
            openXmlShape.ShapeProperties!.RemoveAllChildren<A.SolidFill>();
            openXmlShape.ShapeProperties.RemoveAllChildren<A.Outline>();
            var mixedGradient = new A.GradientFill(
                new A.GradientStopList(
                    new A.GradientStop(new A.SchemeColor {
                        Val = A.SchemeColorValues.PhColor
                    }) { Position = 0 },
                    new A.GradientStop(new A.RgbColorModelHex {
                        Val = "654321"
                    }) { Position = 100000 }),
                new A.LinearGradientFill { Angle = 0, Scaled = true });
            A.FormatScheme scheme = slide.SlidePart.SlideLayoutPart!
                .SlideMasterPart!.ThemePart!.Theme!.ThemeElements!
                .FormatScheme!;
            scheme.FillStyleList!.ReplaceChild(mixedGradient.CloneNode(true),
                scheme.FillStyleList.ChildElements[0]);
            var mixedOutline = new A.Outline(mixedGradient.CloneNode(true)) {
                Width = 38100
            };
            scheme.LineStyleList!.ReplaceChild(mixedOutline,
                scheme.LineStyleList.ChildElements[0]);
            openXmlShape.ShapeStyle = new ShapeStyle(
                new A.LineReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent1
                }) { Index = 1U },
                new A.FillReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent2
                }) { Index = 1U },
                new A.EffectReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent1
                }) { Index = 0U },
                new A.FontReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Dark1
                }) { Index = A.FontCollectionIndexValues.Minor });

            shape.FillTransparency = 40;
            shape.OutlineTransparency = 40;

            A.GradientFill localFill = openXmlShape.ShapeProperties
                .GetFirstChild<A.GradientFill>()!;
            A.GradientFill localOutlineFill = openXmlShape.ShapeProperties
                .GetFirstChild<A.Outline>()!
                .GetFirstChild<A.GradientFill>()!;
            Assert.All(localFill.Descendants<A.GradientStop>(), stop =>
                Assert.Equal(60000, stop.Descendants<A.Alpha>()
                    .Single().Val!.Value));
            Assert.All(localOutlineFill.Descendants<A.GradientStop>(), stop =>
                Assert.Equal(60000, stop.Descendants<A.Alpha>()
                    .Single().Val!.Value));
            Assert.Empty(presentation.ValidateDocument());
        }

        [Fact]
        public void PowerPointSlide_RejectsNonFiniteCustomGeometryOpacityBeforeMutation() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            OfficeShape invalidFill = OfficeShape.Polygon(
                new OfficePoint(0, 0), new OfficePoint(100, 0),
                new OfficePoint(50, 100));
            invalidFill.FillColor = OfficeColor.FromRgb(14, 165, 233);
            invalidFill.FillOpacity = double.NaN;
            OfficeShape invalidStroke = OfficeShape.Path(
                OfficePathCommand.MoveTo(0, 0),
                OfficePathCommand.LineTo(100, 100));
            invalidStroke.StrokeColor = OfficeColor.FromRgb(37, 99, 235);
            invalidStroke.StrokeOpacity = double.PositiveInfinity;

            Assert.Throws<ArgumentException>(() => slide
                .AddCustomGeometryPoints(invalidFill, 10, 10, 100, 80));
            Assert.Throws<ArgumentException>(() => slide
                .AddCustomGeometryPoints(invalidStroke, 10, 10, 100, 80));
            Assert.Empty(slide.Shapes);
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
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape shape = slide.AddCustomGeometryCm(
                polygon, 1, 1, 4, 3, "Shared polygon");
            Assert.Null(shape.ShapeType);
            Assert.Throws<ArgumentException>(() => slide.AddCustomGeometryPoints(
                OfficeShape.Rectangle(10, 10), 0, 0, 10, 10));
        }

        [Fact]
        public void PowerPointSlide_RejectsUnrepresentableCustomGeometryCoordinatesBeforeMutation() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            int originalShapeCount = slide.Shapes.Count;
            OfficeShape infiniteNormalized = OfficeShape.Path(double.Epsilon,
                100D, OfficePathCommand.MoveTo(0D, 0D),
                OfficePathCommand.LineTo(1D, 1D));
            OfficeShape outOfRangeNormalized = OfficeShape.Path(1D, 100D,
                OfficePathCommand.MoveTo(0D, 0D),
                OfficePathCommand.LineTo(300000000D, 1D));

            Assert.Throws<ArgumentException>(() =>
                slide.AddCustomGeometryPoints(infiniteNormalized,
                    0D, 0D, 100D, 100D));
            Assert.Throws<ArgumentException>(() =>
                slide.AddCustomGeometryPoints(outOfRangeNormalized,
                    0D, 0D, 100D, 100D));
            Assert.Equal(originalShapeCount, slide.Shapes.Count);
            Assert.Empty(presentation.ValidateDocument());
        }

    }
}
