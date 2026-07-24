using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        [Fact]
        public void PowerPointSlide_ExportsSolidBackgroundToPngAndSvgThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 160);
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "112233";

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions { Scale = 2D, IncludeSlideContent = false });
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg, new PowerPointImageExportOptions { IncludeSlideContent = false, Scale = 2D });

            Assert.Equal(OfficeImageExportFormat.Png, png.Format);
            Assert.Equal(480, png.Width);
            Assert.Equal(320, png.Height);
            Assert.Equal(480, svg.Width);
            Assert.Equal(320, svg.Height);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(480, image!.Width);
            Assert.Equal(320, image.Height);
            Assert.Equal(OfficeColor.FromRgb(17, 34, 51), image.GetPixel(8, 8));
            AssertNoUnexpectedDiagnostics(png.Diagnostics);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("width=\"480px\"", svgText, StringComparison.Ordinal);
            Assert.Contains("#112233", svgText, StringComparison.OrdinalIgnoreCase);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
        }

        [Fact]
        public void PowerPointSlide_ToImageFluentSavesSvgThroughSharedBuilder() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 160);
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "112233";
            slide.AddTextBoxPoints("Hidden by fluent content toggle", 20, 20, 140, 24);

            using var output = new MemoryStream();
            OfficeImageExportResult result = slide.ToImage()
                .As(OfficeImageExportFormat.Svg)
                .WithoutContent()
                .Save(output);

            Assert.Equal(OfficeImageExportFormat.Svg, result.Format);
            Assert.Equal(result.Bytes.Length, output.Length);
            string svgText = Encoding.UTF8.GetString(output.ToArray());
            Assert.Contains("#112233", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("Hidden by fluent content toggle", svgText, StringComparison.Ordinal);
            AssertNoUnexpectedDiagnostics(result.Diagnostics);
        }

        [Fact]
        public void PowerPointSlide_ToImageUsesConfiguredFormatForBytes() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 90);
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "112233";

            byte[] png = slide.ToImage()
                .WithoutContent()
                .AsPng()
                .ToBytes();
            string svg = Encoding.UTF8.GetString(slide.ToImage()
                .WithoutContent()
                .AsSvg()
                .ToBytes());

            Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, png.Take(4).ToArray());
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("#112233", svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointPresentation_ToImagesFluentSavesVisibleSlidesAndCanIncludeHiddenSlides() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 100);

            PowerPointSlide visible = presentation.AddSlide();
            visible.BackgroundColor = "112233";

            PowerPointSlide hidden = presentation.AddSlide();
            hidden.Hidden = true;
            hidden.BackgroundColor = "445566";

            string folder = Path.Combine(Path.GetTempPath(), "officeimo-ppt-images-" + Guid.NewGuid().ToString("N"));
            try {
                IReadOnlyList<OfficeImageExportResult> visibleResults = presentation.ToImages()
                    .AsSvg()
                    .Save(folder);

                Assert.Single(visibleResults);
                Assert.Equal("Slide 1", visibleResults[0].Name);
                string visiblePath = Path.Combine(folder, "Slide 1.svg");
                Assert.True(File.Exists(visiblePath));
                string visibleSvg = File.ReadAllText(visiblePath);
                Assert.Contains("#112233", visibleSvg, StringComparison.OrdinalIgnoreCase);
                Assert.DoesNotContain("#445566", visibleSvg, StringComparison.OrdinalIgnoreCase);
                Assert.False(File.Exists(Path.Combine(folder, "Slide 2.svg")));

                string allFolder = Path.Combine(folder, "all");
                IReadOnlyList<OfficeImageExportResult> allResults = presentation.ToImages()
                    .IncludeHiddenSlides()
                    .AsSvg()
                    .Save(allFolder);

                Assert.Equal(2, allResults.Count);
                Assert.Equal("Slide 2", allResults[1].Name);
                string hiddenPath = Path.Combine(allFolder, "Slide 2.svg");
                Assert.True(File.Exists(hiddenPath));
                Assert.Contains("#445566", File.ReadAllText(hiddenPath), StringComparison.OrdinalIgnoreCase);
            } finally {
                if (Directory.Exists(folder)) {
                    Directory.Delete(folder, recursive: true);
                }
            }
        }

        [Fact]
        public void PowerPointPresentation_ToImagesFluentSelectsSlidesAndRangesThroughSharedBatchBuilder() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 100);

            PowerPointSlide first = presentation.AddSlide();
            first.BackgroundColor = "112233";
            first.AddTextBoxPoints("First selected", 20, 20, 120, 24);

            PowerPointSlide hidden = presentation.AddSlide();
            hidden.Hidden = true;
            hidden.BackgroundColor = "445566";
            hidden.AddTextBoxPoints("Hidden selected", 20, 20, 120, 24);

            PowerPointSlide third = presentation.AddSlide();
            third.BackgroundColor = "778899";
            third.AddTextBoxPoints("Third selected", 20, 20, 120, 24);

            string folder = Path.Combine(Path.GetTempPath(), "officeimo-ppt-filtered-images-" + Guid.NewGuid().ToString("N"));
            try {
                IReadOnlyList<OfficeImageExportResult> selectedResults = presentation.ToImages()
                    .ForSlides(3, 1, 3)
                    .AsSvg()
                    .Save(folder);

                Assert.Equal(2, selectedResults.Count);
                Assert.Equal("Slide 1", selectedResults[0].Name);
                Assert.Equal("Slide 3", selectedResults[1].Name);
                Assert.True(File.Exists(Path.Combine(folder, "Slide 1.svg")));
                Assert.False(File.Exists(Path.Combine(folder, "Slide 2.svg")));
                Assert.True(File.Exists(Path.Combine(folder, "Slide 3.svg")));
                Assert.Contains("#112233", File.ReadAllText(Path.Combine(folder, "Slide 1.svg")), StringComparison.OrdinalIgnoreCase);
                Assert.Contains("#778899", File.ReadAllText(Path.Combine(folder, "Slide 3.svg")), StringComparison.OrdinalIgnoreCase);

                string rangeFolder = Path.Combine(folder, "range");
                IReadOnlyList<OfficeImageExportResult> rangeResults = presentation.ToImages()
                    .ForSlideRange(2, 3)
                    .IncludeHiddenSlides()
                    .AsSvg()
                    .Save(rangeFolder);

                Assert.Equal(2, rangeResults.Count);
                Assert.Equal("Slide 2", rangeResults[0].Name);
                Assert.Equal("Slide 3", rangeResults[1].Name);
                Assert.Contains("#445566", File.ReadAllText(Path.Combine(rangeFolder, "Slide 2.svg")), StringComparison.OrdinalIgnoreCase);
            } finally {
                if (Directory.Exists(folder)) {
                    Directory.Delete(folder, recursive: true);
                }
            }
        }

        [Fact]
        public void PowerPointSlide_ExportsGradientBackgroundThroughDrawingGradient() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 80);
            PowerPointSlide slide = presentation.AddSlide();
            slide.SetBackgroundGradient("112233", "445566", 45D);

            OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Svg, new PowerPointImageExportOptions { IncludeSlideContent = false });
            string svgText = Encoding.UTF8.GetString(result.Bytes);

            Assert.Contains("<linearGradient", svgText, StringComparison.Ordinal);
            Assert.Contains("#112233", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#445566", svgText, StringComparison.OrdinalIgnoreCase);
            AssertNoUnexpectedDiagnostics(result.Diagnostics);
        }

        [Fact]
        public void PowerPointSlide_ExportsShapeGradientThroughDrawingGradient() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 80);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox textBox = slide.AddTextBoxPoints(string.Empty, 10, 10, 100, 60);
            Shape shape = Assert.IsType<Shape>(textBox.Element);
            shape.ShapeProperties!.Append(new A.GradientFill(
                new A.GradientStopList(
                    new A.GradientStop(new A.RgbColorModelHex { Val = "112233" }) {
                        Position = 0
                    },
                    new A.GradientStop(new A.RgbColorModelHex { Val = "445566" }) {
                        Position = 100000
                    }),
                new A.LinearGradientFill { Angle = 45 * 60000 }));

            OfficeImageExportResult result = slide.ExportImage(
                OfficeImageExportFormat.Svg,
                new PowerPointImageExportOptions { IncludeSlideBackground = false });
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(
                new PowerPointImageExportOptions { IncludeSlideBackground = false });
            string svgText = Encoding.UTF8.GetString(result.Bytes);

            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>(), item => item.Shape.FillGradient != null);
            Assert.Equal("#112233", rendered.Shape.FillGradient!.Stops[0].Color.ToString());
            Assert.Equal("#445566", rendered.Shape.FillGradient.Stops[1].Color.ToString());
            Assert.Contains("<linearGradient", svgText, StringComparison.Ordinal);
            Assert.Contains("#112233", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#445566", svgText, StringComparison.OrdinalIgnoreCase);
            AssertNoUnexpectedDiagnostics(result.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void PowerPointSlide_ResolvesThemeOverrideShapeGradientWithPlaceholderColor(
            bool useSlideOverride) {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 80);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape source = slide.AddShapePoints(A.ShapeTypeValues.Rectangle,
                10, 10, 100, 60);
            Shape shape = Assert.IsType<Shape>(source.Element);
            A.ThemeElements masterElements = slide.SlidePart.SlideLayoutPart!
                .SlideMasterPart!.ThemePart!.Theme!.ThemeElements!;
            A.ColorScheme colors = (A.ColorScheme)masterElements.ColorScheme!
                .CloneNode(true);
            A.Accent1Color accent1 = colors.GetFirstChild<A.Accent1Color>()!;
            accent1.RemoveAllChildren();
            accent1.Append(new A.RgbColorModelHex { Val = "112233" });
            A.FormatScheme format = (A.FormatScheme)masterElements.FormatScheme!
                .CloneNode(true);
            A.FillStyleList fillStyles = format.GetFirstChild<A.FillStyleList>()!;
            fillStyles.RemoveAllChildren();
            fillStyles.Append(new A.GradientFill(
                new A.GradientStopList(
                    new A.GradientStop(new A.SchemeColor {
                        Val = A.SchemeColorValues.PhColor
                    }) { Position = 0 },
                    new A.GradientStop(new A.RgbColorModelHex { Val = "445566" }) {
                        Position = 100000
                    }),
                new A.LinearGradientFill { Angle = 0 }));
            DocumentFormat.OpenXml.Packaging.ThemeOverridePart overridePart =
                useSlideOverride
                ? slide.SlidePart.AddNewPart<DocumentFormat.OpenXml.Packaging.ThemeOverridePart>()
                : slide.SlidePart.SlideLayoutPart!
                    .AddNewPart<DocumentFormat.OpenXml.Packaging.ThemeOverridePart>();
            overridePart.ThemeOverride = new A.ThemeOverride(
                colors,
                masterElements.FontScheme!.CloneNode(true),
                format);
            shape.ShapeStyle = new ShapeStyle(
                new A.LineReference(new A.SchemeColor {
                    Val = A.SchemeColorValues.Accent1
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

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(
                new PowerPointImageExportOptions { IncludeSlideBackground = false });

            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>(), item => item.Shape.FillGradient != null);
            Assert.Equal("#112233", rendered.Shape.FillGradient!.Stops[0].Color.ToString());
            Assert.Equal("#445566", rendered.Shape.FillGradient.Stops[1].Color.ToString());
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
        }

        [Fact]
        public void PowerPointSlide_BoundsOutOfRangeThemeFillIndex() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape source = slide.AddShapePoints(
                A.ShapeTypeValues.Rectangle, 10, 10, 100, 60);
            Shape shape = Assert.IsType<Shape>(source.Element);
            shape.ShapeProperties!.RemoveAllChildren<A.SolidFill>();
            shape.ShapeStyle = new ShapeStyle(
                new A.LineReference { Index = 1U },
                new A.FillReference { Index = uint.MaxValue },
                new A.EffectReference { Index = 0U },
                new A.FontReference {
                    Index = A.FontCollectionIndexValues.Minor
                });

            PowerPointSlideVisualSnapshot snapshot =
                slide.CreateVisualSnapshot(new PowerPointImageExportOptions {
                    IncludeSlideBackground = false
                });

            Assert.NotNull(snapshot.Drawing);
        }

        [Fact]
        public void PowerPointSlide_ReportsUnsupportedPathGradientWithoutApproximatingIt() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 80);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape circle = slide.AddShapePoints(A.ShapeTypeValues.Rectangle,
                10, 10, 100, 60);
            PowerPointAutoShape shapePath = slide.AddShapePoints(A.ShapeTypeValues.Rectangle,
                130, 10, 100, 60);
            AddShapePathGradient(Assert.IsType<Shape>(circle.Element),
                A.PathShadeValues.Circle);
            AddShapePathGradient(Assert.IsType<Shape>(shapePath.Element),
                A.PathShadeValues.Shape);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(
                new PowerPointImageExportOptions { IncludeSlideBackground = false });

            OfficeDrawingShape[] rendered = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .ToArray();
            Assert.Single(rendered, item => item.Shape.FillRadialGradient != null);
            Assert.Single(snapshot.Diagnostics, diagnostic =>
                diagnostic.Code == "unsupported-powerpoint-shape"
                && diagnostic.Message.Contains("gradient", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void PowerPointSlide_ReportsUnsupportedTextBearingPresetFrame() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 90);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox textBox = slide.AddTextBoxPoints("Keep this text", 20, 20,
                120, 50);
            Shape shape = Assert.IsType<Shape>(textBox.Element);
            shape.ShapeProperties!.GetFirstChild<A.PresetGeometry>()!.Preset =
                A.ShapeTypeValues.Funnel;

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(
                new PowerPointImageExportOptions { IncludeSlideBackground = false });

            Assert.Contains(snapshot.Diagnostics, diagnostic =>
                diagnostic.Code == "unsupported-powerpoint-shape"
                && diagnostic.Message.Contains("frame geometry",
                    StringComparison.OrdinalIgnoreCase));
            Assert.Contains(snapshot.Drawing.Elements,
                element => element is OfficeDrawingText or OfficeDrawingRichText);
        }

        [Fact]
        public void PowerPointSlide_HonorsGradientRotateWithShapeInSnapshotAndRaster() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 120);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape rotating = slide.AddShapePoints(A.ShapeTypeValues.Rectangle,
                20, 30, 80, 40);
            PowerPointAutoShape fixedToSlide = slide.AddShapePoints(A.ShapeTypeValues.Rectangle,
                140, 30, 80, 40);
            rotating.Rotation = 90D;
            fixedToSlide.Rotation = 90D;
            AddShapeLinearGradient(Assert.IsType<Shape>(rotating.Element),
                rotateWithShape: true);
            AddShapeLinearGradient(Assert.IsType<Shape>(fixedToSlide.Element),
                rotateWithShape: false);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(
                new PowerPointImageExportOptions { IncludeSlideBackground = false });
            OfficeDrawingShape[] rendered = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(item => item.Shape.FillGradient != null)
                .ToArray();

            Assert.Equal(2, rendered.Length);
            Assert.Equal(rendered[0].Shape.FillGradient!.StartY,
                rendered[0].Shape.FillGradient.EndY, 6);
            Assert.Equal(rendered[1].Shape.FillGradient!.StartX,
                rendered[1].Shape.FillGradient.EndX, 6);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png,
                new PowerPointImageExportOptions { IncludeSlideBackground = false });
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(image!.GetPixel(60, 18).R > image.GetPixel(60, 82).R);
            Assert.True(image.GetPixel(164, 50).R > image.GetPixel(196, 50).R);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
        }

        [Fact]
        public void PowerPointSlide_PreservesSlideFixedGradientAngleOnRotatedNonSquareShape() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 100);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape source = slide.AddShapePoints(A.ShapeTypeValues.Rectangle,
                30, 25, 100, 40);
            source.Rotation = 31D;
            AddShapeLinearGradient(Assert.IsType<Shape>(source.Element),
                rotateWithShape: false, angleDegrees: 37D);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(
                new PowerPointImageExportOptions { IncludeSlideBackground = false });
            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>(), item => item.Shape.FillGradient != null);
            OfficeLinearGradient gradient = rendered.Shape.FillGradient!;
            OfficeTransform transform = rendered.Shape.Transform!.Value;
            OfficePoint start = transform.TransformPoint(new OfficePoint(
                gradient.StartX * rendered.Shape.Width,
                gradient.StartY * rendered.Shape.Height));
            OfficePoint end = transform.TransformPoint(new OfficePoint(
                gradient.EndX * rendered.Shape.Width,
                gradient.EndY * rendered.Shape.Height));
            double actual = Math.Atan2(end.Y - start.Y, end.X - start.X)
                * 180D / Math.PI;
            if (actual < 0D) actual += 360D;

            Assert.InRange(actual, 36.999D, 37.001D);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
        }

        [Fact]
        public void PowerPointSlide_ReportsUnrepresentableRotatedSlideFixedRadialGradient() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 100);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape source = slide.AddShapePoints(A.ShapeTypeValues.Rectangle,
                30, 25, 100, 40);
            source.Rotation = 37D;
            AddShapePathGradient(Assert.IsType<Shape>(source.Element),
                A.PathShadeValues.Circle, rotateWithShape: false);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(
                new PowerPointImageExportOptions { IncludeSlideBackground = false });

            Assert.DoesNotContain(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(),
                item => item.Shape.FillRadialGradient != null);
            Assert.Contains(snapshot.Diagnostics, diagnostic =>
                diagnostic.Code == "unsupported-powerpoint-shape"
                && diagnostic.Message.Contains("gradient",
                    StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void PowerPointSlide_ReportsSlideFixedGradientInsideTransformedGroup() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 100);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape source = slide.AddShapePoints(A.ShapeTypeValues.Rectangle,
                30, 25, 80, 30);
            PowerPointAutoShape anchor = slide.AddShapePoints(A.ShapeTypeValues.Rectangle,
                120, 25, 10, 10);
            anchor.FillColor = "E5E7EB";
            AddShapeLinearGradient(Assert.IsType<Shape>(source.Element),
                rotateWithShape: false, angleDegrees: 37D);
            slide.GroupShapes(new PowerPointShape[] { source, anchor },
                "Rotated gradient group");
            GroupShape group = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<GroupShape>()
                .Single();
            group.GroupShapeProperties!.TransformGroup!.Rotation = 23 * 60000;

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(
                new PowerPointImageExportOptions { IncludeSlideBackground = false });

            Assert.Contains(snapshot.Diagnostics, diagnostic =>
                diagnostic.Code == "unsupported-powerpoint-shape"
                && diagnostic.Message.Contains("gradient",
                    StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void PowerPointSlide_SharedSnapshotStopsAtConfiguredGroupDepth() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape first = slide.AddShapePoints(A.ShapeTypeValues.Rectangle, 10, 10, 20, 20);
            PowerPointAutoShape second = slide.AddShapePoints(A.ShapeTypeValues.Rectangle, 40, 10, 20, 20);
            PowerPointAutoShape third = slide.AddShapePoints(A.ShapeTypeValues.Rectangle, 70, 10, 20, 20);
            PowerPointGroupShape inner = slide.GroupShapes(new PowerPointShape[] { first, second }, "Inner");
            slide.GroupShapes(new PowerPointShape[] { inner, third }, "Outer");

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(
                new PowerPointImageExportOptions { MaxGroupShapeDepth = 1 });

            Assert.Contains(snapshot.Diagnostics, diagnostic =>
                diagnostic.Message.Contains(nameof(PowerPointImageExportOptions.MaxGroupShapeDepth), StringComparison.Ordinal));
        }

        private static void AddShapeLinearGradient(Shape shape, bool rotateWithShape,
            double angleDegrees = 0D) {
            shape.ShapeProperties!.Append(new A.GradientFill(
                new A.GradientStopList(
                    new A.GradientStop(new A.RgbColorModelHex { Val = "FF0000" }) {
                        Position = 0
                    },
                    new A.GradientStop(new A.RgbColorModelHex { Val = "0000FF" }) {
                        Position = 100000
                    }),
                new A.LinearGradientFill {
                    Angle = checked((int)Math.Round(angleDegrees * 60000D,
                        MidpointRounding.AwayFromZero))
                }) {
                RotateWithShape = rotateWithShape
            });
        }

        private static void AddShapePathGradient(Shape shape, A.PathShadeValues path,
            bool? rotateWithShape = null) {
            var gradient = new A.GradientFill(
                new A.GradientStopList(
                    new A.GradientStop(new A.RgbColorModelHex { Val = "112233" }) {
                        Position = 0
                    },
                    new A.GradientStop(new A.RgbColorModelHex { Val = "445566" }) {
                        Position = 100000
                    }),
                new A.PathGradientFill { Path = path }) {
                RotateWithShape = rotateWithShape
            };
            shape.ShapeProperties!.Append(gradient);
        }

        [Fact]
        public void PowerPointSlide_ExportsInheritedThemeGradientBackgroundStyleThroughDrawingGradient() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 80);
            presentation.SetThemeColor(PowerPointThemeColor.Accent1, "112233");
            presentation.SetThemeColor(PowerPointThemeColor.Accent2, "445566");
            PowerPointSlide slide = presentation.AddSlide();
            DocumentFormat.OpenXml.Packaging.SlideMasterPart masterPart = slide.SlidePart.SlideLayoutPart!.SlideMasterPart!;
            A.BackgroundFillStyleList backgroundFills = masterPart.ThemePart!.Theme!.ThemeElements!.FormatScheme!
                .GetFirstChild<A.BackgroundFillStyleList>()!;
            backgroundFills.RemoveAllChildren();
            backgroundFills.Append(new A.GradientFill(
                new A.GradientStopList(
                    new A.GradientStop(new A.SchemeColor { Val = A.SchemeColorValues.Accent1 }) { Position = 0 },
                    new A.GradientStop(new A.SchemeColor { Val = A.SchemeColorValues.Accent2 }) { Position = 100000 }),
                new A.LinearGradientFill { Angle = 0 }));
            masterPart.SlideMaster.CommonSlideData!.Background = new Background(new BackgroundStyleReference { Index = 1001U });

            PowerPointSlideBackground background = slide.GetBackground();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg, new PowerPointImageExportOptions { IncludeSlideContent = false });
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions { IncludeSlideContent = false });

            Assert.Equal(PowerPointSlideBackgroundKind.LinearGradient, background.Kind);
            Assert.Equal("112233", background.GradientStartColor);
            Assert.Equal("445566", background.GradientEndColor);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(120, image!.Width);
            Assert.Equal(80, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<linearGradient", svgText, StringComparison.Ordinal);
            Assert.Contains("#112233", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#445566", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ExportsGradientBackgroundAlphaThroughSharedDrawingStops() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 80);
            PowerPointSlide slide = presentation.AddSlide();
            slide.SetBackgroundGradient("112233", "445566", 45D);

            A.GradientFill gradient = slide.SlidePart.Slide.CommonSlideData!.Background!.BackgroundProperties!
                .GetFirstChild<A.GradientFill>()!;
            A.GradientStop[] stops = gradient.GetFirstChild<A.GradientStopList>()!
                .Elements<A.GradientStop>()
                .OrderBy(stop => stop.Position?.Value ?? 0)
                .ToArray();
            stops[0].GetFirstChild<A.RgbColorModelHex>()!.Append(new A.Alpha { Val = 50000 });
            stops[1].GetFirstChild<A.RgbColorModelHex>()!.Append(new A.Alpha { Val = 25000 });

            PowerPointSlideBackground background = slide.GetBackground();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg, new PowerPointImageExportOptions { IncludeSlideContent = false });
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions { IncludeSlideContent = false });

            Assert.Equal("11223380", background.GradientStartColor);
            Assert.Equal("44556640", background.GradientEndColor);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("stop-color=\"#112233\"", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stop-opacity=\"0.502\"", svgText, StringComparison.Ordinal);
            Assert.Contains("stop-color=\"#445566\"", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stop-opacity=\"0.251\"", svgText, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(120, image!.Width);
            Assert.Equal(80, image.Height);
        }

        [Fact]
        public void PowerPointSlide_ExportsSystemColorGradientBackgroundThroughSharedDrawingStops() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 80);
            PowerPointSlide slide = presentation.AddSlide();
            slide.SetBackgroundGradient("112233", "445566", 45D);

            A.GradientFill gradient = slide.SlidePart.Slide.CommonSlideData!.Background!.BackgroundProperties!
                .GetFirstChild<A.GradientFill>()!;
            A.GradientStop[] stops = gradient.GetFirstChild<A.GradientStopList>()!
                .Elements<A.GradientStop>()
                .OrderBy(stop => stop.Position?.Value ?? 0)
                .ToArray();
            stops[0].RemoveAllChildren<A.RgbColorModelHex>();
            stops[0].Append(new A.SystemColor(new A.Alpha { Val = 50000 }) { Val = A.SystemColorValues.Window, LastColor = "336699" });

            PowerPointSlideBackground background = slide.GetBackground();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg, new PowerPointImageExportOptions { IncludeSlideContent = false });
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions { IncludeSlideContent = false });

            Assert.Equal("33669980", background.GradientStartColor);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("stop-color=\"#336699\"", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stop-opacity=\"0.502\"", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_RendersRectangleAndTextBoxThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape rectangle = slide.AddRectanglePoints(20, 20, 60, 30);
            rectangle.FillColor = "22AA66";
            rectangle.OutlineColor = "114433";
            rectangle.OutlineWidthPoints = 1D;

            PowerPointTextBox textBox = slide.AddTextBoxPoints("Shared renderer", 24, 54, 100, 24);
            textBox.FontSize = 12;
            textBox.Color = "111111";

            OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Png);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(result.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.FromRgb(34, 170, 102), image!.GetPixel(30, 30));
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText);
        }

        [Fact]
        public void PowerPointSlide_ProjectsThemeAutoShapeFillAndOutlineThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            presentation.SetThemeColor(PowerPointThemeColor.Accent1, "204060");
            presentation.SetThemeColor(PowerPointThemeColor.Accent2, "884422");
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape rectangle = slide.AddRectanglePoints(20, 20, 80, 40);
            rectangle.FillColor = "FFFFFF";
            rectangle.OutlineColor = "000000";
            rectangle.OutlineWidthPoints = 2D;
            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<Shape>()
                .Last();
            ShapeProperties properties = shape.ShapeProperties!;
            A.Outline outline = properties.GetFirstChild<A.Outline>()!;
            A.SolidFill themeFill = new A.SolidFill(
                new A.SchemeColor(new A.LuminanceModulation { Val = 50000 }) { Val = A.SchemeColorValues.Accent1 });
            A.SolidFill themeOutline = new A.SolidFill(
                new A.SchemeColor(new A.Shade { Val = 50000 }) { Val = A.SchemeColorValues.Accent2 });
            properties.RemoveAllChildren<A.SolidFill>();
            properties.InsertBefore(themeFill, outline);
            outline.RemoveAllChildren<A.SolidFill>();
            outline.InsertAt(themeOutline, 0);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), drawingShape =>
                Math.Abs(drawingShape.X - 20D) < 0.000001D &&
                Math.Abs(drawingShape.Y - 20D) < 0.000001D);
            Assert.Equal(OfficeColor.FromRgb(16, 32, 48), rendered.Shape.FillColor);
            Assert.Equal(OfficeColor.FromRgb(68, 34, 17), rendered.Shape.StrokeColor);
            Assert.Equal(2D, rendered.Shape.StrokeWidth);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("#102030", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#442211", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.FromRgb(16, 32, 48), image!.GetPixel(30, 30));
        }

        [Fact]
        public void PowerPointSlide_ProjectsSystemColorSolidFillFallbackThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape rectangle = slide.AddRectanglePoints(20, 20, 80, 40);
            rectangle.FillColor = "FFFFFF";
            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<Shape>()
                .Last();
            ShapeProperties properties = shape.ShapeProperties!;
            properties.RemoveAllChildren<A.SolidFill>();
            properties.InsertAt(new A.SolidFill(
                new A.SystemColor(new A.Alpha { Val = 50000 }) {
                    Val = A.SystemColorValues.WindowText,
                    LastColor = "336699"
                }), 0);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), drawingShape =>
                Math.Abs(drawingShape.X - 20D) < 0.000001D &&
                Math.Abs(drawingShape.Y - 20D) < 0.000001D);
            Assert.Equal(OfficeColor.FromRgba(51, 102, 153, 128), rendered.Shape.FillColor);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("fill=\"#336699\" fill-opacity=\"0.502\"", svgText, StringComparison.OrdinalIgnoreCase);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
        }

        [Fact]
        public void PowerPointSlide_ProjectsAutoShapeFillAndOutlineAlphaThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape rectangle = slide.AddRectanglePoints(20, 20, 80, 40);
            rectangle.FillColor = "FFFFFF";
            rectangle.OutlineColor = "000000";
            rectangle.OutlineWidthPoints = 3D;
            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<Shape>()
                .Last();
            ShapeProperties properties = shape.ShapeProperties!;
            A.Outline outline = properties.GetFirstChild<A.Outline>()!;
            A.SolidFill alphaFill = new A.SolidFill(
                new A.RgbColorModelHex(new A.Alpha { Val = 50000 }) { Val = "CC3366" });
            A.SolidFill alphaOutline = new A.SolidFill(
                new A.RgbColorModelHex(new A.Alpha { Val = 50000 }) { Val = "0369A1" });
            properties.RemoveAllChildren<A.SolidFill>();
            properties.InsertBefore(alphaFill, outline);
            outline.RemoveAllChildren<A.SolidFill>();
            outline.InsertAt(alphaOutline, 0);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), drawingShape =>
                Math.Abs(drawingShape.X - 20D) < 0.000001D &&
                Math.Abs(drawingShape.Y - 20D) < 0.000001D);
            Assert.Equal(OfficeColor.FromRgba(204, 51, 102, 128), rendered.Shape.FillColor);
            Assert.Equal(OfficeColor.FromRgba(3, 105, 161, 128), rendered.Shape.StrokeColor);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("fill=\"#CC3366\" fill-opacity=\"0.502\"", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke=\"#0369A1\"", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-opacity=\"0.502\"", svgText, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(160, image!.Width);
            Assert.Equal(100, image.Height);
        }

        [Fact]
        public void PowerPointSlide_ProjectsAutoShapeOutlineCapAndJoinThroughSharedDrawingStrokeStyle() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape arrow = slide.AddShapePoints(A.ShapeTypeValues.RightArrow, 24, 24, 92, 36);
            arrow.FillColor = "E0F2FE";
            arrow.OutlineColor = "0369A1";
            arrow.OutlineWidthPoints = 3D;
            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<Shape>()
                .Last();
            A.Outline outline = shape.ShapeProperties!.GetFirstChild<A.Outline>()!;
            outline.CapType = A.LineCapValues.Round;
            outline.RemoveAllChildren<A.Bevel>();
            outline.RemoveAllChildren<A.Round>();
            outline.RemoveAllChildren<A.Miter>();
            outline.Append(new A.Bevel());

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingShape drawingShape = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), element =>
                Math.Abs(element.X - 24D) < 0.000001D &&
                Math.Abs(element.Y - 24D) < 0.000001D);
            Assert.Equal(OfficeStrokeLineCap.Round, drawingShape.Shape.StrokeLineCap);
            Assert.Equal(OfficeStrokeLineJoin.Bevel, drawingShape.Shape.StrokeLineJoin);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(160, image!.Width);
            Assert.Equal(100, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("stroke-linecap=\"round\"", svgText, StringComparison.Ordinal);
            Assert.Contains("stroke-linejoin=\"bevel\"", svgText, StringComparison.Ordinal);
            Assert.Contains("#0369A1", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsThemeTextBoxFrameThroughSharedDrawingShape() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            presentation.SetThemeColor(PowerPointThemeColor.Accent3, "336699");
            presentation.SetThemeColor(PowerPointThemeColor.Accent4, "884422");
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints(string.Empty, 20, 20, 80, 40);
            textBox.FillColor = "FFFFFF";
            textBox.OutlineColor = "000000";
            textBox.OutlineWidthPoints = 1.5D;
            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<Shape>()
                .Last();
            ShapeProperties properties = shape.ShapeProperties!;
            A.Outline outline = properties.GetFirstChild<A.Outline>()!;
            A.SolidFill themeFill = new A.SolidFill(
                new A.SchemeColor(new A.LuminanceModulation { Val = 50000 }) { Val = A.SchemeColorValues.Accent3 });
            A.SolidFill themeOutline = new A.SolidFill(
                new A.SchemeColor(new A.Shade { Val = 50000 }) { Val = A.SchemeColorValues.Accent4 });
            properties.RemoveAllChildren<A.SolidFill>();
            properties.InsertBefore(themeFill, outline);
            outline.RemoveAllChildren<A.SolidFill>();
            outline.InsertAt(themeOutline, 0);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            OfficeDrawingShape rendered = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), drawingShape =>
                Math.Abs(drawingShape.X - 20D) < 0.000001D &&
                Math.Abs(drawingShape.Y - 20D) < 0.000001D);
            Assert.Equal(OfficeColor.FromRgb(26, 51, 76), rendered.Shape.FillColor);
            Assert.Equal(OfficeColor.FromRgb(68, 34, 17), rendered.Shape.StrokeColor);
            Assert.Equal(1.5D, rendered.Shape.StrokeWidth);
            Assert.DoesNotContain(snapshot.Drawing.Elements, element => element is OfficeDrawingText);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("#1A334C", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#442211", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.FromRgb(26, 51, 76), image!.GetPixel(30, 30));
        }

        [Fact]
        public void PowerPointSlide_ExpandsTabsInTextBoxesThroughSharedDrawingLayout() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 80);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints("A\tB", 20, 20, 130, 24);
            textBox.FontSize = 12;

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);

            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.DoesNotContain("\t", svgText, StringComparison.Ordinal);
            Assert.Contains("xml:space=\"preserve\"", svgText, StringComparison.Ordinal);
            Assert.Contains("A   B", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTextMarginsThroughSharedDrawingPadding() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 90);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints("Inset", 20, 18, 120, 44);
            textBox.FontSize = 12;
            textBox.SetTextMarginsPoints(12, 6, 10, 4);

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingText drawingText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>());
            Assert.Equal(20D, drawingText.X);
            Assert.Equal(18D, drawingText.Y);
            Assert.Equal(120D, drawingText.Width);
            Assert.Equal(44D, drawingText.Height);
            Assert.True(drawingText.HasPadding);
            Assert.Equal(12D, drawingText.Padding.Left);
            Assert.Equal(6D, drawingText.Padding.Top);
            Assert.Equal(10D, drawingText.Padding.Right);
            Assert.Equal(4D, drawingText.Padding.Bottom);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<text x=\"32\"", svgText, StringComparison.Ordinal);
            Assert.Contains("Inset", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsParagraphHangingIndentThroughSharedDrawingTextIndent() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(220, 120);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints("PowerPoint hanging indent wraps across lines", 20, 20, 120, 54);
            textBox.FontSize = 12;
            textBox.Paragraphs[0].SetHangingPoints(18D);

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingText drawingText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>());
            Assert.True(drawingText.HasParagraphIndent);
            Assert.Equal(0D, drawingText.ParagraphIndent.FirstLineOffset);
            Assert.Equal(18D, drawingText.ParagraphIndent.ContinuationLineOffset);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("PowerPoint", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsJustifiedTextBoxesThroughSharedDrawingTextAlignment() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(220, 120);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints("Justified PowerPoint text wraps across the exported slide image", 20, 20, 120, 54);
            textBox.FontSize = 12;
            textBox.Paragraphs[0].SetAlignment(A.TextAlignmentTypeValues.Justified);

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            OfficeDrawingText text = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(element => element.Text.StartsWith("Justified PowerPoint", StringComparison.Ordinal));
            Assert.Equal(OfficeTextAlignment.Justify, text.Alignment);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Justified", svgText, StringComparison.Ordinal);
            Assert.Contains("PowerPoint", svgText, StringComparison.Ordinal);
            Assert.Contains("textLength=", svgText, StringComparison.Ordinal);
            Assert.Contains("lengthAdjust=\"spacing\"", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsThemeTextBoxTextThroughSharedDrawingText() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 100);
            presentation.SetThemeColor(PowerPointThemeColor.Accent1, "204060");
            PowerPointSlide slide = presentation.AddSlide();

            slide.AddTextBoxPoints("Theme text", 20, 20, 120, 34);
            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<Shape>()
                .Last();
            A.Run run = shape.TextBody!.Elements<A.Paragraph>().Single().Elements<A.Run>().Single();
            run.RunProperties ??= new A.RunProperties();
            run.RunProperties.RemoveAllChildren<A.SolidFill>();
            run.RunProperties.InsertAt(new A.SolidFill(new A.SchemeColor(new A.LuminanceModulation { Val = 50000 }) { Val = A.SchemeColorValues.Accent1 }), 0);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            OfficeDrawingText text = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), drawingText => drawingText.Text == "Theme text");
            Assert.Equal(OfficeColor.FromRgb(16, 32, 48), text.Color);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Theme", svgText, StringComparison.Ordinal);
            Assert.Contains("text", svgText, StringComparison.Ordinal);
            Assert.Contains("#102030", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(180, image!.Width);
        }

        [Fact]
        public void PowerPointSlide_ProjectsPresetAutoShapesThroughSharedDrawingPresets() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(220, 140);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape triangle = slide.AddShapePoints(A.ShapeTypeValues.Triangle, 20, 20, 48, 44);
            triangle.FillColor = "1F4E79";
            triangle.OutlineColor = "0F243E";
            triangle.OutlineWidthPoints = 1D;

            PowerPointAutoShape panel = slide.AddShapePoints(A.ShapeTypeValues.Parallelogram, 82, 20, 70, 44);
            panel.FillColor = "1976D2";
            panel.OutlineColor = "0B3D91";
            panel.OutlineWidthPoints = 1D;

            PowerPointAutoShape arrow = slide.AddShapePoints(A.ShapeTypeValues.RightArrow, 42, 82, 116, 34);
            arrow.FillColor = "16A34A";
            arrow.OutlineColor = "14532D";
            arrow.OutlineWidthPoints = 1D;

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            Assert.Equal(3, snapshot.Drawing.Elements.OfType<OfficeDrawingShape>().Count(shape => shape.Shape.Kind == OfficeShapeKind.Polygon));
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(220, image!.Width);
            Assert.Equal(140, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<polygon", svgText, StringComparison.Ordinal);
            Assert.Contains("#1F4E79", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#1976D2", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#16A34A", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsAdditionalPolygonAndStarPresetsThroughSharedDrawingPresets() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(260, 170);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape heptagon = slide.AddShapePoints(A.ShapeTypeValues.Heptagon, 16, 18, 52, 48);
            heptagon.FillColor = "DBEAFE";
            heptagon.OutlineColor = "1D4ED8";

            PowerPointAutoShape decagon = slide.AddShapePoints(A.ShapeTypeValues.Decagon, 86, 18, 56, 48);
            decagon.FillColor = "FEF3C7";
            decagon.OutlineColor = "B45309";

            PowerPointAutoShape dodecagon = slide.AddShapePoints(A.ShapeTypeValues.Dodecagon, 162, 18, 58, 48);
            dodecagon.FillColor = "DCFCE7";
            dodecagon.OutlineColor = "15803D";

            PowerPointAutoShape star4 = slide.AddShapePoints(A.ShapeTypeValues.Star4, 24, 94, 48, 48);
            star4.FillColor = "FCE7F3";
            star4.OutlineColor = "BE185D";

            PowerPointAutoShape star8 = slide.AddShapePoints(A.ShapeTypeValues.Star8, 102, 94, 48, 48);
            star8.FillColor = "E0F2FE";
            star8.OutlineColor = "0369A1";

            PowerPointAutoShape star16 = slide.AddShapePoints(A.ShapeTypeValues.Star16, 180, 94, 48, 48);
            star16.FillColor = "FEE2E2";
            star16.OutlineColor = "B91C1C";

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            List<OfficeDrawingShape> shapes = snapshot.Drawing.Elements.OfType<OfficeDrawingShape>().ToList();
            Assert.True(shapes.Count >= 6);
            Assert.True(shapes.Count(shape => shape.Shape.Kind == OfficeShapeKind.Polygon) >= 6);
            Assert.Contains(shapes, shape => shape.Shape.Points.Count == 7);
            Assert.Contains(shapes, shape => shape.Shape.Points.Count == 10);
            Assert.Contains(shapes, shape => shape.Shape.Points.Count == 12);
            Assert.Contains(shapes, shape => shape.Shape.Points.Count == 32);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(260, image!.Width);
            Assert.Equal(170, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<polygon", svgText, StringComparison.Ordinal);
            Assert.Contains("#DBEAFE", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FEF3C7", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#DCFCE7", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FCE7F3", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#E0F2FE", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FEE2E2", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsMultiDirectionArrowPresetsThroughSharedDrawingPresets() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 120);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape upDown = slide.AddShapePoints(A.ShapeTypeValues.UpDownArrow, 20, 20, 48, 78);
            upDown.FillColor = "A7F3D0";
            upDown.OutlineColor = "047857";

            PowerPointAutoShape quad = slide.AddShapePoints(A.ShapeTypeValues.QuadArrow, 92, 22, 66, 66);
            quad.FillColor = "FBCFE8";
            quad.OutlineColor = "BE185D";

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            List<OfficeDrawingShape> shapes = snapshot.Drawing.Elements.OfType<OfficeDrawingShape>().ToList();
            Assert.Equal(2, shapes.Count(shape => shape.Shape.Kind == OfficeShapeKind.Polygon));
            Assert.Contains(shapes, shape => shape.Shape.Points.Count == 10);
            Assert.Contains(shapes, shape => shape.Shape.Points.Count == 24);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(180, image!.Width);
            Assert.Equal(120, image.Height);

        string svgText = Encoding.UTF8.GetString(svg.Bytes);
        Assert.Contains("<polygon", svgText, StringComparison.Ordinal);
        Assert.Contains("#A7F3D0", svgText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#FBCFE8", svgText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PowerPointSlide_ProjectsAdditionalArrowPresetsThroughSharedDrawingPresets() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 170);
        PowerPointSlide slide = presentation.AddSlide();

        PowerPointAutoShape leftUp = slide.AddShapePoints(A.ShapeTypeValues.LeftUpArrow, 18, 18, 56, 54);
        leftUp.FillColor = "DBEAFE";
        leftUp.OutlineColor = "1D4ED8";

        PowerPointAutoShape leftRightUp = slide.AddShapePoints(A.ShapeTypeValues.LeftRightUpArrow, 98, 16, 72, 56);
        leftRightUp.FillColor = "DCFCE7";
        leftRightUp.OutlineColor = "15803D";

        PowerPointAutoShape bentUp = slide.AddShapePoints(A.ShapeTypeValues.BentUpArrow, 28, 96, 72, 48);
        bentUp.FillColor = "FEF3C7";
        bentUp.OutlineColor = "B45309";

        PowerPointAutoShape uTurn = slide.AddShapePoints(A.ShapeTypeValues.UTurnArrow, 150, 88, 62, 62);
        uTurn.FillColor = "FCE7F3";
        uTurn.OutlineColor = "BE185D";

        OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
        OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
        PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

        AssertNoUnexpectedDiagnostics(png.Diagnostics);
        AssertNoUnexpectedDiagnostics(svg.Diagnostics);
        AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
        List<OfficeDrawingShape> shapes = snapshot.Drawing.Elements.OfType<OfficeDrawingShape>().ToList();
        Assert.Equal(4, shapes.Count(shape => shape.Shape.Kind == OfficeShapeKind.Polygon));
        Assert.Contains(shapes, shape => shape.Shape.Points.Count == 17);
        Assert.Contains(shapes, shape => shape.Shape.Points.Count == 11);
        Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
        Assert.Equal(260, image!.Width);
        Assert.Equal(170, image.Height);

        string svgText = Encoding.UTF8.GetString(svg.Bytes);
        Assert.Contains("<polygon", svgText, StringComparison.Ordinal);
        Assert.Contains("#DBEAFE", svgText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#DCFCE7", svgText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#FEF3C7", svgText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#FCE7F3", svgText, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PowerPointSlide_ProjectsArrowCalloutPresetsThroughSharedDrawingPresets() {
        using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 160);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape rightCallout = slide.AddShapePoints(A.ShapeTypeValues.RightArrowCallout, 16, 18, 84, 42);
            rightCallout.FillColor = "DBEAFE";
            rightCallout.OutlineColor = "1D4ED8";

            PowerPointAutoShape upDownCallout = slide.AddShapePoints(A.ShapeTypeValues.UpDownArrowCallout, 118, 14, 42, 80);
            upDownCallout.FillColor = "DCFCE7";
            upDownCallout.OutlineColor = "15803D";

            PowerPointAutoShape leftRightCallout = slide.AddShapePoints(A.ShapeTypeValues.LeftRightArrowCallout, 20, 104, 92, 36);
            leftRightCallout.FillColor = "FEF3C7";
            leftRightCallout.OutlineColor = "B45309";

            PowerPointAutoShape quadCallout = slide.AddShapePoints(A.ShapeTypeValues.QuadArrowCallout, 170, 78, 52, 52);
            quadCallout.FillColor = "FBCFE8";
            quadCallout.OutlineColor = "BE185D";

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            List<OfficeDrawingShape> shapes = snapshot.Drawing.Elements.OfType<OfficeDrawingShape>().ToList();
            Assert.Equal(4, shapes.Count(shape => shape.Shape.Kind == OfficeShapeKind.Polygon));
            Assert.Contains(shapes, shape => shape.Shape.Points.Count == 7);
            Assert.Contains(shapes, shape => shape.Shape.Points.Count == 10);
            Assert.Contains(shapes, shape => shape.Shape.Points.Count == 24);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(240, image!.Width);
            Assert.Equal(160, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<polygon", svgText, StringComparison.Ordinal);
            Assert.Contains("#DBEAFE", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#DCFCE7", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FEF3C7", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FBCFE8", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsFlowChartPresetAutoShapesThroughSharedDrawingPresets() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 170);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape process = slide.AddShapePoints(A.ShapeTypeValues.FlowChartProcess, 18, 18, 76, 38);
            process.FillColor = "DBEAFE";
            process.OutlineColor = "1D4ED8";

            PowerPointAutoShape decision = slide.AddShapePoints(A.ShapeTypeValues.FlowChartDecision, 118, 12, 64, 52);
            decision.FillColor = "FEF3C7";
            decision.OutlineColor = "D97706";

            PowerPointAutoShape inputOutput = slide.AddShapePoints(A.ShapeTypeValues.FlowChartInputOutput, 20, 82, 76, 38);
            inputOutput.FillColor = "DCFCE7";
            inputOutput.OutlineColor = "16A34A";

            PowerPointAutoShape terminator = slide.AddShapePoints(A.ShapeTypeValues.FlowChartTerminator, 120, 84, 82, 36);
            terminator.FillColor = "FCE7F3";
            terminator.OutlineColor = "BE185D";

            PowerPointAutoShape document = slide.AddShapePoints(A.ShapeTypeValues.FlowChartDocument, 62, 132, 94, 30);
            document.FillColor = "E0F2FE";
            document.OutlineColor = "0369A1";

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            List<OfficeDrawingShape> shapes = snapshot.Drawing.Elements.OfType<OfficeDrawingShape>().ToList();
            Assert.Contains(shapes, shape => shape.Shape.Kind == OfficeShapeKind.Rectangle);
            Assert.Contains(shapes, shape => shape.Shape.Kind == OfficeShapeKind.RoundedRectangle);
            Assert.True(shapes.Count(shape => shape.Shape.Kind == OfficeShapeKind.Polygon) >= 2);
            Assert.Contains(shapes, shape => shape.Shape.Kind == OfficeShapeKind.Path);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(240, image!.Width);
            Assert.Equal(170, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<rect", svgText, StringComparison.Ordinal);
            Assert.Contains("<polygon", svgText, StringComparison.Ordinal);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("#DBEAFE", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FEF3C7", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#DCFCE7", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FCE7F3", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#E0F2FE", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsAdditionalFlowChartPresetsThroughSharedDrawingPresets() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(260, 190);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape preparation = slide.AddShapePoints(A.ShapeTypeValues.FlowChartPreparation, 16, 18, 64, 42);
            preparation.FillColor = "DBEAFE";
            preparation.OutlineColor = "1D4ED8";

            PowerPointAutoShape manualInput = slide.AddShapePoints(A.ShapeTypeValues.FlowChartManualInput, 98, 18, 74, 42);
            manualInput.FillColor = "FEF3C7";
            manualInput.OutlineColor = "B45309";

            PowerPointAutoShape manualOperation = slide.AddShapePoints(A.ShapeTypeValues.FlowChartManualOperation, 18, 82, 70, 42);
            manualOperation.FillColor = "DCFCE7";
            manualOperation.OutlineColor = "15803D";

            PowerPointAutoShape delay = slide.AddShapePoints(A.ShapeTypeValues.FlowChartDelay, 108, 82, 64, 42);
            delay.FillColor = "FCE7F3";
            delay.OutlineColor = "BE185D";

            PowerPointAutoShape offpage = slide.AddShapePoints(A.ShapeTypeValues.FlowChartOffpageConnector, 184, 42, 38, 54);
            offpage.FillColor = "E0F2FE";
            offpage.OutlineColor = "0369A1";

            PowerPointAutoShape magneticTape = slide.AddShapePoints(A.ShapeTypeValues.FlowChartMagneticTape, 18, 136, 58, 38);
            magneticTape.FillColor = "DDD6FE";
            magneticTape.OutlineColor = "6D28D9";

            PowerPointAutoShape magneticDrum = slide.AddShapePoints(A.ShapeTypeValues.FlowChartMagneticDrum, 98, 136, 58, 38);
            magneticDrum.FillColor = "FEE2E2";
            magneticDrum.OutlineColor = "B91C1C";

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            List<OfficeDrawingShape> shapes = snapshot.Drawing.Elements.OfType<OfficeDrawingShape>().ToList();
            Assert.True(shapes.Count >= 7);
            Assert.True(shapes.Count(shape => shape.Shape.Kind == OfficeShapeKind.Polygon) >= 4);
            Assert.True(shapes.Count(shape => shape.Shape.Kind == OfficeShapeKind.Path) >= 3);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(260, image!.Width);
            Assert.Equal(190, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<polygon", svgText, StringComparison.Ordinal);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("#DBEAFE", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FEF3C7", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#DCFCE7", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FCE7F3", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#E0F2FE", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#DDD6FE", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FEE2E2", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsCalloutAndSymbolPresetAutoShapesThroughSharedDrawingPresets() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(260, 170);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape rectangleCallout = slide.AddShapePoints(A.ShapeTypeValues.WedgeRectangleCallout, 16, 16, 72, 50);
            rectangleCallout.FillColor = "FDE68A";
            rectangleCallout.OutlineColor = "92400E";

            PowerPointAutoShape roundedCallout = slide.AddShapePoints(A.ShapeTypeValues.WedgeRoundRectangleCallout, 104, 16, 82, 52);
            roundedCallout.FillColor = "BFDBFE";
            roundedCallout.OutlineColor = "1D4ED8";

            PowerPointAutoShape cloudCallout = slide.AddShapePoints(A.ShapeTypeValues.CloudCallout, 24, 94, 82, 52);
            cloudCallout.FillColor = "DCFCE7";
            cloudCallout.OutlineColor = "15803D";

            PowerPointAutoShape lightningBolt = slide.AddShapePoints(A.ShapeTypeValues.LightningBolt, 162, 94, 44, 56);
            lightningBolt.FillColor = "FACC15";
            lightningBolt.OutlineColor = "854D0E";

            PowerPointAutoShape moon = slide.AddShapePoints(A.ShapeTypeValues.Moon, 212, 22, 34, 44);
            moon.FillColor = "E0E7FF";
            moon.OutlineColor = "3730A3";

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            List<OfficeDrawingShape> shapes = snapshot.Drawing.Elements.OfType<OfficeDrawingShape>().ToList();
            Assert.True(shapes.Count >= 5);
            Assert.Contains(shapes, shape => shape.Shape.Kind == OfficeShapeKind.Polygon);
            Assert.True(shapes.Count(shape => shape.Shape.Kind == OfficeShapeKind.Path) >= 3);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(260, image!.Width);
            Assert.Equal(170, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<polygon", svgText, StringComparison.Ordinal);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("#FDE68A", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#BFDBFE", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#DCFCE7", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FACC15", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#E0E7FF", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsBracketAndBracePresetsThroughSharedDrawingPresets() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 120);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape bracketPair = slide.AddShapePoints(A.ShapeTypeValues.BracketPair, 18, 18, 62, 84);
            bracketPair.FillColor = "DDD6FE";
            bracketPair.OutlineColor = "6D28D9";

            PowerPointAutoShape bracePair = slide.AddShapePoints(A.ShapeTypeValues.BracePair, 98, 18, 62, 84);
            bracePair.FillColor = "FED7AA";
            bracePair.OutlineColor = "C2410C";

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            List<OfficeDrawingShape> pathShapes = snapshot.Drawing.Elements.OfType<OfficeDrawingShape>()
                .Where(shape => shape.Shape.Kind == OfficeShapeKind.Path)
                .ToList();
            Assert.Equal(2, pathShapes.Count);
            Assert.All(pathShapes, shape => Assert.Equal(2, shape.Shape.PathCommands.Count(command => command.Kind == OfficePathCommandKind.MoveTo)));
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(180, image!.Width);
            Assert.Equal(120, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("#DDD6FE", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FED7AA", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsGroupedShapesThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 160);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape first = slide.AddRectanglePoints(20, 20, 30, 20);
            first.FillColor = "FF0000";
            first.OutlineColor = "7F0000";
            first.OutlineWidthPoints = 2D;

            PowerPointAutoShape second = slide.AddRectanglePoints(60, 20, 30, 20);
            second.FillColor = "00AA00";
            second.OutlineColor = "006600";
            second.OutlineWidthPoints = 1.5D;

            slide.GroupShapes(new PowerPointShape[] { first, second }, "Scaled group");
            GroupShape group = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<GroupShape>()
                .Single();
            A.TransformGroup transform = group.GroupShapeProperties!.TransformGroup!;
            transform.Extents!.Cx = PowerPointUnits.FromPoints(140);
            transform.Extents.Cy = PowerPointUnits.FromPoints(40);
            transform.ChildExtents!.Cx = PowerPointUnits.FromPoints(70);
            transform.ChildExtents.Cy = PowerPointUnits.FromPoints(20);
            slide.SlidePart.Slide.Save();

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingShape red = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), element => element.Shape.FillColor == OfficeColor.FromRgb(255, 0, 0));
            OfficeDrawingShape green = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), element => element.Shape.FillColor == OfficeColor.FromRgb(0, 170, 0));
            Assert.Equal(20D, red.X, 6);
            Assert.Equal(20D, red.Y, 6);
            Assert.Equal(60D, red.Shape.Width, 6);
            Assert.Equal(40D, red.Shape.Height, 6);
            Assert.Equal(4D, red.Shape.StrokeWidth, 6);
            Assert.Equal(100D, green.X, 6);
            Assert.Equal(20D, green.Y, 6);
            Assert.Equal(60D, green.Shape.Width, 6);
            Assert.Equal(40D, green.Shape.Height, 6);
            Assert.Equal(3D, green.Shape.StrokeWidth, 6);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.FromRgb(255, 0, 0), image!.GetPixel(30, 30));
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("#FF0000", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#00AA00", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-width=\"4\"", svgText, StringComparison.Ordinal);
            Assert.Contains("stroke-width=\"3\"", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_UsesVisibleFallbackForUndecodableGroupedPictures() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointPicture picture = slide.AddPicture(
                new MemoryStream(new byte[] { 1, 2, 3, 4, 5, 6 }),
                ImagePartType.Bmp,
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(24),
                PowerPointUnits.FromPoints(18));
            PowerPointAutoShape anchor = slide.AddRectanglePoints(50, 20, 12, 12);
            slide.GroupShapes(new PowerPointShape[] { picture, anchor }, "Grouped unsupported image");

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);

            OfficeImageExportDiagnostic diagnostic = Assert.Single(
                png.Diagnostics,
                item => item.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTransformedGroupedShapesThroughSharedDrawingComposition() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(220, 140);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape first = slide.AddRectanglePoints(40, 35, 30, 20);
            first.FillColor = "FF0000";
            first.OutlineColor = "7F0000";

            PowerPointAutoShape second = slide.AddRectanglePoints(80, 35, 30, 20);
            second.FillColor = "00AA00";
            second.OutlineColor = "006600";

            slide.GroupShapes(new PowerPointShape[] { first, second }, "Transformed group");
            GroupShape group = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<GroupShape>()
                .Single();
            A.TransformGroup transform = group.GroupShapeProperties!.TransformGroup!;
            transform.Rotation = 12 * 60000;
            transform.HorizontalFlip = true;
            slide.SlidePart.Slide.Save();

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingShape red = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), element => element.Shape.FillColor == OfficeColor.FromRgb(255, 0, 0));
            OfficeDrawingShape green = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), element => element.Shape.FillColor == OfficeColor.FromRgb(0, 170, 0));
            Assert.True(red.Shape.Transform.HasValue);
            Assert.True(green.Shape.Transform.HasValue);
            Assert.Equal(40D, red.X, 6);
            Assert.Equal(35D, red.Y, 6);
            Assert.Equal(80D, green.X, 6);
            Assert.Equal(35D, green.Y, 6);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(220, image!.Width);
            Assert.Equal(140, image.Height);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("#FF0000", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#00AA00", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("matrix(", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_RendersRotatedGroupedPicturesThroughSharedDrawingComposition() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 120);
            PowerPointSlide slide = presentation.AddSlide();

            byte[] pngBytes = OfficePngWriter.Encode(new OfficeRasterImage(4, 4, OfficeColor.FromRgb(37, 99, 235)));
            PowerPointPicture picture = slide.AddPicture(
                new MemoryStream(pngBytes),
                ImagePartType.Png,
                PowerPointUnits.FromPoints(30),
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(80),
                PowerPointUnits.FromPoints(80));
            PowerPointAutoShape anchor = slide.AddRectanglePoints(68, 58, 4, 4);
            anchor.FillColor = "2563EB";
            anchor.OutlineColor = "2563EB";
            slide.GroupShapes(new PowerPointShape[] { picture, anchor }, "Rotated picture group");
            GroupShape group = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<GroupShape>()
                .Single();
            A.TransformGroup transform = group.GroupShapeProperties!.TransformGroup!;
            transform.Rotation = 45 * 60000;
            slide.SlidePart.Slide.Save();

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingImage>());
            Assert.True(drawingImage.Projection.HasTransform);
            Assert.Equal(45D, drawingImage.Projection.RotationDegrees, 6);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(160, image!.Width);
            Assert.Equal(120, image.Height);
        }

        [Fact]
        public void PowerPointSlide_ClipsOverflowingGroupedShapesThroughSharedDrawingComposition() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape shape = slide.AddRectanglePoints(20, 20, 100, 20);
            shape.FillColor = "FF0000";
            shape.OutlineColor = "7F0000";
            PowerPointAutoShape anchor = slide.AddRectanglePoints(62, 22, 4, 4);
            anchor.FillColor = "E5E7EB";
            anchor.OutlineColor = "9CA3AF";

            slide.GroupShapes(new PowerPointShape[] { shape, anchor }, "Clipped group");
            GroupShape group = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<GroupShape>()
                .Single();
            A.TransformGroup transform = group.GroupShapeProperties!.TransformGroup!;
            transform.Extents!.Cx = PowerPointUnits.FromPoints(50);
            transform.ChildExtents!.Cx = PowerPointUnits.FromPoints(50);
            slide.SlidePart.Slide.Save();

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingGroup drawingGroup = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingGroup>());
            Assert.Equal(20D, drawingGroup.X, 6);
            Assert.Equal(20D, drawingGroup.Y, 6);
            Assert.Equal(50D, drawingGroup.ClipPath.Width, 6);
            Assert.Contains(drawingGroup.Drawing.Elements, element =>
                element is OfficeDrawingShape drawingShape &&
                drawingShape.Shape.FillColor == OfficeColor.FromRgb(255, 0, 0) &&
                drawingShape.Shape.Width > drawingGroup.ClipPath.Width);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.FromRgb(255, 0, 0), image!.GetPixel(25, 25));
            Assert.NotEqual(OfficeColor.FromRgb(255, 0, 0), image.GetPixel(75, 25));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("officeimo-group-clip-", svgText, StringComparison.Ordinal);
            Assert.Contains("#FF0000", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ClipsTransformedOverflowingGroupedShapesThroughSharedDrawingComposition() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape shape = slide.AddRectanglePoints(20, 20, 100, 20);
            shape.FillColor = "FF0000";
            shape.OutlineColor = "7F0000";
            PowerPointAutoShape anchor = slide.AddRectanglePoints(62, 22, 4, 4);
            anchor.FillColor = "E5E7EB";
            anchor.OutlineColor = "9CA3AF";

            slide.GroupShapes(new PowerPointShape[] { shape, anchor }, "Transformed clipped group");
            GroupShape group = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<GroupShape>()
                .Single();
            A.TransformGroup transform = group.GroupShapeProperties!.TransformGroup!;
            transform.Extents!.Cx = PowerPointUnits.FromPoints(50);
            transform.ChildExtents!.Cx = PowerPointUnits.FromPoints(50);
            transform.HorizontalFlip = true;
            slide.SlidePart.Slide.Save();

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingGroup drawingGroup = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingGroup>());
            Assert.True(drawingGroup.FrameTransform.HasValue);
            Assert.True(drawingGroup.FrameTransform.Value.FlipHorizontal);
            Assert.Equal(20D, drawingGroup.X, 6);
            Assert.Equal(20D, drawingGroup.Y, 6);
            Assert.Equal(50D, drawingGroup.ClipPath.Width, 6);
            Assert.Contains(drawingGroup.Drawing.Elements, element =>
                element is OfficeDrawingShape drawingShape &&
                drawingShape.Shape.FillColor == OfficeColor.FromRgb(255, 0, 0) &&
                drawingShape.Shape.Width > drawingGroup.ClipPath.Width);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.FromRgb(255, 0, 0), image!.GetPixel(35, 25));
            Assert.NotEqual(OfficeColor.FromRgb(255, 0, 0), image.GetPixel(75, 25));

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("officeimo-group-clip-", svgText, StringComparison.Ordinal);
            Assert.Contains("matrix(", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#FF0000", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsScaledGroupedRichTextMetricsThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 140);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints("Base ", 20, 20, 60, 24);
            textBox.FontSize = 10;
            textBox.SetTextMarginsPoints(3, 2, 4, 1);
            PowerPointParagraph paragraph = textBox.Paragraphs[0];
            paragraph.SetLeftMarginPoints(6);
            paragraph.AddRun("Red", run => {
                run.Color = "FF0000";
                run.FontSize = 12;
                run.Bold = true;
            });

            PowerPointAutoShape anchor = slide.AddRectanglePoints(90, 20, 10, 10);
            anchor.FillColor = "E5E7EB";
            anchor.OutlineColor = "9CA3AF";

            slide.GroupShapes(new PowerPointShape[] { textBox, anchor }, "Scaled text group");
            GroupShape group = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<GroupShape>()
                .Single();
            A.TransformGroup transform = group.GroupShapeProperties!.TransformGroup!;
            transform.Extents!.Cx = PowerPointUnits.FromPoints(160);
            transform.Extents.Cy = PowerPointUnits.FromPoints(48);
            transform.ChildExtents!.Cx = PowerPointUnits.FromPoints(80);
            transform.ChildExtents.Cy = PowerPointUnits.FromPoints(24);
            slide.SlidePart.Slide.Save();

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            Assert.Equal("Base Red", richText.PlainText);
            Assert.Equal(20D, richText.X, 6);
            Assert.Equal(20D, richText.Y, 6);
            Assert.Equal(120D, richText.Width, 6);
            Assert.Equal(48D, richText.Height, 6);
            Assert.Equal(20D, richText.Runs[0].FontSize, 6);
            Assert.Equal(24D, richText.Runs[1].FontSize, 6);
            Assert.Equal(OfficeColor.Red, richText.Runs[1].Color);
            Assert.Equal(6D, richText.Padding.Left, 6);
            Assert.Equal(4D, richText.Padding.Top, 6);
            Assert.Equal(8D, richText.Padding.Right, 6);
            Assert.Equal(2D, richText.Padding.Bottom, 6);
            Assert.Equal(12D, richText.ParagraphIndent.FirstLineOffset, 6);
            Assert.Equal(12D, richText.ParagraphIndent.ContinuationLineOffset, 6);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(240, image!.Width);
            Assert.Equal(140, image.Height);
        }

        [Fact]
        public void PowerPointSlide_ProjectsZeroThicknessConnectorsThroughSharedDrawingLines() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 120);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape horizontal = slide.AddShapePoints(A.ShapeTypeValues.StraightConnector1, 20, 40, 100, 0);
            horizontal.OutlineColor = "1E5A96";
            horizontal.OutlineWidthPoints = 2D;

            PowerPointAutoShape vertical = slide.AddShapePoints(A.ShapeTypeValues.StraightConnector1, 140, 30, 0, 70);
            vertical.OutlineColor = "C00000";
            vertical.OutlineWidthPoints = 2D;

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            Assert.Equal(2, snapshot.Drawing.Elements.OfType<OfficeDrawingShape>().Count(element => element.Shape.Kind == OfficeShapeKind.Line));
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(180, image!.Width);
            Assert.Equal(120, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<line", svgText, StringComparison.Ordinal);
            Assert.Contains("#1E5A96", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#C00000", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsConnectorLineEndsThroughSharedDrawingMarkers() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 120);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape horizontal = slide.AddShapePoints(A.ShapeTypeValues.StraightConnector1, 20, 40, 100, 0);
            horizontal.OutlineColor = "1E5A96";
            horizontal.OutlineWidthPoints = 2D;
            horizontal.SetLineEnds(null, A.LineEndValues.Triangle, A.LineEndWidthValues.Large, A.LineEndLengthValues.Large);

            PowerPointAutoShape vertical = slide.AddShapePoints(A.ShapeTypeValues.StraightConnector1, 140, 30, 0, 70);
            vertical.OutlineColor = "C00000";
            vertical.OutlineWidthPoints = 2D;
            vertical.SetLineEnds(A.LineEndValues.Diamond, null, A.LineEndWidthValues.Medium, A.LineEndLengthValues.Medium);

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            IReadOnlyList<OfficeDrawingShape> lines = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(element => element.Shape.Kind == OfficeShapeKind.Line)
                .ToArray();
            Assert.Equal(2, lines.Count);
            OfficeDrawingShape horizontalLine = lines.Single(line => line.Shape.StrokeColor == OfficeColor.FromRgb(30, 90, 150));
            OfficeDrawingShape verticalLine = lines.Single(line => line.Shape.StrokeColor == OfficeColor.FromRgb(192, 0, 0));
            Assert.Equal(OfficeLineMarkerKind.Triangle, horizontalLine.Shape.StrokeEndMarker?.Kind);
            Assert.Equal(OfficeLineMarkerKind.Diamond, verticalLine.Shape.StrokeStartMarker?.Kind);
            Assert.Null(horizontalLine.Shape.StrokeStartMarker);
            Assert.Null(verticalLine.Shape.StrokeEndMarker);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(180, image!.Width);
            Assert.Equal(120, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<polygon", svgText, StringComparison.Ordinal);
            Assert.Contains("#1E5A96", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#C00000", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsBentConnectorsThroughSharedDrawingPaths() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 120);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape connector = slide.AddShapePoints(A.ShapeTypeValues.BentConnector2, 24, 24, 92, 58);
            connector.OutlineColor = "1E5A96";
            connector.OutlineWidthPoints = 2D;
            connector.SetLineEnds(null, A.LineEndValues.Triangle, A.LineEndWidthValues.Medium, A.LineEndLengthValues.Medium);

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingShape connectorPath = Assert.Single(
                snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(),
                element => element.Shape.Kind == OfficeShapeKind.Path);
            Assert.Equal(OfficeLineMarkerKind.Triangle, connectorPath.Shape.StrokeEndMarker?.Kind);
            Assert.Null(connectorPath.Shape.StrokeStartMarker);
            Assert.Equal(3, connectorPath.Shape.PathCommands.Count);
            Assert.Equal(OfficePathCommand.MoveTo(0, 0), connectorPath.Shape.PathCommands[0]);
            Assert.Equal(OfficePathCommand.LineTo(0, 58), connectorPath.Shape.PathCommands[1]);
            Assert.Equal(OfficePathCommand.LineTo(92, 58), connectorPath.Shape.PathCommands[2]);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(180, image!.Width);
            Assert.Equal(120, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("<polygon", svgText, StringComparison.Ordinal);
            Assert.Contains("#1E5A96", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsMultiElbowBentConnectorsThroughSharedDrawingPaths() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(220, 150);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape bent4 = slide.AddShapePoints(A.ShapeTypeValues.BentConnector4, 20, 24, 96, 64);
            bent4.OutlineColor = "1E5A96";
            bent4.OutlineWidthPoints = 2D;
            bent4.SetLineEnds(A.LineEndValues.Diamond, A.LineEndValues.Triangle, A.LineEndWidthValues.Medium, A.LineEndLengthValues.Medium);

            PowerPointAutoShape bent5 = slide.AddShapePoints(A.ShapeTypeValues.BentConnector5, 122, 36, 72, 72);
            bent5.OutlineColor = "C00000";
            bent5.OutlineWidthPoints = 2D;
            bent5.SetLineEnds(null, A.LineEndValues.Triangle, A.LineEndWidthValues.Medium, A.LineEndLengthValues.Medium);

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            IReadOnlyList<OfficeDrawingShape> connectorPaths = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(element => element.Shape.Kind == OfficeShapeKind.Path)
                .ToArray();
            Assert.Equal(2, connectorPaths.Count);

            OfficeDrawingShape blue = connectorPaths.Single(path => path.Shape.StrokeColor == OfficeColor.FromRgb(30, 90, 150));
            Assert.Equal(OfficeLineMarkerKind.Diamond, blue.Shape.StrokeStartMarker?.Kind);
            Assert.Equal(OfficeLineMarkerKind.Triangle, blue.Shape.StrokeEndMarker?.Kind);
            Assert.Equal(5, blue.Shape.PathCommands.Count);
            Assert.Equal(OfficePathCommand.MoveTo(0, 0), blue.Shape.PathCommands[0]);
            Assert.Equal(OfficePathCommand.LineTo(48, 0), blue.Shape.PathCommands[1]);
            Assert.Equal(OfficePathCommand.LineTo(48, 32), blue.Shape.PathCommands[2]);
            Assert.Equal(OfficePathCommand.LineTo(96, 32), blue.Shape.PathCommands[3]);
            Assert.Equal(OfficePathCommand.LineTo(96, 64), blue.Shape.PathCommands[4]);

            OfficeDrawingShape red = connectorPaths.Single(path => path.Shape.StrokeColor == OfficeColor.FromRgb(192, 0, 0));
            Assert.Null(red.Shape.StrokeStartMarker);
            Assert.Equal(OfficeLineMarkerKind.Triangle, red.Shape.StrokeEndMarker?.Kind);
            Assert.Equal(6, red.Shape.PathCommands.Count);
            Assert.Equal(OfficePathCommand.MoveTo(0, 0), red.Shape.PathCommands[0]);
            Assert.Equal(OfficePathCommand.LineTo(24, 0), red.Shape.PathCommands[1]);
            Assert.Equal(OfficePathCommand.LineTo(24, 36), red.Shape.PathCommands[2]);
            Assert.Equal(OfficePathCommand.LineTo(48, 36), red.Shape.PathCommands[3]);
            Assert.Equal(OfficePathCommand.LineTo(48, 72), red.Shape.PathCommands[4]);
            Assert.Equal(OfficePathCommand.LineTo(72, 72), red.Shape.PathCommands[5]);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(220, image!.Width);
            Assert.Equal(150, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("<polygon", svgText, StringComparison.Ordinal);
            Assert.Contains("#1E5A96", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#C00000", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsCurvedConnectorsThroughSharedDrawingPaths() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(220, 150);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape curved2 = slide.AddShapePoints(A.ShapeTypeValues.CurvedConnector2, 20, 28, 86, 56);
            curved2.OutlineColor = "1E5A96";
            curved2.OutlineWidthPoints = 2D;
            curved2.SetLineEnds(null, A.LineEndValues.Triangle, A.LineEndWidthValues.Medium, A.LineEndLengthValues.Medium);

            PowerPointAutoShape curved5 = slide.AddShapePoints(A.ShapeTypeValues.CurvedConnector5, 122, 34, 72, 74);
            curved5.OutlineColor = "C00000";
            curved5.OutlineWidthPoints = 2D;
            curved5.SetLineEnds(A.LineEndValues.Diamond, A.LineEndValues.Triangle, A.LineEndWidthValues.Medium, A.LineEndLengthValues.Medium);

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            IReadOnlyList<OfficeDrawingShape> connectorPaths = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(element => element.Shape.Kind == OfficeShapeKind.Path)
                .ToArray();
            Assert.Equal(2, connectorPaths.Count);

            OfficeDrawingShape blue = connectorPaths.Single(path => path.Shape.StrokeColor == OfficeColor.FromRgb(30, 90, 150));
            Assert.Null(blue.Shape.StrokeStartMarker);
            Assert.Equal(OfficeLineMarkerKind.Triangle, blue.Shape.StrokeEndMarker?.Kind);
            Assert.Equal(2, blue.Shape.PathCommands.Count);
            Assert.Equal(OfficePathCommand.MoveTo(0, 0), blue.Shape.PathCommands[0]);
            Assert.Equal(OfficePathCommand.CubicBezierTo(0, 56, 86, 0, 86, 56), blue.Shape.PathCommands[1]);

            OfficeDrawingShape red = connectorPaths.Single(path => path.Shape.StrokeColor == OfficeColor.FromRgb(192, 0, 0));
            Assert.Equal(OfficeLineMarkerKind.Diamond, red.Shape.StrokeStartMarker?.Kind);
            Assert.Equal(OfficeLineMarkerKind.Triangle, red.Shape.StrokeEndMarker?.Kind);
            Assert.Equal(5, red.Shape.PathCommands.Count);
            Assert.All(red.Shape.PathCommands.Skip(1), command => Assert.Equal(OfficePathCommandKind.CubicBezierTo, command.Kind));

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(220, image!.Width);
            Assert.Equal(150, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("<polygon", svgText, StringComparison.Ordinal);
            Assert.Contains("#1E5A96", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#C00000", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_LoadsNativeConnectionShapesAndProjectsThemThroughSharedDrawingPaths() {
            using var stream = new MemoryStream();
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
                presentation.SlideSize.SetSizePoints(220, 150);
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddRectanglePoints(16, 20, 32, 24, "Start Node");
                slide.AddRectanglePoints(162, 96, 32, 24, "End Node");
                slide.SlidePart.Slide.CommonSlideData!.ShapeTree!.Append(CreateNativeBentConnectionShape());
                presentation.Save();
            }

            stream.Position = 0;
            using PowerPointPresentation loaded = PowerPointPresentation.Load(stream, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly });
            PowerPointSlide loadedSlide = loaded.Slides[0];

            PowerPointConnectionShape connection = Assert.Single(loadedSlide.Shapes.OfType<PowerPointConnectionShape>());
            Assert.Equal(PowerPointShapeContentType.Connector, connection.ShapeContentType);
            Assert.Equal("Native Bent Connector", connection.Name);
            Assert.Equal(A.ShapeTypeValues.BentConnector4, connection.ShapeType);

            PowerPointSlideVisualSnapshot snapshot = loadedSlide.CreateVisualSnapshot();
            OfficeImageExportResult png = loadedSlide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = loadedSlide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            OfficeDrawingShape connectorPath = Assert.Single(
                snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(),
                element => element.Shape.Kind == OfficeShapeKind.Path &&
                    element.Shape.StrokeColor == OfficeColor.FromRgb(30, 90, 150));
            Assert.Equal(OfficeLineMarkerKind.Diamond, connectorPath.Shape.StrokeStartMarker?.Kind);
            Assert.Equal(OfficeLineMarkerKind.Triangle, connectorPath.Shape.StrokeEndMarker?.Kind);
            Assert.Equal(5, connectorPath.Shape.PathCommands.Count);
            Assert.Equal(OfficePathCommand.LineTo(48, 0), connectorPath.Shape.PathCommands[1]);
            Assert.Equal(OfficePathCommand.LineTo(48, 32), connectorPath.Shape.PathCommands[2]);
            Assert.Equal(OfficePathCommand.LineTo(96, 32), connectorPath.Shape.PathCommands[3]);
            Assert.Equal(OfficePathCommand.LineTo(96, 64), connectorPath.Shape.PathCommands[4]);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(220, image!.Width);
            Assert.Equal(150, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("<polygon", svgText, StringComparison.Ordinal);
            Assert.Contains("#1E5A96", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_RendersPictureThroughSharedDrawingImageElement() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();
            byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.CornflowerBlue));
            slide.AddPicture(new MemoryStream(png), ImagePartType.Png, PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(20));

            OfficeImageExportResult pngResult = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svgResult = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(pngResult.Diagnostics);
            AssertNoUnexpectedDiagnostics(svgResult.Diagnostics);
            Assert.Single(snapshot.Drawing.Images);
            Assert.True(OfficePngReader.TryDecode(pngResult.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.CornflowerBlue, image!.GetPixel(25, 25));

            string svgText = Encoding.UTF8.GetString(svgResult.Bytes);
            Assert.Contains("<image", svgText, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_RasterizesExistingSvgPictureWithoutAPlaceholder() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();
            byte[] svg = Encoding.UTF8.GetBytes(
                "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 10 10\">" +
                "<rect x=\"0\" y=\"0\" width=\"10\" height=\"10\" fill=\"#D7263D\"/>" +
                "</svg>");
            slide.AddPicture(
                new MemoryStream(svg),
                ImagePartType.Svg,
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(40),
                PowerPointUnits.FromPoints(30));

            OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Png);

            Assert.DoesNotContain(
                result.Diagnostics,
                diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.FromRgb(0xD7, 0x26, 0x3D), image!.GetPixel(40, 35));
        }

        [Fact]
        public void PowerPointSlide_RendersGeneratedMediaPosterThroughSharedDrawingImageElement() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();
            using MemoryStream video = new(new byte[] { 0, 0, 0, 24, 102, 116, 121, 112, 109, 112, 52, 50 });
            slide.AddVideo(
                video,
                "video/mp4",
                ".mp4",
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(80),
                PowerPointUnits.FromPoints(45));

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            OfficeDrawingImage poster = Assert.Single(snapshot.Drawing.Images);
            Assert.Equal("image/png", poster.ContentType);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(OfficeColor.FromRgb(31, 41, 55), rendered!.GetPixel(24, 24));
            Assert.True(CountPixelsNear(rendered, OfficeColor.FromRgb(249, 250, 251)) > 0);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<image", svgText, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_RendersBmpPicturesThroughSharedRasterDecoder() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 80);
            PowerPointSlide slide = presentation.AddSlide();
            byte[] bmp = CreateBmp24(2, 2, new[] {
                OfficeColor.FromRgb(18, 52, 86), OfficeColor.FromRgb(18, 52, 86),
                OfficeColor.FromRgb(18, 52, 86), OfficeColor.FromRgb(18, 52, 86)
            });
            slide.AddPicture(
                new MemoryStream(bmp),
                ImagePartType.Bmp,
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(24),
                PowerPointUnits.FromPoints(18));

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            Assert.Equal("image/bmp", drawingImage.ContentType);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(OfficeColor.FromRgb(18, 52, 86), rendered!.GetPixel(30, 28));
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_RendersTopDownBmpPicturesThroughSharedRasterDecoder() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 80);
            PowerPointSlide slide = presentation.AddSlide();
            byte[] bmp = CreateBmp24(2, 2, new[] {
                OfficeColor.FromRgb(24, 96, 144), OfficeColor.FromRgb(24, 96, 144),
                OfficeColor.FromRgb(24, 96, 144), OfficeColor.FromRgb(24, 96, 144)
            }, topDown: true);
            slide.AddPicture(
                new MemoryStream(bmp),
                ImagePartType.Bmp,
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(24),
                PowerPointUnits.FromPoints(18));

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.Equal("image/bmp", Assert.Single(snapshot.Drawing.Images).ContentType);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(OfficeColor.FromRgb(24, 96, 144), rendered!.GetPixel(30, 28));
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_RendersBmp32AlphaPicturesThroughSharedRasterDecoder() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 80);
            PowerPointSlide slide = presentation.AddSlide();
            byte[] bmp = CreateBmp32(2, 2, new[] {
                OfficeColor.FromRgba(255, 0, 0, 128), OfficeColor.FromRgba(255, 0, 0, 128),
                OfficeColor.FromRgba(255, 0, 0, 128), OfficeColor.FromRgba(255, 0, 0, 128)
            });
            slide.AddPicture(
                new MemoryStream(bmp),
                ImagePartType.Bmp,
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(24),
                PowerPointUnits.FromPoints(18));

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions { BackgroundColor = OfficeColor.White });
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg, new PowerPointImageExportOptions { BackgroundColor = OfficeColor.White });
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            Assert.Equal("image/bmp", drawingImage.ContentType);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            OfficeColor blended = rendered!.GetPixel(30, 28);
            Assert.True(blended.R >= 252, $"Expected red channel to stay near full after BMP alpha blend, got {blended.R}.");
            Assert.InRange(blended.G, 124, 130);
            Assert.InRange(blended.B, 124, 130);
            Assert.Equal(255, blended.A);
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_RendersGifPicturesThroughSharedRasterDecoder() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 80);
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "000000";
            slide.AddPicture(
                new MemoryStream(CreateSinglePixelGif()),
                ImagePartType.Gif,
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(24),
                PowerPointUnits.FromPoints(18));

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            OfficeDrawingImage drawingImage = Assert.Single(snapshot.Drawing.Images);
            Assert.Equal("image/gif", drawingImage.ContentType);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(OfficeColor.White, rendered!.GetPixel(30, 28));
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("data:image/gif;base64,", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsPictureCropAndTransformThroughSharedDrawingProjection() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 80);
            PowerPointSlide slide = presentation.AddSlide();

            OfficeRasterImage source = new(4, 4, OfficeColor.Transparent);
            for (int y = 0; y < source.Height; y++) {
                for (int x = 0; x < source.Width; x++) {
                    source.SetPixel(x, y, x < 2 ? OfficeColor.FromRgb(220, 38, 38) : OfficeColor.FromRgb(37, 99, 235));
                }
            }

            byte[] png = OfficePngWriter.Encode(source);
            PowerPointPicture picture = slide.AddPicture(
                new MemoryStream(png),
                ImagePartType.Png,
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(10),
                PowerPointUnits.FromPoints(80),
                PowerPointUnits.FromPoints(40));
            picture.Crop(leftPercent: 25, topPercent: 10, rightPercent: 25, bottomPercent: 20);
            picture.Rotation = 15D;
            picture.HorizontalFlip = true;
            picture.VerticalFlip = true;

            OfficeImageExportResult pngResult = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svgResult = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(pngResult.Diagnostics);
            AssertNoUnexpectedDiagnostics(svgResult.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);

            OfficeDrawingImage image = Assert.Single(snapshot.Drawing.Images);
            OfficeImageProjection projection = image.Projection;
            Assert.Equal("image/png", image.ContentType);
            Assert.Equal(20D, projection.X, 6);
            Assert.Equal(10D, projection.Y, 6);
            Assert.Equal(80D, projection.Width, 6);
            Assert.Equal(40D, projection.Height, 6);
            Assert.True(projection.HasCrop);
            Assert.Equal(0.25D, projection.SourceCrop.Left, 6);
            Assert.Equal(0.10D, projection.SourceCrop.Top, 6);
            Assert.Equal(0.25D, projection.SourceCrop.Right, 6);
            Assert.Equal(0.20D, projection.SourceCrop.Bottom, 6);
            Assert.Equal(0.50D, projection.SourceWidth, 6);
            Assert.Equal(0.70D, projection.SourceHeight, 6);
            Assert.True(projection.HasTransform);
            Assert.Equal(15D, projection.RotationDegrees, 6);
            Assert.Equal(60D, projection.RotationCenterX, 6);
            Assert.Equal(30D, projection.RotationCenterY, 6);
            Assert.True(projection.FlipHorizontal);
            Assert.True(projection.FlipVertical);

            Assert.True(OfficePngReader.TryDecode(pngResult.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(120, rendered!.Width);
            Assert.Equal(80, rendered.Height);

            string svgText = Encoding.UTF8.GetString(svgResult.Bytes);
            Assert.Contains("<image", svgText, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
            Assert.Contains("officeimo-image-clip-", svgText, StringComparison.Ordinal);
            Assert.Contains("rotate(15)", svgText, StringComparison.Ordinal);
            Assert.Contains("scale(-1 -1)", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_RendersImageBackgroundThroughSharedDrawingImageElement() {
            string imagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".png");
            try {
                OfficeColor backgroundImageColor = OfficeColor.FromRgb(60, 179, 113);
                byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(4, 4, backgroundImageColor));
                File.WriteAllBytes(imagePath, png);

                using var stream = new MemoryStream();
                using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
                presentation.SlideSize.SetSizePoints(80, 60);
                PowerPointSlide slide = presentation.AddSlide();
                slide.SetBackgroundImage(imagePath);

                OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions { IncludeSlideContent = false });
                PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(new PowerPointImageExportOptions { IncludeSlideContent = false });

                AssertNoUnexpectedDiagnostics(result.Diagnostics);
                Assert.Single(snapshot.Drawing.Images);
                Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
                Assert.Equal(backgroundImageColor, image!.GetPixel(5, 5));
            } finally {
                if (File.Exists(imagePath)) {
                    File.Delete(imagePath);
                }
            }
        }

        [Fact]
        public void PowerPointSlide_RendersTableCellsThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 120);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTablePoints(2, 2, 20, 20, 120, 50);
            table.SetColumnWidthsPoints(45, 75);
            table.SetRowHeightsPoints(20, 30);
            table.Rotation = 5D;

            PowerPointTableCell header = table.GetCell(0, 0);
            header.Text = "Region";
            header.FillColor = "22AA66";
            header.BorderColor = "114433";
            header.Color = "FFFFFF";
            header.Bold = true;
            header.FontSize = 11;
            header.HorizontalAlignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center;
            header.VerticalAlignment = DocumentFormat.OpenXml.Drawing.TextAnchoringTypeValues.Center;

            PowerPointTableCell merged = table.GetCell(1, 0);
            merged.Text = "Shared table renderer";
            merged.FillColor = "EAF7F0";
            merged.BorderColor = "114433";
            merged.Merge = (1, 2);
            table.GetCell(1, 1).Text = string.Empty;

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            Assert.True(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>().Count() >= 3);
            Assert.Contains(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), shape =>
                shape.Shape.FillColor == OfficeColor.FromRgb(34, 170, 102));
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingShape drawingShape && drawingShape.Shape.Transform.HasValue);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText drawingText && drawingText.Text == "Region" && Math.Abs(drawingText.RotationDegrees - 5D) < 0.000001D);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText drawingText && drawingText.Text == "Shared table renderer");

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);
            Assert.Equal(png.Height, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<rect", svgText, StringComparison.Ordinal);
            Assert.Contains("matrix(", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTableStyleFillsThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 120);
            presentation.SetThemeColor(PowerPointThemeColor.Accent1, "336699");
            presentation.SetThemeColor(PowerPointThemeColor.Light1, "F8FAFC");
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTablePoints(3, 1, 20, 18, 100, 72);
            table.SetRowHeightsPoints(24, 24, 24);
            table.FirstRow = true;
            table.BandedRows = true;
            table.GetCell(0, 0).Text = "Header";
            table.GetCell(1, 0).Text = "Band";
            table.GetCell(2, 0).Text = "Whole";

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            List<OfficeDrawingShape> rectangles = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(shape => shape.Shape.Kind == OfficeShapeKind.Rectangle)
                .ToList();
            Assert.Contains(rectangles, shape => shape.Shape.FillColor == OfficeColor.FromRgb(51, 102, 153));
            Assert.Contains(rectangles, shape => shape.Shape.FillColor == OfficeColor.FromRgb(173, 194, 214));
            Assert.Contains(rectangles, shape => shape.Shape.FillColor == OfficeColor.FromRgb(214, 224, 235));
            Assert.Contains(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), shape =>
                shape.Shape.Kind == OfficeShapeKind.Line &&
                shape.Shape.StrokeColor == OfficeColor.FromRgb(248, 250, 252) &&
                Math.Abs(shape.Shape.StrokeWidth - 3D) < 0.000001D);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("#336699", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#ADC2D6", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#D6E0EB", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#F8FAFC", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.FromRgb(51, 102, 153), image!.GetPixel(30, 30));
        }

        [Fact]
        public void PowerPointSlide_ProjectsTableCellBordersThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTable table = slide.AddTablePoints(1, 1, 20, 18, 100, 42);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Edges";
            cell.FillColor = "F8FAFC";
            cell.SetBorders(TableCellBorders.Left, "DC2626", 2D);
            cell.SetBorders(TableCellBorders.Top, "2563EB", 1.5D, A.PresetLineDashValues.Dash);
            cell.SetBorders(TableCellBorders.Right, "16A34A", 3D, A.PresetLineDashValues.Dot);
            cell.SetBorders(TableCellBorders.Bottom, "9333EA", 2.5D, A.PresetLineDashValues.DashDot);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            List<OfficeDrawingShape> borderLines = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line)
                .ToList();
            Assert.Contains(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(220, 38, 38) && line.Shape.StrokeWidth == 2D);
            Assert.Contains(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(37, 99, 235) && line.Shape.StrokeWidth == 1.5D && line.Shape.StrokeDashStyle == OfficeStrokeDashStyle.Dash);
            Assert.Contains(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(22, 163, 74) && line.Shape.StrokeWidth == 3D && line.Shape.StrokeDashStyle == OfficeStrokeDashStyle.Dot);
            Assert.Contains(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(147, 51, 234) && line.Shape.StrokeWidth == 2.5D && line.Shape.StrokeDashStyle == OfficeStrokeDashStyle.DashDot);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Edges", svgText, StringComparison.Ordinal);
            Assert.Contains("#DC2626", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#2563EB", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#16A34A", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#9333EA", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-dasharray", svgText, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);
        }

        [Fact]
        public void PowerPointSlide_ProjectsThemeTableCellBordersThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            presentation.SetThemeColor(PowerPointThemeColor.Accent2, "204060");
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTable table = slide.AddTablePoints(1, 1, 20, 18, 100, 42);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Theme edge";
            cell.FillColor = "FFFFFF";
            cell.SetBorders(TableCellBorders.Left, "000000", 2D);
            A.LinePropertiesType leftBorder = cell.Cell.TableCellProperties!.LeftBorderLineProperties!;
            leftBorder.RemoveAllChildren<A.SolidFill>();
            leftBorder.Append(new A.SolidFill(new A.SchemeColor(new A.LuminanceModulation { Val = 50000 }) { Val = A.SchemeColorValues.Accent2 }));

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            List<OfficeDrawingShape> borderLines = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line)
                .ToList();
            Assert.Contains(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(16, 32, 48) && line.Shape.StrokeWidth == 2D);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Theme edge", svgText, StringComparison.Ordinal);
            Assert.Contains("#102030", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTableCellFillAndBorderAlphaThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTable table = slide.AddTablePoints(1, 1, 20, 18, 100, 42);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Alpha cell";
            cell.FillColor = "FFFFFF";
            cell.SetBorders(TableCellBorders.Left, "000000", 2D);

            A.TableCellProperties properties = cell.Cell.TableCellProperties!;
            properties.RemoveAllChildren<A.SolidFill>();
            properties.Append(new A.SolidFill(
                new A.RgbColorModelHex(new A.Alpha { Val = 50000 }) { Val = "38BDF8" }));

            A.LinePropertiesType leftBorder = properties.LeftBorderLineProperties!;
            leftBorder.RemoveAllChildren<A.SolidFill>();
            leftBorder.Append(new A.SolidFill(
                new A.RgbColorModelHex(new A.Alpha { Val = 50000 }) { Val = "F97316" }));

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), shape =>
                shape.Shape.Kind == OfficeShapeKind.Rectangle &&
                shape.Shape.FillColor == OfficeColor.FromRgba(56, 189, 248, 128));
            Assert.Contains(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), shape =>
                shape.Shape.Kind == OfficeShapeKind.Line &&
                shape.Shape.StrokeColor == OfficeColor.FromRgba(249, 115, 22, 128) &&
                shape.Shape.StrokeWidth == 2D);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Alpha cell", svgText, StringComparison.Ordinal);
            Assert.Contains("fill=\"#38BDF8\" fill-opacity=\"0.502\"", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke=\"#F97316\" stroke-width=\"2\" stroke-opacity=\"0.502\"", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);
            Assert.Equal(png.Height, image.Height);
        }

        [Fact]
        public void PowerPointSlide_ProjectsDiagonalTableCellBordersThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTable table = slide.AddTablePoints(1, 1, 20, 18, 100, 42);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Diagonal";
            cell.FillColor = "FFFFFF";
            cell.SetBorders(TableCellBorders.DiagonalDown, "F97316", 2D);
            cell.SetBorders(TableCellBorders.DiagonalUp, "0EA5E9", 1.5D, A.PresetLineDashValues.Dash);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            List<OfficeDrawingShape> borderLines = snapshot.Drawing.Elements
                .OfType<OfficeDrawingShape>()
                .Where(shape => shape.Shape.Kind == OfficeShapeKind.Line)
                .ToList();
            OfficeDrawingShape down = Assert.Single(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(249, 115, 22));
            Assert.Equal(2D, down.Shape.StrokeWidth);
            Assert.Equal(new OfficePoint(0D, 0D), down.Shape.Points[0]);
            Assert.Equal(new OfficePoint(100D, 42D), down.Shape.Points[1]);

            OfficeDrawingShape up = Assert.Single(borderLines, line => line.Shape.StrokeColor == OfficeColor.FromRgb(14, 165, 233));
            Assert.Equal(1.5D, up.Shape.StrokeWidth);
            Assert.Equal(OfficeStrokeDashStyle.Dash, up.Shape.StrokeDashStyle);
            Assert.Equal(new OfficePoint(0D, 42D), up.Shape.Points[0]);
            Assert.Equal(new OfficePoint(100D, 0D), up.Shape.Points[1]);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Diagonal", svgText, StringComparison.Ordinal);
            Assert.Contains("#F97316", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#0EA5E9", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-dasharray", svgText, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);
        }

        [Fact]
        public void PowerPointSlide_ProjectsThemeTableCellFillsThroughSharedDrawingBorderBox() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            presentation.SetThemeColor(PowerPointThemeColor.Accent3, "336699");
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTable table = slide.AddTablePoints(1, 1, 20, 18, 100, 42);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Theme fill";
            cell.FillColor = "FFFFFF";
            A.TableCellProperties properties = cell.Cell.TableCellProperties!;
            properties.RemoveAllChildren<A.SolidFill>();
            properties.Append(new A.SolidFill(new A.SchemeColor(new A.Tint { Val = 50000 }) { Val = A.SchemeColorValues.Accent3 }));

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), shape =>
                shape.Shape.FillColor == OfficeColor.FromRgb(153, 178, 204) &&
                Math.Abs(shape.X - 20D) < 0.000001D &&
                Math.Abs(shape.Y - 18D) < 0.000001D);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Theme fill", svgText, StringComparison.Ordinal);
            Assert.Contains("#99B2CC", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(png.Width, image!.Width);
            Assert.Equal(png.Height, image.Height);
        }

        [Fact]
        public void PowerPointSlide_ProjectsThemeTableCellTextThroughSharedDrawingText() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            presentation.SetThemeColor(PowerPointThemeColor.Accent4, "884422");
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTable table = slide.AddTablePoints(1, 1, 20, 18, 100, 42);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Theme cell";
            A.Run run = cell.Cell.TextBody!.Elements<A.Paragraph>().Single().Elements<A.Run>().Single();
            run.RunProperties ??= new A.RunProperties();
            run.RunProperties.RemoveAllChildren<A.SolidFill>();
            run.RunProperties.InsertAt(new A.SolidFill(new A.SchemeColor(new A.Shade { Val = 50000 }) { Val = A.SchemeColorValues.Accent4 }), 0);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            OfficeDrawingText text = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), drawingText => drawingText.Text == "Theme cell");
            Assert.Equal(OfficeColor.FromRgb(68, 34, 17), text.Color);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Theme cell", svgText, StringComparison.Ordinal);
            Assert.Contains("#442211", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(160, image!.Width);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTableCellTextAlphaThroughSharedDrawingText() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTable table = slide.AddTablePoints(1, 1, 20, 18, 100, 42);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Alpha text";
            cell.FillColor = "FFFFFF";
            A.Run run = cell.Cell.TextBody!.Elements<A.Paragraph>().Single().Elements<A.Run>().Single();
            run.RunProperties ??= new A.RunProperties();
            run.RunProperties.RemoveAllChildren<A.SolidFill>();
            run.RunProperties.InsertAt(new A.SolidFill(
                new A.RgbColorModelHex(new A.Alpha { Val = 50000 }) { Val = "111827" }), 0);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            OfficeDrawingText text = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), drawingText => drawingText.Text == "Alpha text");
            Assert.Equal(OfficeColor.FromRgba(17, 24, 39, 128), text.Color);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Alpha text", svgText, StringComparison.Ordinal);
            Assert.Contains("fill=\"#111827\" fill-opacity=\"0.502\"", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(160, image!.Width);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTableCellPaddingThroughSharedDrawingPadding() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTablePoints(1, 1, 20, 18, 100, 40);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Padded table cell";
            cell.PaddingLeftPoints = 12D;
            cell.PaddingTopPoints = 6D;
            cell.PaddingRightPoints = 10D;
            cell.PaddingBottomPoints = 4D;

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);

            OfficeDrawingText text = snapshot.Drawing.Elements.OfType<OfficeDrawingText>().Single(drawingText => drawingText.Text == "Padded table cell");
            Assert.Equal(20D, text.X);
            Assert.Equal(18D, text.Y);
            Assert.Equal(100D, text.Width);
            Assert.Equal(40D, text.Height);
            Assert.True(text.HasPadding);
            Assert.Equal(12D, text.Padding.Left);
            Assert.Equal(6D, text.Padding.Top);
            Assert.Equal(10D, text.Padding.Right);
            Assert.Equal(4D, text.Padding.Bottom);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Padded", svgText, StringComparison.Ordinal);
            Assert.Contains("table", svgText, StringComparison.Ordinal);
            Assert.Contains("cell", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTableCellRichTextRunsThroughSharedDrawingRichText() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 110);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTablePoints(1, 1, 24, 24, 132, 46);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Plain ";
            cell.PaddingLeftPoints = 9D;
            cell.PaddingTopPoints = 5D;
            cell.AddRun("Red", run => {
                run.Color = "FF0000";
                run.Bold = true;
                run.FontSize = 13;
            });
            cell.AddRun(" blue", run => {
                run.Color = "0000FF";
                run.Italic = true;
                run.Underline = true;
                run.Strikethrough = true;
                run.FontName = "Aptos";
            });

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            Assert.Equal("Plain Red blue", richText.PlainText);
            Assert.Equal(3, richText.Runs.Count);
            Assert.Equal(24D, richText.X);
            Assert.Equal(24D, richText.Y);
            Assert.Equal(132D, richText.Width);
            Assert.Equal(46D, richText.Height);
            Assert.True(richText.HasPadding);
            Assert.Equal(9D, richText.Padding.Left);
            Assert.Equal(5D, richText.Padding.Top);
            Assert.True(richText.Runs[1].Bold);
            Assert.Equal(OfficeColor.Red, richText.Runs[1].Color);
            Assert.True(richText.Runs[2].Italic);
            Assert.True(richText.Runs[2].Underline);
            Assert.True(richText.Runs[2].Strikethrough);
            Assert.Equal(OfficeColor.Blue, richText.Runs[2].Color);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(180, image!.Width);
            Assert.Equal(110, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Plain", svgText, StringComparison.Ordinal);
            Assert.Contains("Red", svgText, StringComparison.Ordinal);
            Assert.Contains("blue", svgText, StringComparison.Ordinal);
            Assert.Contains("#FF0000", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#0000FF", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTableCellRunHighlightThroughSharedDrawingRichTextBackground() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 110);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTablePoints(1, 1, 24, 24, 132, 46);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Marked";
            cell.FontSize = 18;
            cell.Color = "111111";
            cell.Runs[0].HighlightColor = "C7F9CC";

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            OfficeRichTextRun run = Assert.Single(richText.Runs);
            Assert.Equal("Marked", run.Text);
            Assert.Equal(OfficeColor.FromRgb(199, 249, 204), run.BackgroundColor);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Marked", svgText, StringComparison.Ordinal);
            Assert.Contains("<rect", svgText, StringComparison.Ordinal);
            Assert.Contains("#C7F9CC", svgText, StringComparison.OrdinalIgnoreCase);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(199, 249, 204)) > 20, "Expected highlighted PowerPoint table-cell run background to render through the shared raster rich-text path.");
        }

        [Fact]
        public void PowerPointSlide_ProjectsTableCellRunHighlightAlphaThroughSharedDrawingRichTextBackground() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 110);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTablePoints(1, 1, 24, 24, 132, 46);
            PowerPointTableCell cell = table.GetCell(0, 0);
            cell.Text = "Marked";
            cell.FontSize = 18;
            cell.Color = "111111";
            cell.Runs[0].HighlightColor = "C7F9CC";
            A.Run run = cell.Cell.TextBody!.Elements<A.Paragraph>().Single().Elements<A.Run>().Single();
            run.RunProperties!.GetFirstChild<A.Highlight>()!
                .GetFirstChild<A.RgbColorModelHex>()!
                .Append(new A.Alpha { Val = 50000 });

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            OfficeRichTextRun richRun = Assert.Single(richText.Runs);
            Assert.Equal("Marked", richRun.Text);
            Assert.Equal(OfficeColor.FromRgba(199, 249, 204, 128), richRun.BackgroundColor);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Marked", svgText, StringComparison.Ordinal);
            Assert.Contains("#C7F9CC", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("fill-opacity=\"0.502\"", svgText, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(180, image!.Width);
            Assert.Equal(110, image.Height);
        }

        [Fact]
        public void PowerPointSlide_ProjectsFlippedTablesThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 120);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTable table = slide.AddTablePoints(2, 2, 25, 25, 110, 48);
            table.SetColumnWidthsPoints(50, 60);
            table.SetRowHeightsPoints(22, 26);
            table.HorizontalFlip = true;

            PowerPointTableCell header = table.GetCell(0, 0);
            header.Text = "Mirrored";
            header.FillColor = "22AA66";
            header.Color = "FFFFFF";
            header.Bold = true;

            PowerPointTableCell body = table.GetCell(1, 0);
            body.Text = "Shared table renderer";
            body.FillColor = "EAF7F0";
            body.Merge = (1, 2);
            table.GetCell(1, 1).Text = string.Empty;

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingShape drawingShape && drawingShape.Shape.Transform.HasValue);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText drawingText && drawingText.Text == "Mirrored" && drawingText.FlipHorizontal && !drawingText.FlipVertical);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText drawingText && drawingText.Text == "Shared table renderer" && drawingText.FlipHorizontal && !drawingText.FlipVertical);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("scale(-1 1)", svgText, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(180, image!.Width);
            Assert.Equal(120, image.Height);
        }

        [Fact]
        public void PowerPointSlide_RendersChartThroughSharedDrawingChartRenderer() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(360, 240);
            PowerPointSlide slide = presentation.AddSlide();
            var data = new PowerPointChartData(
                new[] { "Jan", "Feb", "Mar" },
                new[] {
                    new PowerPointChartSeries("Revenue", new[] { 10D, 18D, 24D }),
                    new PowerPointChartSeries("Forecast", new[] { 12D, 19D, 25D })
                });

            PowerPointChart chart = slide.AddChartPoints(data, 30, 25, 280, 180);
            chart.SetTitle("Revenue Trend");

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText drawingText && drawingText.Text == "Revenue Trend");
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText drawingText && drawingText.Text == "Revenue");
            Assert.True(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>().Count() > 5);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(360, image!.Width);
            Assert.Equal(240, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<rect", svgText, StringComparison.Ordinal);
            Assert.Contains("<text", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_SkipsChartFramesTooSmallForSafeRendering() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 80);
            PowerPointSlide slide = presentation.AddSlide();
            var data = new PowerPointChartData(
                new[] { "A", "B" },
                new[] { new PowerPointChartSeries("Series", new[] { 1D, 2D }) });
            slide.AddChartPoints(data, 10, 10, 1, 1);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            Assert.Contains(snapshot.Diagnostics, diagnostic =>
                diagnostic.Code == PowerPointImageExportDiagnosticCodes.UnsupportedShape);
        }

        [Fact]
        public void PowerPointSlide_RendersComboChartsThroughSharedDrawingChartRenderer() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(360, 240);
            PowerPointSlide slide = presentation.AddSlide();
            var data = new PowerPointChartData(
                new[] { "Q1", "Q2", "Q3" },
                new[] {
                    new PowerPointChartSeries("Bars", new[] { 10D, 16D, 22D }),
                    new PowerPointChartSeries("Trend", new[] { 12D, 18D, 24D })
                });

            PowerPointChart chart = slide.AddChartPoints(data, 30, 25, 280, 180);
            chart.SetTitle("Combo Trend");
            chart.SetSeriesLineColor(1, "DC2626", widthPoints: 2.5D);
            ConvertSecondBarSeriesToLineChart(chart);

            Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot snapshot));
            Assert.Equal(PowerPointChartSnapshotKind.ClusteredColumn, snapshot.ChartKind);
            Assert.Equal(PowerPointChartSnapshotKind.ClusteredColumn, snapshot.Data.Series[0].ChartKind);
            Assert.Equal(PowerPointChartSnapshotKind.Line, snapshot.Data.Series[1].ChartKind);

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(360, image!.Width);
            Assert.Equal(240, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("#DC2626", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<line", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_RejectsComboChartsWithScatterSeriesUntilNumericAxisIsModeled() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(360, 240);
            PowerPointSlide slide = presentation.AddSlide();
            var data = new PowerPointChartData(
                new[] { "Q1", "Q2", "Q3" },
                new[] {
                    new PowerPointChartSeries("Columns", new[] { 10D, 16D, 22D }),
                    new PowerPointChartSeries("Scatter", new[] { 12D, 18D, 24D })
                });

            PowerPointChart chart = slide.AddChartPoints(data, 30, 25, 280, 180);
            ConvertSecondBarSeriesToScatterChart(chart);

            Assert.False(chart.TryGetSnapshot(out _));
        }

        [Fact]
        public void PowerPointSlide_RejectsHorizontalBarComboChartsWithLineSeriesUntilAxisMappingIsModeled() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(360, 240);
            PowerPointSlide slide = presentation.AddSlide();
            var data = new PowerPointChartData(
                new[] { "Q1", "Q2", "Q3" },
                new[] {
                    new PowerPointChartSeries("Bars", new[] { 10D, 16D, 22D }),
                    new PowerPointChartSeries("Trend", new[] { 12D, 18D, 24D })
                });

            PowerPointChart chart = slide.AddChartPoints(data, 30, 25, 280, 180);
            SetBarChartShape(chart, C.BarDirectionValues.Bar, C.BarGroupingValues.Clustered);
            ConvertSecondBarSeriesToLineChart(chart);

            Assert.False(chart.TryGetSnapshot(out _));
        }

        [Fact]
        public void PowerPointSlide_RendersThemeChartSeriesColorsThroughSharedDrawingChartRenderer() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(360, 240);
            presentation.SetThemeColor(PowerPointThemeColor.Accent2, "884422");
            PowerPointSlide slide = presentation.AddSlide();
            var data = new PowerPointChartData(
                new[] { "Q1", "Q2", "Q3" },
                new[] {
                    new PowerPointChartSeries("Theme Trend", new[] { 12D, 18D, 24D })
                });

            PowerPointChart chart = slide.AddLineChartPoints(data, 30, 25, 280, 180);
            chart.SetSeriesLineColor(0, "000000", widthPoints: 2.5D);
            SetFirstLineSeriesOutlineSchemeColor(chart, A.SchemeColorValues.Accent2);

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);

            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(360, image!.Width);
            Assert.Equal(240, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("#884422", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<line", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTransformedChartsThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(360, 240);
            PowerPointSlide slide = presentation.AddSlide();
            var data = new PowerPointChartData(
                new[] { "Jan", "Feb", "Mar" },
                new[] {
                    new PowerPointChartSeries("Revenue", new[] { 10D, 18D, 24D }),
                    new PowerPointChartSeries("Forecast", new[] { 12D, 19D, 25D })
                });

            PowerPointChart chart = slide.AddChartPoints(data, 30, 25, 280, 180);
            chart.SetTitle("Revenue Trend");
            chart.HorizontalFlip = true;

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingShape drawingShape && drawingShape.Shape.Transform.HasValue);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText drawingText && drawingText.Text == "Revenue Trend" && drawingText.FlipHorizontal);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("scale(-1 1)", svgText, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(360, image!.Width);
            Assert.Equal(240, image.Height);
        }

        [Fact]
        public void PowerPointSlide_ProjectsGroupedScaledChartsAtMappedFrameSize() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(360, 240);
            PowerPointSlide slide = presentation.AddSlide();
            var data = new PowerPointChartData(
                new[] { "Jan", "Feb", "Mar" },
                new[] {
                    new PowerPointChartSeries("Revenue", new[] { 10D, 18D, 24D }),
                    new PowerPointChartSeries("Forecast", new[] { 12D, 19D, 25D })
                });

            PowerPointChart chart = slide.AddChartPoints(data, 40, 30, 280, 180);
            chart.SetTitle("Grouped Revenue");
            PowerPointAutoShape marker = slide.AddRectanglePoints(325, 210, 8, 8);
            marker.FillColor = "FFFFFF";
            marker.OutlineColor = "FFFFFF";
            slide.GroupShapes(new PowerPointShape[] { chart, marker }, "Scaled chart group");
            GroupShape group = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<GroupShape>()
                .Single();
            A.TransformGroup transform = group.GroupShapeProperties!.TransformGroup!;
            transform.Extents!.Cx = PowerPointUnits.FromPoints(140);
            transform.Extents.Cy = PowerPointUnits.FromPoints(90);
            transform.ChildExtents!.Cx = PowerPointUnits.FromPoints(280);
            transform.ChildExtents.Cy = PowerPointUnits.FromPoints(180);
            slide.SlidePart.Slide.Save();

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);

            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            OfficeDrawingGroup chartGroup = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingGroup>());
            Assert.Equal(40D, chartGroup.X, 1);
            Assert.Equal(30D, chartGroup.Y, 1);
            Assert.Equal(140D, chartGroup.ClipPath.Width, 1);
            Assert.Equal(90D, chartGroup.ClipPath.Height, 1);
            Assert.Contains(chartGroup.Drawing.Elements, element => element is OfficeDrawingText drawingText && drawingText.Text == "Grouped Revenue");
        }

        [Fact]
        public void PowerPointSlide_ProjectsSupportedTransformsThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointAutoShape rectangle = slide.AddRectanglePoints(40, 30, 40, 20);
            rectangle.FillColor = "22AA66";
            rectangle.Rotation = 15D;
            rectangle.HorizontalFlip = true;

            PowerPointTextBox textBox = slide.AddTextBoxPoints("Rotated", 70, 42, 60, 24);
            textBox.FontSize = 12;
            textBox.Color = "111111";
            textBox.Rotation = 10D;

            OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(result.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingShape drawingShape && drawingShape.Shape.Transform.HasValue);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText drawingText && Math.Abs(drawingText.RotationDegrees - 10D) < 0.000001D);

            string svgText = Encoding.UTF8.GetString(result.Bytes);
            Assert.Contains("matrix(", svgText, StringComparison.Ordinal);
            Assert.Contains("rotate(10", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsFlippedTextThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(160, 100);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints("Mirrored", 40, 30, 80, 24);
            textBox.FontSize = 14;
            textBox.Color = "111111";
            textBox.HorizontalFlip = true;

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingText drawingText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingText>());
            Assert.True(drawingText.FlipHorizontal);
            Assert.False(drawingText.FlipVertical);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("scale(-1 1)", svgText, StringComparison.Ordinal);
            Assert.Contains("Mirrored", svgText, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.Equal(160, rendered!.Width);
            Assert.Equal(100, rendered.Height);
        }

        [Fact]
        public void PowerPointSlide_ProjectsRichTextRunsThroughSharedDrawingRichText() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 110);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints("Plain ", 24, 24, 132, 58);
            PowerPointParagraph paragraph = textBox.Paragraphs[0];
            paragraph.AddRun("Red", run => {
                run.Color = "FF0000";
                run.Bold = true;
                run.FontSize = 14;
            });
            paragraph.AddRun(" blue", run => {
                run.Color = "0000FF";
                run.Italic = true;
                run.Underline = true;
                run.Strikethrough = true;
                run.FontName = "Aptos";
            });

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            Assert.Equal("Plain Red blue", richText.PlainText);
            Assert.Equal(3, richText.Runs.Count);
            Assert.True(richText.Runs[1].Bold);
            Assert.Equal(OfficeColor.Red, richText.Runs[1].Color);
            Assert.True(richText.Runs[2].Italic);
            Assert.True(richText.Runs[2].Underline);
            Assert.True(richText.Runs[2].Strikethrough);
            Assert.Equal(OfficeColor.Blue, richText.Runs[2].Color);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(180, image!.Width);
            Assert.Equal(110, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Plain", svgText, StringComparison.Ordinal);
            Assert.Contains("Red", svgText, StringComparison.Ordinal);
            Assert.Contains("blue", svgText, StringComparison.Ordinal);
            Assert.Contains("#FF0000", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#0000FF", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsRichTextRunColorAlphaThroughSharedDrawingRichText() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 110);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints("Plain ", 24, 24, 132, 58);
            PowerPointParagraph paragraph = textBox.Paragraphs[0];
            paragraph.AddRun("Red", run => {
                run.Color = "FF0000";
                run.Bold = true;
                run.FontSize = 14;
            });
            paragraph.AddRun(" blue", run => {
                run.Color = "0000FF";
                run.Italic = true;
            });
            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<Shape>()
                .Last();
            A.Run redRun = shape.TextBody!.Elements<A.Paragraph>().Single().Elements<A.Run>().ElementAt(1);
            redRun.RunProperties!.RemoveAllChildren<A.SolidFill>();
            redRun.RunProperties.InsertAt(new A.SolidFill(
                new A.RgbColorModelHex(new A.Alpha { Val = 50000 }) { Val = "FF0000" }), 0);

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            Assert.Equal("Plain Red blue", richText.PlainText);
            Assert.Equal(OfficeColor.FromRgba(255, 0, 0, 128), richText.Runs[1].Color);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Red", svgText, StringComparison.Ordinal);
            Assert.Contains("fill=\"#FF0000\" fill-opacity=\"0.502\"", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(180, image!.Width);
            Assert.Equal(110, image.Height);
        }

        [Fact]
        public void PowerPointSlide_ProjectsTextBoxRunHighlightThroughSharedDrawingRichTextBackground() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 110);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints("Marked", 24, 24, 132, 40);
            textBox.FontSize = 18;
            textBox.Color = "111111";
            textBox.Paragraphs[0].Runs[0].HighlightColor = "FFE680";

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            OfficeRichTextRun run = Assert.Single(richText.Runs);
            Assert.Equal("Marked", run.Text);
            Assert.Equal(OfficeColor.FromRgb(255, 230, 128), run.BackgroundColor);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Marked", svgText, StringComparison.Ordinal);
            Assert.Contains("<rect", svgText, StringComparison.Ordinal);
            Assert.Contains("#FFE680", svgText, StringComparison.OrdinalIgnoreCase);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(CountPixelsNear(image!, OfficeColor.FromRgb(255, 230, 128)) > 20, "Expected highlighted PowerPoint text-box run background to render through the shared raster rich-text path.");
        }

        [Fact]
        public void PowerPointSlide_ProjectsTextBoxRunHighlightAlphaThroughSharedDrawingRichTextBackground() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 110);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints("Marked", 24, 24, 132, 40);
            textBox.FontSize = 18;
            textBox.Color = "111111";
            textBox.Paragraphs[0].Runs[0].HighlightColor = "FFE680";
            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<Shape>()
                .Last();
            A.Run run = shape.TextBody!.Elements<A.Paragraph>().Single().Elements<A.Run>().Single();
            run.RunProperties!.GetFirstChild<A.Highlight>()!
                .GetFirstChild<A.RgbColorModelHex>()!
                .Append(new A.Alpha { Val = 50000 });

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            OfficeRichTextRun richRun = Assert.Single(richText.Runs);
            Assert.Equal("Marked", richRun.Text);
            Assert.Equal(OfficeColor.FromRgba(255, 230, 128, 128), richRun.BackgroundColor);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Marked", svgText, StringComparison.Ordinal);
            Assert.Contains("#FFE680", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("fill-opacity=\"0.502\"", svgText, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(180, image!.Width);
            Assert.Equal(110, image.Height);
        }

        [Fact]
        public void PowerPointSlide_ProjectsThemeTextBoxRunHighlightThroughSharedDrawingRichTextBackground() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 110);
            presentation.SetThemeColor(PowerPointThemeColor.Accent5, "66CCFF");
            PowerPointSlide slide = presentation.AddSlide();

            slide.AddTextBoxPoints("Theme mark", 24, 24, 132, 40);
            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<Shape>()
                .Last();
            A.Run themeRun = shape.TextBody!.Elements<A.Paragraph>().Single().Elements<A.Run>().Single();
            themeRun.RunProperties ??= new A.RunProperties();
            themeRun.RunProperties.RemoveAllChildren<A.Highlight>();
            themeRun.RunProperties.Append(new A.Highlight(
                new A.SchemeColor(new A.LuminanceModulation { Val = 50000 }) { Val = A.SchemeColorValues.Accent5 }));

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            OfficeRichTextRun run = Assert.Single(richText.Runs);
            Assert.Equal("Theme mark", run.Text);
            Assert.Equal(OfficeColor.FromRgb(0, 119, 178), run.BackgroundColor);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Theme", svgText, StringComparison.Ordinal);
            Assert.Contains("#0077B2", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsThemeRichTextRunsThroughSharedDrawingRichText() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 110);
            presentation.SetThemeColor(PowerPointThemeColor.Accent5, "2468AC");
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints("Plain ", 24, 24, 132, 58);
            textBox.Paragraphs[0].AddRun("Theme", run => {
                run.Bold = true;
                run.FontSize = 14;
            });
            Shape shape = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
                .Elements<Shape>()
                .Last();
            A.Run themeRun = shape.TextBody!.Elements<A.Paragraph>().Single().Elements<A.Run>().Last();
            themeRun.RunProperties ??= new A.RunProperties();
            themeRun.RunProperties.RemoveAllChildren<A.SolidFill>();
            themeRun.RunProperties.InsertAt(new A.SolidFill(new A.SchemeColor(new A.LuminanceModulation { Val = 50000 }) { Val = A.SchemeColorValues.Accent5 }), 0);

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            OfficeRichTextRun theme = Assert.Single(richText.Runs, run => run.Text == "Theme");
            Assert.Equal(OfficeColor.FromRgb(18, 52, 86), theme.Color);
            Assert.True(theme.Bold);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(180, image!.Width);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Theme", svgText, StringComparison.Ordinal);
            Assert.Contains("#123456", svgText, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PowerPointSlide_ProjectsSingleDecoratedTextRunThroughSharedDrawingRichText() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 110);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints("Deprecated", 24, 24, 132, 40);
            textBox.ApplyTextStyle(PowerPointTextStyle.Body.WithUnderline(true).WithStrikethrough(true), applyToRuns: true);

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            Assert.DoesNotContain(snapshot.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "Deprecated");
            Assert.Equal("Deprecated", richText.PlainText);
            OfficeRichTextRun run = Assert.Single(richText.Runs);
            Assert.True(run.Underline);
            Assert.True(run.Strikethrough);
            Assert.Equal(18D, run.FontSize);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(180, image!.Width);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Deprecated", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsMarkdownStrikethroughThroughSharedDrawingRichText() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(220, 120);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints(string.Empty, 24, 24, 172, 46);
            textBox.SetMarkdown("Keep ~~obsolete~~ current");

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            Assert.Equal("Keep obsolete current", richText.PlainText);
            Assert.Equal(3, richText.Runs.Count);
            Assert.False(richText.Runs[0].Strikethrough);
            Assert.True(richText.Runs[1].Strikethrough);
            Assert.False(richText.Runs[2].Strikethrough);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(220, image!.Width);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("obsolete", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsListMarkersThroughSharedDrawingRichText() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(320, 260);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox bullets = slide.AddTextBoxPoints(string.Empty, 20, 20, 280, 220);
            bullets.SetBullets(new[] { "First bullet", "Second bullet" });
            bullets.AddNumberedList(new[] { "First number", "Second number" }, startAt: 3);
            bullets.AddBullets(new[] { "Nested bullet" }, level: 1);

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingRichText[] richTexts = snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>().ToArray();
            Assert.Equal(5, richTexts.Length);
            Assert.Contains(richTexts[0].Runs, run => run.Text == "\u2022 ");
            Assert.Contains(richTexts[2].Runs, run => run.Text == "3. ");
            Assert.Contains(richTexts[3].Runs, run => run.Text == "4. ");
            Assert.Contains(richTexts[4].Runs, run => run.Text == "\u2022 ");
            Assert.Contains("First bullet", richTexts[0].PlainText, StringComparison.Ordinal);
            Assert.Contains("Second number", richTexts[3].PlainText, StringComparison.Ordinal);
            Assert.Contains("Nested bullet", richTexts[4].PlainText, StringComparison.Ordinal);
            Assert.Equal(0D, richTexts[0].ParagraphIndent.FirstLineOffset);
            Assert.Equal(18D, richTexts[0].ParagraphIndent.ContinuationLineOffset);
            Assert.Equal(18D, richTexts[4].ParagraphIndent.FirstLineOffset);
            Assert.Equal(36D, richTexts[4].ParagraphIndent.ContinuationLineOffset);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(320, image!.Width);
            Assert.Equal(260, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("First bullet", svgText, StringComparison.Ordinal);
            Assert.Contains("Second number", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsAutoNumberSchemesThroughSharedDrawingRichText() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(340, 260);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox alpha = slide.AddTextBoxPoints(string.Empty, 20, 20, 300, 90);
            alpha.AddNumberedList(
                new[] { "Alpha lower", "Alpha next" },
                A.TextAutoNumberSchemeValues.AlphaLowerCharacterPeriod,
                startAt: 2);

            PowerPointTextBox roman = slide.AddTextBoxPoints(string.Empty, 20, 130, 300, 90);
            roman.AddNumberedList(
                new[] { "Roman upper", "Roman next" },
                A.TextAutoNumberSchemeValues.RomanUpperCharacterParenR,
                startAt: 4);

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingRichText[] richTexts = snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>().ToArray();
            Assert.Equal(4, richTexts.Length);
            Assert.Contains(richTexts[0].Runs, run => run.Text == "b. ");
            Assert.Contains(richTexts[1].Runs, run => run.Text == "c. ");
            Assert.Contains(richTexts[2].Runs, run => run.Text == "IV) ");
            Assert.Contains(richTexts[3].Runs, run => run.Text == "V) ");
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(340, image!.Width);
            Assert.Equal(260, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("b.", svgText, StringComparison.Ordinal);
            Assert.Contains("IV)", svgText, StringComparison.Ordinal);
            Assert.Contains("Roman next", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_ProjectsFlippedRichTextRunsThroughSharedDrawingRichText() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(180, 110);
            PowerPointSlide slide = presentation.AddSlide();

            PowerPointTextBox textBox = slide.AddTextBoxPoints("Plain ", 24, 24, 132, 58);
            textBox.HorizontalFlip = true;
            PowerPointParagraph paragraph = textBox.Paragraphs[0];
            paragraph.AddRun("Red", run => {
                run.Color = "FF0000";
                run.Bold = true;
            });
            paragraph.AddRun(" blue", run => {
                run.Color = "0000FF";
                run.Italic = true;
            });

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            OfficeDrawingRichText richText = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingRichText>());
            Assert.True(richText.FlipHorizontal);
            Assert.False(richText.FlipVertical);
            Assert.Empty(snapshot.Drawing.Elements.OfType<OfficeDrawingText>());
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(180, image!.Width);
            Assert.Equal(110, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("scale(-1 1)", svgText, StringComparison.Ordinal);
            Assert.Contains("Plain", svgText, StringComparison.Ordinal);
            Assert.Contains("Red", svgText, StringComparison.Ordinal);
            Assert.Contains("blue", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointSlide_RendersInheritedLayoutTextThroughSharedDrawing() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 160);

            int layoutIndex = presentation.GetLayoutIndex(SlideLayoutValues.TitleOnly);
            var bounds = new PowerPointLayoutBox(
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(24),
                PowerPointUnits.FromPoints(180),
                PowerPointUnits.FromPoints(32));
            PowerPointTextBox layoutTitle = presentation.EnsureLayoutPlaceholderTextBox(0, layoutIndex, PlaceholderValues.Title, bounds: bounds);
            presentation.SetLayoutPlaceholderBounds(0, layoutIndex, PlaceholderValues.Title, bounds);
            layoutTitle.Text = "Inherited layout title";
            layoutTitle.FontSize = 14;
            layoutTitle.Color = "1F4E79";

            PowerPointSlide slide = presentation.AddSlide(SlideLayoutValues.TitleOnly);

            OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Png);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            Assert.Empty(slide.Shapes);
            AssertNoUnexpectedDiagnostics(result.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText text && text.Text == "Inherited layout title");
            Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
            Assert.Equal(240, image!.Width);
            Assert.Equal(160, image.Height);
        }

        [Fact]
        public void PowerPointSlide_SuppressesInheritedLayoutPlaceholderWhenSlidePlaceholderOverridesIt() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 160);

            int layoutIndex = presentation.GetLayoutIndex(SlideLayoutValues.TitleOnly);
            var bounds = new PowerPointLayoutBox(
                PowerPointUnits.FromPoints(20),
                PowerPointUnits.FromPoints(24),
                PowerPointUnits.FromPoints(180),
                PowerPointUnits.FromPoints(32));
            PowerPointTextBox layoutTitle = presentation.EnsureLayoutPlaceholderTextBox(0, layoutIndex, PlaceholderValues.Title, bounds: bounds);
            presentation.SetLayoutPlaceholderBounds(0, layoutIndex, PlaceholderValues.Title, bounds);
            layoutTitle.Text = "Inherited layout title";
            layoutTitle.Color = "C00000";

            PowerPointSlide slide = presentation.AddSlide(SlideLayoutValues.TitleOnly);
            PowerPointTextBox slideTitle = slide.AddTitlePoints("Slide local title", 20, 24, 180, 32);
            slideTitle.Color = "1F4E79";

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg);
            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot();

            Assert.Single(slide.Shapes);
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);
            AssertNoUnexpectedDiagnostics(snapshot.Diagnostics);
            Assert.Contains(snapshot.Drawing.Elements, element => element is OfficeDrawingText text && text.Text == "Slide local title");
            Assert.DoesNotContain(snapshot.Drawing.Elements, element => element is OfficeDrawingText text && text.Text == "Inherited layout title");
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(240, image!.Width);
            Assert.Equal(160, image.Height);

            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("Slide local title", svgText, StringComparison.Ordinal);
            Assert.DoesNotContain("Inherited layout title", svgText, StringComparison.Ordinal);
        }

        [Fact]
        public void PowerPointImageExportOptionsReuseSharedOfficeImageExportOptions() {
            PowerPointImageExportOptions options = new PowerPointImageExportOptions {
                Scale = 1.25D,
                BackgroundColor = OfficeColor.AliceBlue
            };

            Assert.IsAssignableFrom<OfficeImageExportOptions>(options);
            Assert.Equal(1.25D, options.Scale);
            Assert.Equal(OfficeColor.AliceBlue, options.BackgroundColor);

            PowerPointPresentationImageExportOptions presentationOptions = new PowerPointPresentationImageExportOptions {
                Scale = 1.5D,
                BackgroundColor = OfficeColor.White,
                IncludeHiddenSlides = true,
                SlideNumbers = new[] { 1, 3 }
            };

            Assert.IsAssignableFrom<OfficeImageExportOptions>(presentationOptions);
            Assert.Equal(1.5D, presentationOptions.Scale);
            Assert.Equal(OfficeColor.White, presentationOptions.BackgroundColor);
            Assert.True(presentationOptions.IncludeHiddenSlides);
            Assert.Equal(new[] { 1, 3 }, presentationOptions.SlideNumbers);

            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                presentation.AddSlide().ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions { Scale = 0D }));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                presentation.ExportImages(OfficeImageExportFormat.Png, new PowerPointPresentationImageExportOptions { SlideNumbers = new[] { 0 } }));
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                presentation.ExportImages(OfficeImageExportFormat.Png, new PowerPointPresentationImageExportOptions { SlideNumbers = new[] { 2 } }));
            Assert.Throws<ArgumentOutOfRangeException>(() => presentation.ToImages().ForSlides(0));
            Assert.Throws<ArgumentOutOfRangeException>(() => presentation.ToImages().ForSlideRange(2, 1));
        }

        [Fact]
        public void PresentationImageExportPropagatesAllSlideContentFilters() {
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddPicture(imagePath);
            PowerPointAutoShape shape = slide.AddRectanglePoints(12, 12, 30, 20, "Filtered shape");
            shape.FillColor = "ABCDEF";
            slide.AddTextBoxPoints("FILTERED TEXT", 48, 12, 100, 24);
            PowerPointTable table = slide.AddTablePoints(1, 1, 12, 44, 120, 30);
            table.GetCell(0, 0).Text = "FILTERED TABLE";
            slide.AddChartPoints(OfficeChartKind.ColumnClustered,
                new OfficeChartData(new[] { "FILTERED CATEGORY" }, new[] {
                    new OfficeChartSeries("FILTERED CHART", new[] { 42D })
                }), 150, 44, 120, 80);

            OfficeImageExportResult result = Assert.Single(presentation.ExportImages(
                OfficeImageExportFormat.Svg, new PowerPointPresentationImageExportOptions {
                    IncludeSlideBackground = false,
                    IncludePictures = false,
                    IncludeAutoShapes = false,
                    IncludeTextBoxes = false,
                    IncludeTables = false,
                    IncludeCharts = false
                }));
            string svg = Encoding.UTF8.GetString(result.Bytes);

            Assert.DoesNotContain("<image", svg, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("ABCDEF", svg, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("FILTERED TEXT", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("FILTERED TABLE", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("FILTERED CATEGORY", svg, StringComparison.Ordinal);
        }

        private static int CountPixelsNear(OfficeRasterImage image, OfficeColor expected) {
            int count = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor actual = image.GetPixel(x, y);
                    if (Math.Abs(actual.R - expected.R) <= 8 &&
                        Math.Abs(actual.G - expected.G) <= 8 &&
                        Math.Abs(actual.B - expected.B) <= 8 &&
                        Math.Abs(actual.A - expected.A) <= 8) {
                        count++;
                    }
                }
            }

            return count;
        }

        private static void ConvertSecondBarSeriesToLineChart(PowerPointChart chart) {
            DocumentFormat.OpenXml.Packaging.ChartPart chartPart = GetChartPart(chart);
            C.PlotArea plotArea = chartPart.ChartSpace!.Descendants<C.PlotArea>().Single();
            C.BarChart barChart = plotArea.Elements<C.BarChart>().Single();
            C.BarChartSeries barSeries = barChart.Elements<C.BarChartSeries>().Skip(1).Single();
            var lineChart = new C.LineChart(new C.Grouping { Val = C.GroupingValues.Standard });
            var lineSeries = new C.LineChartSeries();
            foreach (OpenXmlElement child in barSeries.ChildElements) {
                lineSeries.Append(child.CloneNode(true));
            }

            lineChart.Append(lineSeries);
            foreach (C.AxisId axisId in barChart.Elements<C.AxisId>()) {
                lineChart.Append((C.AxisId)axisId.CloneNode(true));
            }

            barSeries.Remove();
            barChart.InsertAfterSelf(lineChart);
            chartPart.ChartSpace.Save();
        }

        private static void SetBarChartShape(PowerPointChart chart, C.BarDirectionValues direction, C.BarGroupingValues grouping) {
            DocumentFormat.OpenXml.Packaging.ChartPart chartPart = GetChartPart(chart);
            C.BarChart barChart = chartPart.ChartSpace!.Descendants<C.BarChart>().Single();
            barChart.GetFirstChild<C.BarDirection>()!.Val = direction;
            barChart.GetFirstChild<C.BarGrouping>()!.Val = grouping;
            chartPart.ChartSpace.Save();
        }

        private static void ConvertSecondBarSeriesToScatterChart(PowerPointChart chart) {
            DocumentFormat.OpenXml.Packaging.ChartPart chartPart = GetChartPart(chart);
            C.PlotArea plotArea = chartPart.ChartSpace!.Descendants<C.PlotArea>().Single();
            C.BarChart barChart = plotArea.Elements<C.BarChart>().Single();
            C.BarChartSeries barSeries = barChart.Elements<C.BarChartSeries>().Skip(1).Single();
            var scatterChart = new C.ScatterChart(new C.ScatterStyle { Val = C.ScatterStyleValues.LineMarker });
            var scatterSeries = new C.ScatterChartSeries();
            C.Index? index = barSeries.GetFirstChild<C.Index>()?.CloneNode(true) as C.Index;
            C.Order? order = barSeries.GetFirstChild<C.Order>()?.CloneNode(true) as C.Order;
            C.SeriesText? text = barSeries.GetFirstChild<C.SeriesText>()?.CloneNode(true) as C.SeriesText;
            if (index != null) scatterSeries.Append(index);
            if (order != null) scatterSeries.Append(order);
            if (text != null) scatterSeries.Append(text);
            scatterSeries.Append(
                new C.XValues(CreateNumberReference("Sheet1!$D$2:$D$3", new[] { 1.5D, 2.5D })),
                new C.YValues(CreateNumberReference("Sheet1!$E$2:$E$3", new[] { 11D, 13D })));

            scatterChart.Append(scatterSeries);
            foreach (C.AxisId axisId in barChart.Elements<C.AxisId>()) {
                scatterChart.Append((C.AxisId)axisId.CloneNode(true));
            }

            barSeries.Remove();
            barChart.InsertAfterSelf(scatterChart);
            chartPart.ChartSpace.Save();
        }

        private static C.NumberReference CreateNumberReference(string formula, IReadOnlyList<double> values) {
            C.NumberingCache cache = new(new C.FormatCode { Text = "General" }, new C.PointCount { Val = (uint)values.Count });
            for (int i = 0; i < values.Count; i++) {
                cache.Append(new C.NumericPoint {
                    Index = (uint)i,
                    NumericValue = new C.NumericValue { Text = values[i].ToString(CultureInfo.InvariantCulture) }
                });
            }

            return new C.NumberReference(new C.Formula { Text = formula }, cache);
        }

        private static void SetFirstLineSeriesOutlineSchemeColor(PowerPointChart chart, A.SchemeColorValues schemeColor) {
            DocumentFormat.OpenXml.Packaging.ChartPart chartPart = GetChartPart(chart);
            C.LineChartSeries series = chartPart.ChartSpace!.Descendants<C.LineChartSeries>().First();
            C.ChartShapeProperties properties = series.GetFirstChild<C.ChartShapeProperties>() ?? series.AppendChild(new C.ChartShapeProperties());
            A.Outline outline = properties.GetFirstChild<A.Outline>() ?? properties.AppendChild(new A.Outline());
            outline.RemoveAllChildren<A.SolidFill>();
            outline.InsertAt(new A.SolidFill(new A.SchemeColor { Val = schemeColor }), 0);
            chartPart.ChartSpace.Save();
        }

        private static DocumentFormat.OpenXml.Packaging.ChartPart GetChartPart(PowerPointChart chart) {
            MethodInfo method = typeof(PowerPointChart).GetMethod("GetChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!;
            return (DocumentFormat.OpenXml.Packaging.ChartPart)method.Invoke(chart, Array.Empty<object>())!;
        }

        private static byte[] CreateBmp24(int width, int height, IReadOnlyList<OfficeColor> pixels, bool topDown = false) {
            int rowStride = ((width * 24) + 31) / 32 * 4;
            int pixelOffset = 54;
            byte[] bytes = new byte[pixelOffset + (rowStride * height)];
            bytes[0] = (byte)'B';
            bytes[1] = (byte)'M';
            WriteInt32LittleEndian(bytes, 2, bytes.Length);
            WriteInt32LittleEndian(bytes, 10, pixelOffset);
            WriteInt32LittleEndian(bytes, 14, 40);
            WriteInt32LittleEndian(bytes, 18, width);
            WriteInt32LittleEndian(bytes, 22, topDown ? -height : height);
            WriteUInt16LittleEndian(bytes, 26, 1);
            WriteUInt16LittleEndian(bytes, 28, 24);

            for (int y = 0; y < height; y++) {
                int sourceY = topDown ? y : height - 1 - y;
                int rowOffset = pixelOffset + (y * rowStride);
                for (int x = 0; x < width; x++) {
                    OfficeColor color = pixels[(sourceY * width) + x];
                    int offset = rowOffset + (x * 3);
                    bytes[offset] = color.B;
                    bytes[offset + 1] = color.G;
                    bytes[offset + 2] = color.R;
                }
            }

            return bytes;
        }

        private static byte[] CreateBmp32(int width, int height, IReadOnlyList<OfficeColor> pixels) {
            int rowStride = width * 4;
            int pixelOffset = 54;
            byte[] bytes = new byte[pixelOffset + (rowStride * height)];
            bytes[0] = (byte)'B';
            bytes[1] = (byte)'M';
            WriteInt32LittleEndian(bytes, 2, bytes.Length);
            WriteInt32LittleEndian(bytes, 10, pixelOffset);
            WriteInt32LittleEndian(bytes, 14, 40);
            WriteInt32LittleEndian(bytes, 18, width);
            WriteInt32LittleEndian(bytes, 22, height);
            WriteUInt16LittleEndian(bytes, 26, 1);
            WriteUInt16LittleEndian(bytes, 28, 32);

            for (int y = 0; y < height; y++) {
                int sourceY = height - 1 - y;
                int rowOffset = pixelOffset + (y * rowStride);
                for (int x = 0; x < width; x++) {
                    OfficeColor color = pixels[(sourceY * width) + x];
                    int offset = rowOffset + (x * 4);
                    bytes[offset] = color.B;
                    bytes[offset + 1] = color.G;
                    bytes[offset + 2] = color.R;
                    bytes[offset + 3] = color.A;
                }
            }

            return bytes;
        }

        private static byte[] CreateSinglePixelGif() =>
            Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");

        private static void WriteInt32LittleEndian(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
            bytes[offset + 2] = (byte)(value >> 16);
            bytes[offset + 3] = (byte)(value >> 24);
        }

        private static void WriteUInt16LittleEndian(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
        }

        private static ConnectionShape CreateNativeBentConnectionShape() {
            return new ConnectionShape(
                new NonVisualConnectionShapeProperties(
                    new NonVisualDrawingProperties { Id = 20U, Name = "Native Bent Connector" },
                    new NonVisualConnectorShapeDrawingProperties(
                        new A.StartConnection { Id = 2U, Index = 3U },
                        new A.EndConnection { Id = 3U, Index = 1U }),
                    new ApplicationNonVisualDrawingProperties()),
                new ShapeProperties(
                    new A.Transform2D(
                        new A.Offset {
                            X = PowerPointUnits.FromPoints(54),
                            Y = PowerPointUnits.FromPoints(50)
                        },
                        new A.Extents {
                            Cx = PowerPointUnits.FromPoints(96),
                            Cy = PowerPointUnits.FromPoints(64)
                        }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.BentConnector4 },
                    new A.Outline(
                        new A.SolidFill(new A.RgbColorModelHex { Val = "1E5A96" }),
                        new A.HeadEnd {
                            Type = A.LineEndValues.Diamond,
                            Width = A.LineEndWidthValues.Medium,
                            Length = A.LineEndLengthValues.Medium
                        },
                        new A.TailEnd {
                            Type = A.LineEndValues.Triangle,
                            Width = A.LineEndWidthValues.Medium,
                            Length = A.LineEndLengthValues.Medium
                        }) {
                        Width = 25400
                    }));
        }

        private static void AssertNoUnexpectedDiagnostics(IEnumerable<OfficeImageExportDiagnostic> diagnostics) {
            OfficeImageExportDiagnostic[] unexpected = diagnostics
                .Where(diagnostic =>
                    diagnostic.Code != OfficeImageExportDiagnosticCodes.FontSubstituted)
                .ToArray();
            Assert.Empty(unexpected);
        }
    }
}
