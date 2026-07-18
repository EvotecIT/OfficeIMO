using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        [Fact]
        public void PowerPointSlide_ProjectsOuterShadowThroughSharedDrawingImageExport() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 90);
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "FFFFFF";

            PowerPointAutoShape shape = slide.AddShapePoints(A.ShapeTypeValues.Rectangle, 20, 20, 60, 30);
            shape.FillColor = "E0F2FE";
            shape.OutlineColor = "0284C7";
            shape.SetShadow("000000", distancePoints: 8, angleDegrees: 45, transparencyPercent: 0);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(new PowerPointImageExportOptions { Scale = 1D });
            OfficeDrawingShape drawingShape = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), element =>
                Math.Abs(element.X - 20D) < 0.000001D &&
                Math.Abs(element.Y - 20D) < 0.000001D);
            Assert.NotNull(drawingShape.Shape.Shadow);
            Assert.Equal(OfficeColor.FromRgb(0, 0, 0), drawingShape.Shape.Shadow!.Color);
            Assert.Equal(1D, drawingShape.Shape.Shadow.Opacity, precision: 3);
            Assert.Equal(8D / Math.Sqrt(2D), drawingShape.Shape.Shadow.OffsetX, precision: 3);
            Assert.Equal(8D / Math.Sqrt(2D), drawingShape.Shape.Shadow.OffsetY, precision: 3);
            Assert.Equal(4D, drawingShape.Shape.Shadow.BlurRadius, precision: 3);

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg, new PowerPointImageExportOptions { Scale = 1D });
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("#000000", svgText, StringComparison.OrdinalIgnoreCase);
            Assert.True(CountOccurrences(svgText, "#000000", StringComparison.OrdinalIgnoreCase) >= 3, "Expected SVG output to layer the blurred shadow.");
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions { Scale = 1D });
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.FromRgb(0, 0, 0), image!.GetPixel(83, 53));
            Assert.True(ContainsVisibleNonWhitePixel(image!, 87, 35, 8, 14), "Expected blurred shadow halo pixels outside the hard shadow bounds.");
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
        }

        [Fact]
        public void PowerPointSlide_ProjectsGlowThroughSharedDrawingImageExport() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(120, 90);
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "FFFFFF";

            PowerPointAutoShape shape = slide.AddShapePoints(A.ShapeTypeValues.Rectangle, 32, 24, 48, 28);
            shape.FillColor = "E0F2FE";
            shape.OutlineColor = "0284C7";
            shape.SetGlow("FF00FF", radiusPoints: 6, transparencyPercent: 25);

            PowerPointSlideVisualSnapshot snapshot = slide.CreateVisualSnapshot(new PowerPointImageExportOptions { Scale = 1D });
            OfficeDrawingShape drawingShape = Assert.Single(snapshot.Drawing.Elements.OfType<OfficeDrawingShape>(), element =>
                Math.Abs(element.X - 32D) < 0.000001D &&
                Math.Abs(element.Y - 24D) < 0.000001D);
            Assert.NotNull(drawingShape.Shape.Glow);
            Assert.Equal(OfficeColor.FromRgb(255, 0, 255), drawingShape.Shape.Glow!.Color);
            Assert.Equal(191D / 255D, drawingShape.Shape.Glow.Opacity, precision: 3);
            Assert.Equal(6D, drawingShape.Shape.Glow.Radius, precision: 3);

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg, new PowerPointImageExportOptions { Scale = 1D });
            string svgText = Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("#FF00FF", svgText, StringComparison.OrdinalIgnoreCase);
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions { Scale = 1D });
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.True(ContainsMagentaGlowPixel(image!, 24, 16, 64, 44));
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
        }

        private static bool ContainsMagentaGlowPixel(OfficeRasterImage image, int left, int top, int width, int height) {
            int right = Math.Min(image.Width, left + width);
            int bottom = Math.Min(image.Height, top + height);
            for (int y = Math.Max(0, top); y < bottom; y++) {
                for (int x = Math.Max(0, left); x < right; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.R > 240 && pixel.B > 240 && pixel.G < 245) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool ContainsVisibleNonWhitePixel(OfficeRasterImage image, int left, int top, int width, int height) {
            int right = Math.Min(image.Width, left + width);
            int bottom = Math.Min(image.Height, top + height);
            for (int y = Math.Max(0, top); y < bottom; y++) {
                for (int x = Math.Max(0, left); x < right; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.A > 0 && (pixel.R < 245 || pixel.G < 245 || pixel.B < 245)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static int CountOccurrences(string text, string value, StringComparison comparison) {
            int count = 0;
            int index = 0;
            while ((index = text.IndexOf(value, index, comparison)) >= 0) {
                count++;
                index += value.Length;
            }

            return count;
        }
    }
}
