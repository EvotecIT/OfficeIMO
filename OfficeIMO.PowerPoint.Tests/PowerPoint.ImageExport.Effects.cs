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
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(120),OfficeIMO.OfficeLength.Points(90));
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "FFFFFF";

            PowerPointAutoShape shape = slide.AddShape(OfficePresetShapeType.Rectangle, OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(60), OfficeIMO.OfficeLength.Points(30));
            shape.FillColor = "E0F2FE";
            shape.OutlineColor = "0284C7";
            shape.SetShadow("000000", distancePoints: 16, angleDegrees: 45, transparencyPercent: 0);

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg, new PowerPointImageExportOptions { Scale = 1D });
            OfficeRasterImage svgImage = ReadSvgOutput(svg);
            Assert.Equal(OfficeColor.FromRgb(0, 0, 0), svgImage.GetPixel(83, 53));
            Assert.True(ContainsVisibleNonWhitePixel(svgImage, 93, 35, 8, 14), "Expected the shadow halo in SVG output.");
            AssertNoUnexpectedDiagnostics(svg.Diagnostics);

            OfficeImageExportResult png = slide.ExportImage(OfficeImageExportFormat.Png, new PowerPointImageExportOptions { Scale = 1D });
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            // The sample lies inside the opaque shadow, beyond the foreground shape
            // and the blur transition. Test the halo separately outside its hard bounds.
            Assert.Equal(OfficeColor.FromRgb(0, 0, 0), image!.GetPixel(83, 53));
            Assert.True(ContainsVisibleNonWhitePixel(image!, 93, 35, 8, 14), "Expected blurred shadow halo pixels outside the hard shadow bounds.");
            AssertNoUnexpectedDiagnostics(png.Diagnostics);
        }

        [Fact]
        public void PowerPointSlide_ProjectsGlowThroughSharedDrawingImageExport() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(120),OfficeIMO.OfficeLength.Points(90));
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "FFFFFF";

            PowerPointAutoShape shape = slide.AddShape(OfficePresetShapeType.Rectangle, OfficeIMO.OfficeLength.Points(32), OfficeIMO.OfficeLength.Points(24), OfficeIMO.OfficeLength.Points(48), OfficeIMO.OfficeLength.Points(28));
            shape.FillColor = "E0F2FE";
            shape.OutlineColor = "0284C7";
            shape.SetGlow("FF00FF", radiusPoints: 6, transparencyPercent: 25);

            OfficeImageExportResult svg = slide.ExportImage(OfficeImageExportFormat.Svg, new PowerPointImageExportOptions { Scale = 1D });
            OfficeRasterImage svgImage = ReadSvgOutput(svg);
            Assert.True(ContainsMagentaGlowPixel(svgImage, 24, 16, 64, 44), "Expected the glow in SVG output.");
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

    }
}
