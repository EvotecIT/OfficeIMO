using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class DrawingGalleryPdf {
        public static void Example_Pdf_DrawingGallery(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.DrawingGallery.pdf");

            PdfDoc.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 10,
                    DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                    HeaderFont = PdfStandardFont.Helvetica,
                    HeaderFontSize = 8,
                    HeaderFormat = "OfficeIMO.Drawing shared vector gate",
                    HeaderAlign = PdfAlign.Left,
                    ShowHeader = true,
                    FooterFont = PdfStandardFont.Helvetica,
                    FooterFontSize = 8,
                    FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                    FooterAlign = PdfAlign.Right,
                    ShowPageNumbers = true
                })
                .Meta(title: "OfficeIMO.Pdf Drawing Gallery", author: "OfficeIMO")
                .H1("Drawing Gallery", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
                .Paragraph(p => p.Text("A visual baseline for shared OfficeIMO.Drawing vector descriptors rendered by the dependency-free PDF engine."))
                .Drawing(CreateDrawingScene(), PdfAlign.Center, spacingBefore: 8, spacingAfter: 10)
                .Paragraph(p => p
                    .Text("Covered: ")
                    .Bold("gradients")
                    .Text(", shadows, dashed strokes, line caps and joins, clipping paths, transforms, grouped scenes, and freeform paths."))
                .Save(path);

            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }

        private static OfficeDrawing CreateDrawingScene() {
            var drawing = new OfficeDrawing(420, 170);

            var background = OfficeShape.RoundedRectangle(420, 170, 12);
            background.FillColor = OfficeColor.FromRgb(248, 250, 252);
            background.StrokeColor = OfficeColor.FromRgb(183, 194, 207);
            background.StrokeWidth = 0.8;
            drawing.AddShape(background, 0, 0);

            var ribbon = OfficeShape.RoundedRectangle(132, 42, 10);
            ribbon.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.FromRgb(25, 55, 85), OfficeColor.FromRgb(20, 90, 180));
            ribbon.StrokeColor = OfficeColor.FromRgb(25, 55, 85);
            ribbon.StrokeWidth = 0.8;
            ribbon.Shadow = new OfficeShadow(OfficeColor.FromRgb(15, 23, 42), 0.18, 3, 3);
            drawing.AddShape(ribbon, 18, 18);

            var ellipse = OfficeShape.Ellipse(64, 42);
            ellipse.FillColor = OfficeColor.FromRgb(220, 252, 231);
            ellipse.StrokeColor = OfficeColor.FromRgb(22, 101, 52);
            ellipse.StrokeWidth = 1.2;
            drawing.AddShape(ellipse, 176, 18);

            var triangle = OfficeShape.Polygon(new OfficePoint(0, 38), new OfficePoint(36, 0), new OfficePoint(72, 38));
            triangle.FillColor = OfficeColor.FromRgb(254, 243, 199);
            triangle.StrokeColor = OfficeColor.FromRgb(180, 83, 9);
            triangle.StrokeWidth = 1.2;
            triangle.StrokeLineJoin = OfficeStrokeLineJoin.Round;
            drawing.AddShape(triangle, 276, 20);

            var rule = OfficeShape.Line(0, 0, 380, 0);
            rule.StrokeColor = OfficeColor.FromRgb(80, 80, 80);
            rule.StrokeWidth = 1.4;
            rule.StrokeDashStyle = OfficeStrokeDashStyle.DashDot;
            rule.StrokeLineCap = OfficeStrokeLineCap.Round;
            drawing.AddShape(rule, 20, 84);

            var clipped = OfficeShape.Rectangle(76, 44);
            clipped.FillColor = OfficeColor.FromRgb(219, 234, 254);
            clipped.StrokeColor = OfficeColor.FromRgb(20, 90, 180);
            clipped.StrokeWidth = 1;
            clipped.ClipPath = OfficeClipPath.RoundedRectangle(76, 44, 12);
            drawing.AddShape(clipped, 26, 110);

            var transformed = OfficeShape.Rectangle(74, 34);
            transformed.FillColor = OfficeColor.FromRgb(237, 233, 254);
            transformed.StrokeColor = OfficeColor.FromRgb(91, 33, 182);
            transformed.StrokeWidth = 1;
            transformed.Transform = OfficeTransform.RotateDegrees(8, 37, 17);
            drawing.AddShape(transformed, 132, 114);

            var path = OfficeShape.Path(
                OfficePathCommand.MoveTo(0, 34),
                OfficePathCommand.CubicBezierTo(22, -8, 60, -8, 82, 34),
                OfficePathCommand.LineTo(82, 44),
                OfficePathCommand.LineTo(0, 44),
                OfficePathCommand.Close());
            path.FillColor = OfficeColor.FromRgb(252, 231, 243);
            path.StrokeColor = OfficeColor.FromRgb(157, 23, 77);
            path.StrokeWidth = 1;
            drawing.AddShape(path, 238, 108);

            return drawing;
        }
    }
}
