using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using static OfficeIMO.Tests.Pdf.PdfTransformedPageEditFixtures;

namespace OfficeIMO.Tests.Pdf;

/// <summary>
/// The operations that add content to, or flatten onto, an existing page. Requested positions are written in
/// top-left displayed-page points and mapped to the user-space coordinates each API takes.
/// </summary>
internal static class PdfTransformedPageEditOperations {
    private static readonly OfficeIMO.OfficeColor Red = OfficeIMO.OfficeColor.FromNormalizedRgb(1, 0, 0);
    private static OfficeIMO.OfficeLength Pt(double value) => OfficeIMO.OfficeLength.Points(value);

    internal static readonly string[] PlacedOperations = {
        "StampTextAt", "StampTextDefault", "StampTextBehind", "WatermarkText", "StampImageAt", "StampImageDefault",
        "CanvasText", "OverlayPage", "Redaction", "SquareFlatten", "FreeTextFlatten", "TextAdd", "ImageAdd", "Bates", "NUp"
    };

    internal static readonly string[] AppendingOperations = PlacedOperations
        .Concat(new[] { "Resize", "Crop", "NormalizeRotation" }).ToArray();

    internal static byte[] Apply(string operation, Fixture f) {
        switch (operation) {
            case "StampTextAt":
            case "StampTextBehind": {
                var p = f.ToUser(60, 150);
                return PdfStamper.StampText(f.Bytes, "EDITMARK", new PdfTextStampOptions {
                    X = Pt(p.X), Y = Pt(p.Y), Color = Red, FontSize = Pt(14), BehindContent = operation == "StampTextBehind"
                });
            }
            case "StampTextDefault":
                return PdfStamper.StampText(f.Bytes, "EDITMARK", new PdfTextStampOptions { Color = Red, FontSize = Pt(14) });
            case "WatermarkText":
                return PdfStamper.WatermarkText(f.Bytes, "DRAFT", new PdfTextStampOptions { Color = Red, Opacity = 1, FontSize = Pt(40) });
            case "StampImageAt": {
                var p = f.ToUser(60, 150);
                return PdfStamper.StampImage(f.Bytes, RedImagePngWithGreenTop(), new PdfImageStampOptions {
                    X = Pt(p.X), Y = Pt(p.Y), Width = Pt(40), Height = Pt(20)
                });
            }
            case "StampImageDefault":
                return PdfStamper.StampImage(f.Bytes, RedImagePngWithGreenTop(), new PdfImageStampOptions { Width = Pt(40), Height = Pt(20) });
            case "CanvasText":
                return PdfStamper.StampCanvas(f.Bytes, (canvas, _) => canvas.Text("EDITMARK", Pt(60), Pt(130), Pt(200), Pt(30), Pt(14), Red));
            case "OverlayPage":
                return PdfStamper.OverlayPage(f.Bytes, OverlaySource(f.DisplaySize), new PdfPageOverlayOptions());
            case "Redaction": {
                (double x, double y, double w, double h) = UserRectangle(f, 200, 60, 250, 90);
                return PdfDocument.Load(f.Bytes).Redactions.Apply(new[] { new PdfRedactionArea(1, x, y, w, h) },
                    new PdfRedactionApplyOptions { FillColor = Red }).ToBytes();
            }
            case "SquareFlatten": {
                (double x, double y, double w, double h) = UserRectangle(f, 200, 60, 250, 90);
                byte[] annotated = PdfDocument.Load(f.Bytes).Annotations.Add(new PdfAnnotationCreateOptions {
                    PageNumber = 1, Subtype = "Square", Rectangle = new[] { x, y, x + w, y + h },
                    Color = new double[] { 1, 0, 0 }, InteriorColor = new double[] { 1, 0, 0 }, GenerateAppearance = true
                }).Bytes;
                return PdfDocument.Load(annotated).Annotations.Flatten().Bytes;
            }
            case "FreeTextFlatten": {
                (double x, double y, double w, double h) = UserRectangle(f, 60, 120, 260, 160);
                byte[] annotated = PdfDocument.Load(f.Bytes).Annotations.Add(new PdfAnnotationCreateOptions {
                    PageNumber = 1, Subtype = "FreeText", Contents = "NOTEMARK", Rectangle = new[] { x, y, x + w, y + h },
                    Color = new double[] { 1, 0, 0 }, GenerateAppearance = true
                }).Bytes;
                return PdfDocument.Load(annotated).Annotations.Flatten().Bytes;
            }
            case "TextAdd": {
                // Text and image editor regions are measured from the page box origin.
                (double x, double y, double w, double h) = UserRectangle(f, 60, 120, 260, 160);
                return PdfDocument.Load(f.Bytes).Text.Add(new PdfPageRegion(1, x - f.Left, y - f.Bottom, w, h), "EDITMARK",
                    new PdfTextEditOptions { Color = Red, FontSize = 14 }).Document.ToBytes();
            }
            case "ImageAdd": {
                (double x, double y, double w, double h) = UserRectangle(f, 60, 130, 100, 150);
                return PdfDocument.Load(f.Bytes).Images.Add(new PdfPageRegion(1, x - f.Left, y - f.Bottom, w, h), RedImagePngWithGreenTop()).Document.ToBytes();
            }
            case "Bates":
                return PdfBatesNumberer.Apply(new[] { new PdfBatesDocument(f.Bytes) },
                    new PdfBatesNumberingOptions { Color = Red, Prefix = "BATES" }).Documents[0].ToBytes();
            case "NUp":
                return PdfDocument.Load(f.Bytes).Pages.ImposeNUp(new PdfNUpOptions(new PdfPageSize(Pt(900), Pt(500)), 2, 1)).Bytes;
            case "Resize":
                return PdfPageEditor.ResizePages(f.Bytes, new PdfPageSize(Pt(500), Pt(500)), 1);
            case "Crop":
                return PdfPageEditor.CropAndTranslatePages(f.Bytes, f.Left + 10, f.Bottom + 10, f.Right - 10, f.Top - 10, 1);
            case "NormalizeRotation":
                return PdfPageEditor.NormalizePageRotation(f.Bytes, null, null, default);
            default:
                throw new ArgumentOutOfRangeException(nameof(operation), operation, null);
        }
    }

    /// <summary>Returns null when the edit sits at its requested displayed-page position and reads upright.</summary>
    internal static string? CheckPlacement(string operation, Fixture f, OfficeRasterImage after, byte[] edited) {
        var size = f.DisplaySize;
        switch (operation) {
            case "StampTextAt":
            case "StampTextBehind":
                return ExpectText(f, after, edited, 60, 150, "EDITMARK");
            case "StampTextDefault":
                return ExpectText(f, after, edited, 36, size.Height - 36, "EDITMARK");
            case "WatermarkText": {
                Box red = Find(after, Ink.Red);
                if (red.IsEmpty) return "no watermark";
                if (Math.Abs(red.CenterX - size.Width / 2) > 8 || Math.Abs(red.CenterY - size.Height / 2) > 20) return $"watermark {red} not centred";
                return ExpectAngle(f, edited, "DRAFT", 0);
            }
            case "StampImageAt":
                return ExpectImage(after, 60, 150);
            case "StampImageDefault":
                return ExpectImage(after, 36, size.Height - 36);
            case "CanvasText": {
                Box red = Find(after, Ink.Red);
                if (red.IsEmpty) return "no canvas text";
                if (Math.Abs(red.Left - 60) > 6 || red.Top < 128 || red.Bottom > 165) return $"canvas text {red}";
                return ExpectAngle(f, edited, "EDITMARK", 0);
            }
            case "OverlayPage": {
                Box red = Find(after, Ink.Red);
                if (red.IsEmpty) return "no overlay";
                if (Math.Abs(red.Left - 150) > 3 || Math.Abs(red.Top - 105) > 3) return $"overlay red {red} expected at 150,105";
                return HasStripAbove(after, red) ? null : $"overlay not upright {red}";
            }
            case "Redaction":
                return ExpectRectangle(after, 200, 60, 50, 30, 2);
            case "SquareFlatten":
                return ExpectRectangle(after, 200, 60, 50, 30, 3);
            case "FreeTextFlatten":
                return ExpectAngle(f, edited, "NOTEMARK", 0);
            case "TextAdd": {
                Box red = Find(after, Ink.Red);
                if (red.IsEmpty) return "no text";
                if (red.Left < 55 || red.Right > 265 || red.Top < 115 || red.Bottom > 165) return $"text {red} outside region";
                return ExpectAngle(f, edited, "EDITMARK", 0);
            }
            case "ImageAdd":
                return ExpectImage(after, 60, 150);
            case "Bates": {
                Box red = Find(after, Ink.Red);
                if (red.IsEmpty) return "no Bates number";
                if (red.CenterX < size.Width / 2 || red.CenterY < size.Height / 2) return $"Bates number {red} not bottom-right";
                return ExpectAngle(f, edited, "BATES", 0);
            }
            case "NUp": {
                // The imposed page keeps the displayed orientation: the green mark stays on the same side of the blue one.
                OfficeRasterImage source = Render(f.Bytes);
                Box b0 = Find(source, Ink.Blue), g0 = Find(source, Ink.Green), b1 = Find(after, Ink.Blue), g1 = Find(after, Ink.Green);
                if (b1.IsEmpty || g1.IsEmpty) return "marks missing";
                bool same = Math.Sign(g0.CenterX - b0.CenterX) == Math.Sign(g1.CenterX - b1.CenterX) &&
                    Math.Sign(g0.CenterY - b0.CenterY) == Math.Sign(g1.CenterY - b1.CenterY);
                return same ? null : "imposed page orientation changed";
            }
            default:
                throw new ArgumentOutOfRangeException(nameof(operation), operation, null);
        }
    }

    /// <summary>
    /// A page carrying one annotation or push-button widget whose appearance is drawn in a frame turned by its /Matrix:
    /// a red body with a green strip along the frame's top edge.
    /// </summary>
    internal static byte[] BuildPageWithTurnedAppearance(int matrixRotation, bool widget) {
        const double rectLeft = 150, rectBottom = 100, rectWidth = 100, rectHeight = 50;
        bool quarter = matrixRotation % 180 != 0;
        double frameWidth = quarter ? rectHeight : rectWidth, frameHeight = quarter ? rectWidth : rectHeight;
        // A bounding box away from the origin checks that flattening maps the box, not the origin, into /Rect.
        double boxLeft = 10, boxBottom = 20;
        string appearance = "1 0 0 rg " + Fmt(boxLeft, boxBottom, frameWidth, frameHeight) + " re f 0 1 0 rg " +
            Fmt(boxLeft, boxBottom + frameHeight - 6, frameWidth, 6) + " re f";
        double radians = matrixRotation * Math.PI / 180D;
        double cos = Math.Round(Math.Cos(radians)), sin = Math.Round(Math.Sin(radians));
        string annotation = widget
            ? "<< /Type /Annot /Subtype /Widget /FT /Btn /Ff 65536 /T (Turned) /F 4 /P 3 0 R /Rect [" + Fmt(rectLeft, rectBottom, rectLeft + rectWidth, rectBottom + rectHeight) +
              "] /MK << /R " + matrixRotation + " >> /AP << /N 6 0 R >> >>"
            : "<< /Type /Annot /Subtype /Square /F 4 /P 3 0 R /Rect [" + Fmt(rectLeft, rectBottom, rectLeft + rectWidth, rectBottom + rectHeight) +
              "] /C [1 0 0] /AP << /N 6 0 R >> >>";
        string page = "0 0 1 rg 30 30 60 40 re f";
        return Serialize(new[] {
            "<< /Type /Catalog /Pages 2 0 R" + (widget ? " /AcroForm << /Fields [5 0 R] >>" : string.Empty) + " >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "<< /Length " + page.Length + " >>\nstream\n" + page + "\nendstream",
            annotation,
            "<< /Type /XObject /Subtype /Form /BBox [" + Fmt(boxLeft, boxBottom, boxLeft + frameWidth, boxBottom + frameHeight) +
                "] /Matrix [" + Fmt(cos, sin, -sin, cos, 0, 0) + "] /Length " + appearance.Length + " >>\nstream\n" + appearance + "\nendstream"
        });
    }

    private static (double X, double Y, double W, double H) UserRectangle(Fixture f, double left, double top, double right, double bottom) {
        var p1 = f.ToUser(left, top);
        var p2 = f.ToUser(right, bottom);
        return (Math.Min(p1.X, p2.X), Math.Min(p1.Y, p2.Y), Math.Abs(p2.X - p1.X), Math.Abs(p2.Y - p1.Y));
    }

    private static string? ExpectText(Fixture f, OfficeRasterImage after, byte[] edited, double left, double baseline, string word) {
        Box red = Find(after, Ink.Red);
        if (red.IsEmpty) return "no added text";
        if (Math.Abs(red.Left - left) > 6 || Math.Abs(red.Bottom - baseline) > 6)
            return $"text {red} expected to start at {left:0} on baseline {baseline:0}";
        if (red.PixelWidth <= red.PixelHeight) return $"text {red} is not horizontal";
        return ExpectAngle(f, edited, word, 0);
    }

    private static string? ExpectAngle(Fixture f, byte[] edited, string word, double angle) {
        PdfTextSpan[] spans = PdfReadDocument.Open(edited).Pages[0].GetTextSpans().Where(span => span.Text.Contains(word.Substring(0, 2))).ToArray();
        if (spans.Length == 0) return "added text was not extracted";
        double actual = DisplayAngle(spans[0], f.Rotation);
        if (Math.Abs(actual - angle) > 0.5D && Math.Abs(actual - angle - 360D) > 0.5D) return $"text reads at {actual:0} degrees, expected {angle:0}";
        string text = string.Concat(spans.Select(span => span.Text)).Replace(" ", string.Empty);
        return text.Contains(word) ? null : $"extracted '{text}'";
    }

    private static string? ExpectImage(OfficeRasterImage after, double left, double bottom) {
        Box red = Find(after, Ink.Red);
        if (red.IsEmpty) return "no added image";
        if (Math.Abs(red.Left - left) > 3 || Math.Abs(red.Bottom + 1 - bottom) > 3) return $"image {red} expected at left {left:0}, bottom {bottom:0}";
        return HasStripAbove(after, red) ? null : $"image {red} is not upright";
    }

    private static bool HasStripAbove(OfficeRasterImage after, Box red) {
        Box strip = Find(after, Ink.Green, red.Left, red.Top - 8, red.Right, red.Top - 1);
        return !strip.IsEmpty && strip.PixelWidth >= red.PixelWidth - 3;
    }

    private static string? ExpectRectangle(OfficeRasterImage after, double left, double top, double width, double height, double tolerance) {
        Box red = Find(after, Ink.Red);
        if (red.IsEmpty) return "no added rectangle";
        bool placed = Math.Abs(red.Left - left) <= tolerance && Math.Abs(red.Top - top) <= tolerance &&
            Math.Abs(red.PixelWidth - width) <= tolerance + 1 && Math.Abs(red.PixelHeight - height) <= tolerance + 1;
        return placed ? null : $"rectangle {red} expected at [{left:0},{top:0} {width:0}x{height:0}]";
    }

    private static byte[] OverlaySource((double Width, double Height) size) {
        // A red block under a green strip, 150 points from the left and 105 points from the top of the page.
        double top = size.Height - 100D;
        string content = "0 1 0 rg 150 " + Fmt(top - 5) + " 40 5 re f 1 0 0 rg 150 " + Fmt(top - 25) + " 40 20 re f";
        return Serialize(new[] {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " + Fmt(size.Width, size.Height) + "] /Contents 4 0 R >>",
            "<< /Length " + content.Length + " >>\nstream\n" + content + "\nendstream"
        });
    }
}
