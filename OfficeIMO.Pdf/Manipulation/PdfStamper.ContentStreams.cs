using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfStamper {
    private static PdfStream BuildStampStream(
        string text,
        string fontResourceName,
        PdfReadPage page,
        PdfTextStampOptions options,
        bool watermarkDefaults,
        string? glyphHex = null,
        double? textWidth = null) {
        PdfPageGeometry geometry = page.GetGeometry();
        var size = PdfReadPage.GetPageSize(geometry);
        double userUnit = geometry.UserUnit ?? 1D;
        double fontSize = options.FontSize.ToPoints() / userUnit;
        // Authoring inputs and measured embedded-font widths are physical points; operators use raw page user space.
        double width = textWidth.HasValue ? textWidth.Value / userUnit
            : PdfWriter.EstimateSimpleTextWidth(text, options.Font, fontSize);
        double left = geometry.EffectiveBox?.Left ?? 0D;
        double bottom = geometry.EffectiveBox?.Bottom ?? 0D;
        double y = options.Y?.ToPoints() / userUnit ?? (watermarkDefaults ? bottom + size.Height / 2D : 36D / userUnit);
        double radians = options.RotationDegrees * Math.PI / 180D;
        double cos = Math.Cos(radians), sin = Math.Sin(radians);
        double x;
        if (options.Align == PdfAlign.Left) {
            x = options.X?.ToPoints() / userUnit ?? (watermarkDefaults ? left + (size.Width - width) / 2D : 36D / userUnit);
        } else {
            double anchorX = options.X?.ToPoints() / userUnit ?? (options.Align == PdfAlign.Center
                ? left + size.Width / 2D : left + size.Width - 36D / userUnit);
            double offset = options.Align == PdfAlign.Center ? width / 2D : width;
            x = anchorX - offset * cos;
            y -= offset * sin;
        }

        var sb = new StringBuilder();
        new ContentStreamBuilder(sb)
            .SaveState()
            .FillColor(options.Color)
            .BeginText()
            .Font(fontResourceName, fontSize, preserveLogicalPrecision: true)
            .TextMatrix(cos, sin, -sin, cos, x, y, precisePosition: true)
            .ShowHexText(glyphHex ?? EncodeWinAnsiHex(text))
            .EndText()
            .RestoreState();

        return new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes(sb.ToString()));
    }

    private static PdfStream BuildImageStampStream(
        string imageResourceName,
        PdfReadPage page,
        int pixelWidth,
        int pixelHeight,
        OfficeImageOrientation orientation,
        PdfImageStampOptions options,
        bool watermarkDefaults) {
        PdfPageGeometry geometry = page.GetGeometry();
        var size = PdfReadPage.GetPageSize(geometry);
        double userUnit = geometry.UserUnit ?? 1D;
        // Preserve the intrinsic one-pixel-to-one-physical-point default.
        double imageWidth = (options.Width?.ToPoints() ?? pixelWidth) / userUnit;
        double imageHeight = (options.Height?.ToPoints() ?? pixelHeight) / userUnit;
        double x = options.X?.ToPoints() / userUnit ?? (watermarkDefaults
            ? (geometry.EffectiveBox?.Left ?? 0D) + (size.Width - imageWidth) / 2D
            : 36D / userUnit);
        double y = options.Y?.ToPoints() / userUnit ?? (watermarkDefaults
            ? (geometry.EffectiveBox?.Bottom ?? 0D) + (size.Height - imageHeight) / 2D
            : 36D / userUnit);
        OfficeTransform imageTransform = new OfficeImageProjection(
            new OfficeImagePlacement(x, y, imageWidth, imageHeight),
            rotationDegrees: options.RotationDegrees,
            rotationCenterX: x,
            rotationCenterY: y)
            .CreateUnitSquareTransform();

        if (orientation != OfficeImageOrientation.Normal) {
            imageTransform = PdfWriter.GetImageOrientationTransform(orientation).Then(imageTransform);
        }
        var sb = new StringBuilder();
        new ContentStreamBuilder(sb)
            .SaveState()
            .TransformMatrix(imageTransform.M11, imageTransform.M12, imageTransform.M21, imageTransform.M22, imageTransform.OffsetX, imageTransform.OffsetY, preciseCoordinates: true)
            .XObject(imageResourceName)
            .RestoreState();

        return new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes(sb.ToString()));
    }
    private static string EncodeWinAnsiHex(string text) {
        var bytes = PdfWinAnsiEncoding.Encode(text);
        var sb = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            sb.Append(bytes[i].ToString("X2", CultureInfo.InvariantCulture));
        }

        return sb.ToString();
    }
}
