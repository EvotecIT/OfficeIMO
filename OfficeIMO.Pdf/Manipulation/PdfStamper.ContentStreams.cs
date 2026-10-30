using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfStamper {
    private static PdfStream BuildStampStream(
        string text,
        string fontResourceName,
        PdfReadPage page,
        PdfTextStampOptions options,
        bool watermarkDefaults) {
        PdfPageGeometry geometry = page.GetGeometry();
        var size = PdfReadPage.GetPageSize(geometry);
        double userUnit = geometry.UserUnit ?? 1D;
        double fontSize = options.FontSize.ToPoints() / userUnit;
        // Content operators use raw user space; physical inputs are measured in points.
        double x = options.X?.ToPoints() / userUnit ?? (watermarkDefaults
            ? (geometry.EffectiveBox?.Left ?? 0D) + (size.Width - PdfWriter.EstimateSimpleTextWidth(text, options.Font, fontSize)) / 2D
            : 36D / userUnit);
        double y = options.Y?.ToPoints() / userUnit ?? (watermarkDefaults
            ? (geometry.EffectiveBox?.Bottom ?? 0D) + size.Height / 2D
            : 36D / userUnit);
        double radians = options.RotationDegrees * Math.PI / 180.0;
        double cos = Math.Cos(radians);
        double sin = Math.Sin(radians);

        var sb = new StringBuilder();
        new ContentStreamBuilder(sb)
            .SaveState()
            .FillColor(options.Color)
            .BeginText()
            .Font(fontResourceName, fontSize, preserveLogicalPrecision: true)
            .TextMatrix(cos, sin, -sin, cos, x, y, precisePosition: true)
            .ShowHexText(EncodeWinAnsiHex(text))
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
