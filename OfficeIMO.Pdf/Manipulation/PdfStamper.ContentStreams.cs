using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfStamper {
    private static PdfStream BuildStampStream(
        string text,
        string fontResourceName,
        double pageWidth,
        double pageHeight,
        PdfTextStampOptions options,
        bool watermarkDefaults) {
        double fontSize = options.FontSize;
        double x = options.X ?? (watermarkDefaults ? (pageWidth - PdfWriter.EstimateSimpleTextWidth(text, options.Font, fontSize)) / 2.0 : 36);
        double y = options.Y ?? (watermarkDefaults ? pageHeight / 2.0 : 36);
        double radians = options.RotationDegrees * Math.PI / 180.0;
        double cos = Math.Cos(radians);
        double sin = Math.Sin(radians);

        var sb = new StringBuilder();
        new ContentStreamBuilder(sb)
            .SaveState()
            .FillColor(options.Color)
            .BeginText()
            .Font(fontResourceName, fontSize)
            .TextMatrix(cos, sin, -sin, cos, x, y)
            .ShowHexText(EncodeWinAnsiHex(text))
            .EndText()
            .RestoreState();

        return new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes(sb.ToString()));
    }

    private static PdfStream BuildImageStampStream(
        string imageResourceName,
        double pageWidth,
        double pageHeight,
        int pixelWidth,
        int pixelHeight,
        PdfImageStampOptions options,
        bool watermarkDefaults) {
        double imageWidth = options.Width ?? pixelWidth;
        double imageHeight = options.Height ?? pixelHeight;
        double x = options.X ?? (watermarkDefaults ? (pageWidth - imageWidth) / 2.0 : 36);
        double y = options.Y ?? (watermarkDefaults ? (pageHeight - imageHeight) / 2.0 : 36);
        OfficeTransform imageTransform = new OfficeImageProjection(
            new OfficeImagePlacement(x, y, imageWidth, imageHeight),
            rotationDegrees: options.RotationDegrees,
            rotationCenterX: x,
            rotationCenterY: y)
            .CreateUnitSquareTransform();

        var sb = new StringBuilder();
        new ContentStreamBuilder(sb)
            .SaveState()
            .TransformMatrix(imageTransform)
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
