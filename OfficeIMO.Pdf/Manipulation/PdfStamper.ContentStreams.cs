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
        double? textWidth = null,
        string? graphicsStateResourceName = null) {
        PdfPageGeometry geometry = page.GetGeometry();
        var frame = DisplayFrame.Create(page, geometry);
        double userUnit = geometry.UserUnit ?? 1D;
        double fontSize = options.FontSize.ToPoints() / userUnit;
        // Authoring inputs and measured embedded-font widths are physical points; operators use raw page user space.
        double width = textWidth.HasValue ? textWidth.Value / userUnit
            : PdfWriter.EstimateSimpleTextWidth(text, options.Font, fontSize);
        // Defaults are measured on the displayed page; explicit coordinates stay in unrotated user space.
        double margin = 36D / userUnit;
        double anchorDisplayX = options.Align switch {
            OfficeAlignment.Center => frame.Width / 2D,
            OfficeAlignment.Right => frame.Width - margin,
            _ => watermarkDefaults ? (frame.Width - width) / 2D : margin
        };
        (double x, double y) = frame.ToUser(anchorDisplayX, watermarkDefaults ? frame.Height / 2D : margin);
        if (options.X.HasValue) x = options.X.Value.ToPoints() / userUnit;
        if (options.Y.HasValue) y = options.Y.Value.ToPoints() / userUnit;
        // Rotation is relative to the displayed page, so text reads upright on rotated pages.
        double radians = (options.RotationDegrees + frame.Rotation) * Math.PI / 180D;
        double cos = Math.Cos(radians), sin = Math.Sin(radians);
        if (options.Align != OfficeAlignment.Left) {
            double offset = options.Align == OfficeAlignment.Center ? width / 2D : width;
            x -= offset * cos;
            y -= offset * sin;
        }

        var sb = new StringBuilder();
        var content = new ContentStreamBuilder(sb).SaveState();
        if (graphicsStateResourceName != null) content.GraphicsState(graphicsStateResourceName);
        content.FillColor(options.Color.Resolve(OfficeTheme.Default))
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
        bool watermarkDefaults,
        string? graphicsStateResourceName = null) {
        PdfPageGeometry geometry = page.GetGeometry();
        var frame = DisplayFrame.Create(page, geometry);
        double userUnit = geometry.UserUnit ?? 1D;
        // Preserve the intrinsic one-pixel-to-one-physical-point default.
        double imageWidth = (options.Width?.ToPoints() ?? pixelWidth) / userUnit;
        double imageHeight = (options.Height?.ToPoints() ?? pixelHeight) / userUnit;
        // Defaults are measured on the displayed page; explicit coordinates stay in unrotated user space.
        (double x, double y) = watermarkDefaults
            ? frame.ToUser((frame.Width - imageWidth) / 2D, (frame.Height - imageHeight) / 2D)
            : frame.ToUser(36D / userUnit, 36D / userUnit);
        if (options.X.HasValue) x = options.X.Value.ToPoints() / userUnit;
        if (options.Y.HasValue) y = options.Y.Value.ToPoints() / userUnit;
        OfficeTransform imageTransform = new OfficeImageProjection(
            new OfficeImagePlacement(x, y, imageWidth, imageHeight),
            // Rotation is relative to the displayed page unless an editor replays an existing user-space placement.
            rotationDegrees: options.RotationDegrees + (options.RotationInUserSpace ? 0 : frame.Rotation),
            rotationCenterX: x,
            rotationCenterY: y)
            .CreateUnitSquareTransform();

        if (orientation != OfficeImageOrientation.Normal) {
            imageTransform = PdfWriter.GetImageOrientationTransform(orientation).Then(imageTransform);
        }
        var sb = new StringBuilder();
        var content = new ContentStreamBuilder(sb).SaveState();
        if (graphicsStateResourceName != null) content.GraphicsState(graphicsStateResourceName);
        content.TransformMatrix(imageTransform.M11, imageTransform.M12, imageTransform.M21, imageTransform.M22, imageTransform.OffsetX, imageTransform.OffsetY, preciseCoordinates: true)
            .XObject(imageResourceName)
            .RestoreState();

        return new PdfStream(new PdfDictionary(), PdfEncoding.Latin1GetBytes(sb.ToString()));
    }
    /// <summary>The effective page box as it is displayed after /Rotate, expressed in raw user-space units.</summary>
    internal readonly struct DisplayFrame {
        private readonly double _left, _bottom, _right, _top;

        private DisplayFrame(double left, double bottom, double right, double top, int rotation) {
            _left = left; _bottom = bottom; _right = right; _top = top;
            Rotation = rotation;
        }

        /// <summary>Clockwise display rotation: 0, 90, 180 or 270.</summary>
        internal int Rotation { get; }
        internal double Width => Rotation % 180 == 0 ? _right - _left : _top - _bottom;
        internal double Height => Rotation % 180 == 0 ? _top - _bottom : _right - _left;

        internal static DisplayFrame Create(PdfReadPage page, PdfPageGeometry geometry) {
            var size = PdfReadPage.GetPageSize(geometry);
            double left = geometry.EffectiveBox?.Left ?? 0D;
            double bottom = geometry.EffectiveBox?.Bottom ?? 0D;
            int rotation = (int)(Math.Round(page.GetRotationDegrees() / 90D) % 4D) * 90;
            return new DisplayFrame(left, bottom, left + size.Width, bottom + size.Height, rotation);
        }

        /// <summary>Frames a user-space box shown with the given clockwise page rotation.</summary>
        internal static DisplayFrame ForBox(double left, double bottom, double right, double top, int rotation) =>
            new(left, bottom, right, top, rotation);

        /// <summary>Maps a point measured from the displayed page's lower-left corner to user space.</summary>
        internal (double X, double Y) ToUser(double x, double y) => Rotation switch {
            90 => (_right - y, _bottom + x),
            180 => (_right - x, _top - y),
            270 => (_left + y, _top - x),
            _ => (_left + x, _bottom + y)
        };
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
