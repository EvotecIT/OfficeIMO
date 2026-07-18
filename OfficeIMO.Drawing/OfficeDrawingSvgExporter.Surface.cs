using System;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingSvgExporter {
    /// <summary>
    /// Converts a drawing to an SVG document with an explicit root size unit.
    /// </summary>
    /// <param name="drawing">Drawing to export.</param>
    /// <param name="scale">Scale applied to the exported SVG width and height.</param>
    /// <param name="sizeUnit">Unit written on the root width and height attributes.</param>
    /// <returns>SVG markup representing the drawing.</returns>
    public static string ToSvg(OfficeDrawing drawing, double scale, OfficeSvgSizeUnit sizeUnit) {
        return ToSvg(drawing, scale, sizeUnit, null);
    }

    /// <summary>Converts a drawing to SVG and uses an optional shared codec for source images that require transcoding.</summary>
    public static string ToSvg(
        OfficeDrawing drawing,
        double scale,
        OfficeSvgSizeUnit sizeUnit,
        IOfficeRasterImageCodec? imageCodec) {
        return ToSvg(drawing, scale, sizeUnit, imageCodec, null);
    }

    /// <summary>Converts a drawing to SVG and prefixes every generated resource identifier for safe inline composition.</summary>
    public static string ToSvg(
        OfficeDrawing drawing,
        double scale,
        OfficeSvgSizeUnit sizeUnit,
        IOfficeRasterImageCodec? imageCodec,
        string? resourceIdPrefix) {
        if (drawing == null) throw new ArgumentNullException(nameof(drawing));
        if (double.IsNaN(scale) || double.IsInfinity(scale) || scale <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(scale), "Scale must be a positive finite value.");
        }
        if (!Enum.IsDefined(typeof(OfficeSvgSizeUnit), sizeUnit)) {
            throw new ArgumentOutOfRangeException(nameof(sizeUnit));
        }
        string idPrefix = ValidateResourceIdPrefix(resourceIdPrefix);

        double width = drawing.Width * scale;
        double height = drawing.Height * scale;
        if (sizeUnit == OfficeSvgSizeUnit.Pixel) {
            width = Math.Ceiling(width);
            height = Math.Ceiling(height);
        }
        string unit = sizeUnit == OfficeSvgSizeUnit.Pixel ? "px" : "pt";
        var builder = new StringBuilder();
        builder.Append("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"")
            .Append(Format(width))
            .Append(unit)
            .Append("\" height=\"")
            .Append(Format(height))
            .Append(unit)
            .Append("\" viewBox=\"0 0 ")
            .Append(Format(drawing.Width))
            .Append(' ')
            .Append(Format(drawing.Height))
            .Append("\" role=\"img\">");

        AppendEmbeddedFonts(builder, drawing.Fonts);
        int gradientId = 0;
        int clipPathId = 0;
        AppendElements(builder, drawing.Elements, imageCodec, idPrefix, ref gradientId, ref clipPathId);
        builder.Append("</svg>");
        return builder.ToString();
    }

    /// <summary>Converts a drawing to UTF-8 SVG bytes with an explicit root size unit.</summary>
    public static byte[] ToSvgBytes(OfficeDrawing drawing, double scale, OfficeSvgSizeUnit sizeUnit) =>
        Encoding.UTF8.GetBytes(ToSvg(drawing, scale, sizeUnit));

    /// <summary>Converts a drawing to UTF-8 SVG bytes and uses an optional shared image codec.</summary>
    public static byte[] ToSvgBytes(
        OfficeDrawing drawing,
        double scale,
        OfficeSvgSizeUnit sizeUnit,
        IOfficeRasterImageCodec? imageCodec) =>
        Encoding.UTF8.GetBytes(ToSvg(drawing, scale, sizeUnit, imageCodec));

    /// <summary>Converts a drawing to UTF-8 SVG bytes and prefixes generated resource identifiers for safe inline composition.</summary>
    public static byte[] ToSvgBytes(
        OfficeDrawing drawing,
        double scale,
        OfficeSvgSizeUnit sizeUnit,
        IOfficeRasterImageCodec? imageCodec,
        string? resourceIdPrefix) =>
        Encoding.UTF8.GetBytes(ToSvg(drawing, scale, sizeUnit, imageCodec, resourceIdPrefix));

    private static string ValidateResourceIdPrefix(string? value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        for (int index = 0; index < value!.Length; index++) {
            char character = value[index];
            if ((character >= 'a' && character <= 'z') ||
                (character >= 'A' && character <= 'Z') ||
                (character >= '0' && character <= '9') ||
                character == '-' || character == '_' || character == '.' || character == ':') continue;
            throw new ArgumentException("An SVG resource identifier prefix can contain only ASCII letters, digits, '-', '_', '.', and ':'.", nameof(value));
        }
        return value;
    }
}
