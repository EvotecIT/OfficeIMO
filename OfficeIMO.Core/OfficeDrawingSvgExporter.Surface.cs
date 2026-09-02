using System;
using System.Text;
using System.Threading;

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
        return ToSvg(drawing, scale, sizeUnit, imageCodec, resourceIdPrefix, CancellationToken.None);
    }

    /// <summary>Converts a drawing to SVG with bounded cancellation and a safe generated-resource prefix.</summary>
    public static string ToSvg(
        OfficeDrawing drawing,
        double scale,
        OfficeSvgSizeUnit sizeUnit,
        IOfficeRasterImageCodec? imageCodec,
        string? resourceIdPrefix,
        CancellationToken cancellationToken) {
        return ToSvg(
            drawing,
            scale,
            sizeUnit,
            imageCodec,
            resourceIdPrefix,
            cancellationToken,
            new SvgNearestNeighborRectangleBudget());
    }

    internal static string ToSvg(
        OfficeDrawing drawing,
        double scale,
        OfficeSvgSizeUnit sizeUnit,
        IOfficeRasterImageCodec? imageCodec,
        string? resourceIdPrefix,
        CancellationToken cancellationToken,
        SvgNearestNeighborRectangleBudget nearestNeighborRectangleBudget) {
        return ToSvgCore(
            drawing,
            scale,
            sizeUnit,
            imageCodec,
            resourceIdPrefix,
            cancellationToken,
            nearestNeighborRectangleBudget,
            maximumCharacters: null);
    }

    private static string ToSvgCore(
        OfficeDrawing drawing,
        double scale,
        OfficeSvgSizeUnit sizeUnit,
        IOfficeRasterImageCodec? imageCodec,
        string? resourceIdPrefix,
        CancellationToken cancellationToken,
        SvgNearestNeighborRectangleBudget nearestNeighborRectangleBudget,
        int? maximumCharacters) {
        if (drawing == null) throw new ArgumentNullException(nameof(drawing));
        if (nearestNeighborRectangleBudget == null) throw new ArgumentNullException(nameof(nearestNeighborRectangleBudget));
        cancellationToken.ThrowIfCancellationRequested();
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
        var builder = maximumCharacters.HasValue
            ? new StringBuilder(Math.Min(256, maximumCharacters.Value), maximumCharacters.Value)
            : new StringBuilder();
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

        AppendEmbeddedFonts(builder, drawing.Fonts, cancellationToken);
        int gradientId = 0;
        int clipPathId = 0;
        var tilingExpansionBudget = new SvgTilingExpansionBudget();
        AppendElements(builder, drawing.Elements, imageCodec, idPrefix, ref gradientId, ref clipPathId, cancellationToken, tilingExpansionBudget, nearestNeighborRectangleBudget);
        builder.Append("</svg>");
        cancellationToken.ThrowIfCancellationRequested();
        string svg = builder.ToString();
        cancellationToken.ThrowIfCancellationRequested();
        return svg;
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

    /// <summary>Converts a drawing to cancellable UTF-8 SVG bytes and prefixes generated resource identifiers.</summary>
    public static byte[] ToSvgBytes(
        OfficeDrawing drawing,
        double scale,
        OfficeSvgSizeUnit sizeUnit,
        IOfficeRasterImageCodec? imageCodec,
        string? resourceIdPrefix,
        CancellationToken cancellationToken) {
        string svg = ToSvg(drawing, scale, sizeUnit, imageCodec, resourceIdPrefix, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = Encoding.UTF8.GetBytes(svg);
        cancellationToken.ThrowIfCancellationRequested();
        return bytes;
    }

    internal static byte[] ToSvgBytes(
        OfficeDrawing drawing,
        double scale,
        OfficeSvgSizeUnit sizeUnit,
        IOfficeRasterImageCodec? imageCodec,
        string? resourceIdPrefix,
        long maximumUtf8Bytes,
        CancellationToken cancellationToken) {
        if (maximumUtf8Bytes < 1L) throw new ArgumentOutOfRangeException(nameof(maximumUtf8Bytes));
        if (drawing == null) throw new ArgumentNullException(nameof(drawing));
        if (double.IsNaN(scale) || double.IsInfinity(scale) || scale <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(scale), "Scale must be a positive finite value.");
        }
        if (!Enum.IsDefined(typeof(OfficeSvgSizeUnit), sizeUnit)) {
            throw new ArgumentOutOfRangeException(nameof(sizeUnit));
        }
        int maximumCharacters = maximumUtf8Bytes > int.MaxValue
            ? int.MaxValue
            : (int)maximumUtf8Bytes;
        string svg;
        try {
            svg = ToSvgCore(
                drawing,
                scale,
                sizeUnit,
                imageCodec,
                resourceIdPrefix,
                cancellationToken,
                new SvgNearestNeighborRectangleBudget(),
                maximumCharacters);
        } catch (ArgumentOutOfRangeException) {
            throw new OfficeImageExportBatchLimitException(
                nameof(OfficeImageExportOptions.MaximumTotalEncodedBytes),
                maximumUtf8Bytes == long.MaxValue ? long.MaxValue : maximumUtf8Bytes + 1L,
                maximumUtf8Bytes);
        }
        cancellationToken.ThrowIfCancellationRequested();
        long byteCount = GetUtf8ByteCount(svg, maximumUtf8Bytes, cancellationToken);
        if (byteCount > maximumUtf8Bytes) {
            throw new OfficeImageExportBatchLimitException(
                nameof(OfficeImageExportOptions.MaximumTotalEncodedBytes),
                byteCount,
                maximumUtf8Bytes);
        }
        byte[] bytes = EncodeUtf8(svg, checked((int)byteCount), cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return bytes;
    }

    private static long GetUtf8ByteCount(
        string value,
        long maximumUtf8Bytes,
        CancellationToken cancellationToken) {
        long total = 0L;
        var characters = new char[16 * 1024];
        for (int offset = 0; offset < value.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = GetUtf8ChunkLength(value, offset);
            value.CopyTo(offset, characters, 0, count);
            total = checked(total + Encoding.UTF8.GetByteCount(characters, 0, count));
            if (total > maximumUtf8Bytes) return total;
            offset += count;
        }
        return total;
    }

    private static byte[] EncodeUtf8(
        string value,
        int byteCount,
        CancellationToken cancellationToken) {
        var bytes = new byte[byteCount];
        var characters = new char[16 * 1024];
        int byteOffset = 0;
        for (int offset = 0; offset < value.Length;) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = GetUtf8ChunkLength(value, offset);
            value.CopyTo(offset, characters, 0, count);
            byteOffset += Encoding.UTF8.GetBytes(characters, 0, count, bytes, byteOffset);
            offset += count;
        }
        return bytes;
    }

    private static int GetUtf8ChunkLength(string value, int offset) {
        const int ChunkCharacters = 16 * 1024;
        int count = Math.Min(ChunkCharacters, value.Length - offset);
        int end = offset + count;
        if (end < value.Length &&
            char.IsHighSurrogate(value[end - 1]) &&
            char.IsLowSurrogate(value[end])) {
            count--;
        }
        return count;
    }

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
