using OfficeIMO.Drawing;
using System.Text.RegularExpressions;

namespace OfficeIMO.OpenDocument;

internal static class OdfImageCropCodec {
    internal static OdfInsets? Parse(string? lexical) {
        if (lexical == null) return null;
        lexical = lexical.Trim();
        if (lexical == "auto") return default(OdfInsets);
        if (!lexical.StartsWith("rect(", StringComparison.Ordinal) || !lexical.EndsWith(")", StringComparison.Ordinal))
            throw new InvalidDataException("Image crop must be auto or rect(top right bottom left).");
        string[] edges = lexical.Substring(5, lexical.Length - 6).Split(new[] { ' ', ',', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (edges.Length != 4) throw new InvalidDataException("Image crop requires four edge offsets.");
        var values = edges.Select(edge => edge == "auto" ? default : OdfLength.Parse(edge)).ToArray();
        var crop = new OdfInsets(values[0], values[1], values[2], values[3]);
        Validate(crop); return crop;
    }

    internal static string? Serialize(OdfInsets? crop) {
        if (!crop.HasValue) return null;
        Validate(crop.Value);
        return "rect(" + crop.Value.Top + ", " + crop.Value.Right + ", " + crop.Value.Bottom + ", " + crop.Value.Left + ")";
    }

    private static void Validate(OdfInsets crop) {
        if (new[] { crop.Top, crop.Right, crop.Bottom, crop.Left }.Any(edge =>
            !Regex.IsMatch(edge.ToString(), @"\A-?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:cm|mm|in|pt|pc)\z", RegexOptions.CultureInvariant) || !edge.TryToPoints(out _)))
            throw new InvalidDataException("Image crop offsets must be finite absolute ODF lengths.");
    }

    internal static OfficeImageProjection? Project(OdfInsets? crop, OfficeImageInfo image, double width, double height) {
        var frame = new OfficeImagePlacement(0, 0, width, height);
        if (!crop.HasValue) return new OfficeImageProjection(frame);
        OdfInsets c = crop.Value; Validate(c);
        double left = c.Left.ToPoints(), top = c.Top.ToPoints(), right = c.Right.ToPoints(), bottom = c.Bottom.ToPoints();
        if (left == 0 && top == 0 && right == 0 && bottom == 0) return new OfficeImageProjection(frame);
        GetIntrinsicSize(image, out double intrinsicWidth, out double intrinsicHeight);
        double visibleWidth = intrinsicWidth - left - right, visibleHeight = intrinsicHeight - top - bottom;
        if (!FinitePositive(intrinsicWidth) || !FinitePositive(intrinsicHeight) || !FinitePositive(visibleWidth) || !FinitePositive(visibleHeight))
            throw new NotSupportedException("Image crop must leave finite positive intrinsic dimensions; collapsed source crops are not qualified.");
        // Negative native offsets expand the source window with transparent space. Project only
        // the intersection with image pixels, keeping that space inside the original frame.
        double sourceLeft = Math.Max(left, 0), sourceTop = Math.Max(top, 0);
        double sourceRight = Math.Min(intrinsicWidth - right, intrinsicWidth), sourceBottom = Math.Min(intrinsicHeight - bottom, intrinsicHeight);
        if (sourceRight <= sourceLeft || sourceBottom <= sourceTop) return null;
        var sourceCrop = OfficeImageSourceCrop.FromStrictFractions(sourceLeft / intrinsicWidth, sourceTop / intrinsicHeight,
            (intrinsicWidth - sourceRight) / intrinsicWidth, (intrinsicHeight - sourceBottom) / intrinsicHeight);
        double x = left < 0 ? (sourceLeft - left) / visibleWidth * width : 0;
        double y = top < 0 ? (sourceTop - top) / visibleHeight * height : 0;
        double endX = right < 0 ? (sourceRight - left) / visibleWidth * width : width;
        double endY = bottom < 0 ? (sourceBottom - top) / visibleHeight * height : height;
        var placement = new OfficeImagePlacement(x, y, Math.Min(width - x, endX - x), Math.Min(height - y, endY - y));
        return new OfficeImageProjection(placement, sourceCrop);
    }

    /// <summary>Converts source-image fractions to native ODF offsets, independent of the displayed frame.</summary>
    /// <remarks>Only nonnegative crops that leave a visible raster source region are supported.</remarks>
    internal static OdfInsets? FromSourceFractions(OfficeImageInfo image, double left, double top, double right, double bottom) {
        ValidateSourceFractions(left, top, right, bottom);
        if (left == 0 && top == 0 && right == 0 && bottom == 0) return null;
        GetIntrinsicSize(image, out double width, out double height);
        var crop = new OdfInsets(OdfLength.Points(top * height), OdfLength.Points(right * width),
            OdfLength.Points(bottom * height), OdfLength.Points(left * width));
        // Absolute ODF lengths have finite stored precision. Recheck the serialized offsets
        // so a very narrow source window cannot become a collapsed native crop.
        _ = ToSourceFractions(crop, image);
        return crop;
    }

    /// <summary>Converts native ODF offsets to strict source-image fractions, independent of the displayed frame.</summary>
    /// <remarks>Negative offsets add blank margins and cannot be expressed by this source-crop-only profile.</remarks>
    internal static OfficeImageSourceCrop ToSourceFractions(OdfInsets crop, OfficeImageInfo image) {
        Validate(crop);
        double left = crop.Left.ToPoints(), top = crop.Top.ToPoints(), right = crop.Right.ToPoints(), bottom = crop.Bottom.ToPoints();
        if (left == 0 && top == 0 && right == 0 && bottom == 0) return default;
        GetIntrinsicSize(image, out double width, out double height);
        left /= width; right /= width; top /= height; bottom /= height;
        ValidateSourceFractions(left, top, right, bottom);
        return OfficeImageSourceCrop.FromStrictFractions(left, top, right, bottom);
    }

    private static void ValidateSourceFractions(double left, double top, double right, double bottom) {
        if (!FiniteNonnegative(left) || !FiniteNonnegative(top) || !FiniteNonnegative(right) || !FiniteNonnegative(bottom)
            || left + right >= 1 || top + bottom >= 1)
            throw new NotSupportedException("Source image crop must have finite nonnegative edges and leave a visible source area.");
    }

    private static void GetIntrinsicSize(OfficeImageInfo image, out double width, out double height) {
        if (image.Format is OfficeImageFormat.Svg or OfficeImageFormat.Emf or OfficeImageFormat.Wmf || image.Width <= 0 || image.Height <= 0)
            throw new NotSupportedException("Cropped vector images and images without intrinsic dimensions are outside this crop profile.");
        width = image.Width * 72D / image.DpiX; height = image.Height * 72D / image.DpiY;
        if (!FinitePositive(width) || !FinitePositive(height))
            throw new NotSupportedException("Image crop requires finite positive intrinsic dimensions.");
    }

    private static bool FiniteNonnegative(double value) => value >= 0 && !double.IsInfinity(value);
    private static bool FinitePositive(double value) => value > 0 && !double.IsInfinity(value);
}
