using OfficeIMO.Drawing;

namespace OfficeIMO.Word;

public partial class WordImage {
    private const double CropPercentageUnits = 100000D;

    /// <summary>Gets or sets the crop from the top in centimeters of the original image height, before displayed scaling.</summary>
    /// <remarks>Uses the embedded image's intrinsic dimensions and DPI. Setting a crop preserves its displayed scale, adjusts the frame and floating anchor, and respects rotation and mirrors. Negative values add empty margins. Null removes this edge and restores its area. Values are rounded to DrawingML percentage precision. Reading leaves the document XML unchanged.</remarks>
    /// <exception cref="InvalidOperationException">The embedded source size or floating anchor position cannot be resolved without pagination.</exception>
    /// <exception cref="ArgumentOutOfRangeException">The crop is not finite, removes the entire source area, or overflows the frame or anchor range.</exception>
    public double? CropTopCentimeters {
        get => GetCropCentimeters(CropTop, horizontal: false);
        set => SetCropCentimeters(value, horizontal: false, leading: true);
    }

    /// <summary>Gets or sets the crop from the bottom in centimeters of the original image height, before displayed scaling.</summary>
    /// <remarks>Uses intrinsic image dimensions and DPI, preserves displayed scale and the opposite edge, and adjusts the frame and floating anchor for rotation and mirrors. Negative values add empty margins; null removes this edge. Reading leaves the document XML unchanged.</remarks>
    /// <exception cref="InvalidOperationException">The embedded source size or floating anchor position cannot be resolved without pagination.</exception>
    /// <exception cref="ArgumentOutOfRangeException">The crop is not finite, removes the entire source area, or overflows the frame or anchor range.</exception>
    public double? CropBottomCentimeters {
        get => GetCropCentimeters(CropBottom, horizontal: false);
        set => SetCropCentimeters(value, horizontal: false, leading: false);
    }

    /// <summary>Gets or sets the crop from the left in centimeters of the original image width, before displayed scaling.</summary>
    /// <remarks>Uses intrinsic image dimensions and DPI, preserves displayed scale and the opposite edge, and adjusts the frame and floating anchor for rotation and mirrors. Negative values add empty margins; null removes this edge. Reading leaves the document XML unchanged.</remarks>
    /// <exception cref="InvalidOperationException">The embedded source size or floating anchor position cannot be resolved without pagination.</exception>
    /// <exception cref="ArgumentOutOfRangeException">The crop is not finite, removes the entire source area, or overflows the frame or anchor range.</exception>
    public double? CropLeftCentimeters {
        get => GetCropCentimeters(CropLeft, horizontal: true);
        set => SetCropCentimeters(value, horizontal: true, leading: true);
    }

    /// <summary>Gets or sets the crop from the right in centimeters of the original image width, before displayed scaling.</summary>
    /// <remarks>Uses intrinsic image dimensions and DPI, preserves displayed scale and the opposite edge, and adjusts the frame and floating anchor for rotation and mirrors. Negative values add empty margins; null removes this edge. Reading leaves the document XML unchanged.</remarks>
    /// <exception cref="InvalidOperationException">The embedded source size or floating anchor position cannot be resolved without pagination.</exception>
    /// <exception cref="ArgumentOutOfRangeException">The crop is not finite, removes the entire source area, or overflows the frame or anchor range.</exception>
    public double? CropRightCentimeters {
        get => GetCropCentimeters(CropRight, horizontal: true);
        set => SetCropCentimeters(value, horizontal: true, leading: false);
    }

    private double? GetCropCentimeters(int? crop, bool horizontal) {
        if (!crop.HasValue) return null;
        var size = GetIntrinsicCropSizeCentimeters();
        return crop.Value / CropPercentageUnits * (horizontal ? size.Width : size.Height);
    }

    private void SetCropCentimeters(double? value, bool horizontal, bool leading) {
        int? first = horizontal ? CropLeft : CropTop;
        int? last = horizontal ? CropRight : CropBottom;
        if (!value.HasValue && !(leading ? first : last).HasValue) return;
        int? crop = null;
        if (value.HasValue) {
            var size = GetIntrinsicCropSizeCentimeters();
            crop = ConvertCentimetersToCrop(value.Value, horizontal ? size.Width : size.Height, nameof(value));
        }
        if (leading) first = crop;
        else last = crop;
        ApplyPhysicalCrop(horizontal ? first : CropLeft, horizontal ? CropTop : first,
            horizontal ? last : CropRight, horizontal ? CropBottom : last);
    }

    /// <summary>Applies the fluent crop rectangle as one edit, validating both axes before changing either.</summary>
    internal void SetCropCentimeters(double left, double top, double right, double bottom) {
        var size = GetIntrinsicCropSizeCentimeters();
        int? cropLeft = ConvertCentimetersToCrop(left, size.Width, nameof(left));
        int? cropRight = ConvertCentimetersToCrop(right, size.Width, nameof(right));
        int? cropTop = ConvertCentimetersToCrop(top, size.Height, nameof(top));
        int? cropBottom = ConvertCentimetersToCrop(bottom, size.Height, nameof(bottom));
        ApplyPhysicalCrop(cropLeft, cropTop, cropRight, cropBottom);
    }

    private (double Width, double Height) GetIntrinsicCropSizeCentimeters() {
        if (GetPicture() == null)
            throw new InvalidOperationException("Physical cropping requires a DrawingML picture.");
        using var stream = OpenRead();
        if (!OfficeImageReader.TryIdentifyByContent(stream, FileName, out OfficeImageInfo info) ||
            info.Width <= 0 || info.Height <= 0)
            throw new InvalidOperationException("Physical cropping requires readable embedded image dimensions; use the raw percentage crop properties when the original size is unavailable.");
        double width = info.Width / info.DpiX * 2.54D;
        double height = info.Height / info.DpiY * 2.54D;
        if (!IsPositiveFiniteCropValue(width) || !IsPositiveFiniteCropValue(height))
            throw new InvalidOperationException("Physical cropping requires finite intrinsic image dimensions and DPI.");
        return (width, height);
    }

    private static int ConvertCentimetersToCrop(double centimeters, double originalSize, string parameter) {
        double crop = Math.Round(centimeters / originalSize * CropPercentageUnits, MidpointRounding.AwayFromZero);
        if (double.IsNaN(crop) || double.IsInfinity(crop) || crop < int.MinValue || crop > int.MaxValue)
            throw new ArgumentOutOfRangeException(parameter, "The crop must be finite and fit the DrawingML percentage range.");
        return checked((int)crop);
    }

    private long GetCroppedAxisExtent(bool horizontal, int? first, int? last) {
        if (GetPicture() == null)
            throw new InvalidOperationException("Physical cropping requires a DrawingML picture.");
        double currentExtent = (horizontal ? EmuWidth : EmuHeight) ?? 0D;
        double previousVisible = horizontal ? VisibleCropFraction(CropLeft, CropRight) : VisibleCropFraction(CropTop, CropBottom);
        double nextVisible = VisibleCropFraction(first, last);
        if (!IsPositiveFiniteCropValue(currentExtent) || !IsPositiveFiniteCropValue(previousVisible))
            throw new InvalidOperationException("The current picture must have a positive frame and a nonempty source crop before applying a physical crop.");
        if (!IsPositiveFiniteCropValue(nextVisible))
            throw new ArgumentOutOfRangeException(nameof(first), "Opposing crop edges must retain a nonempty image area.");
        double nextExtent = Math.Round(currentExtent * (nextVisible / previousVisible), MidpointRounding.AwayFromZero);
        if (nextExtent < 1D || nextExtent >= (double)long.MaxValue || double.IsNaN(nextExtent) || double.IsInfinity(nextExtent))
            throw new ArgumentOutOfRangeException(nameof(first), "The cropped picture frame must remain positive and fit the DrawingML extent range.");
        return checked((long)nextExtent);
    }

    private void ApplyPhysicalCrop(int? left, int? top, int? right, int? bottom) {
        long width = GetCroppedAxisExtent(horizontal: true, left, right);
        long height = GetCroppedAxisExtent(horizontal: false, top, bottom);
        // Resolve all geometry before changing the XML, including unsupported or malformed anchors.
        var position = GetPhysicalCropAnchorPosition(left, top, right, bottom, width, height);
        var inner = GetPicture()?.ShapeProperties?.Transform2D?.Extents;
        CropLeft = left; CropRight = right; CropTop = top; CropBottom = bottom;
        WordDrawingLayoutReader.SetExtent(_Image, width, height);
        if (inner != null) {
            inner.Cx = width;
            inner.Cy = height;
        }
        ApplyPhysicalCropAnchorPosition(position);
    }

    private static double VisibleCropFraction(int? first, int? last) =>
        1D - (first ?? 0) / CropPercentageUnits - (last ?? 0) / CropPercentageUnits;

    private static bool IsPositiveFiniteCropValue(double value) => value > 0D && !double.IsInfinity(value);
}
