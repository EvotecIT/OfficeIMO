using OfficeIMO.Drawing;

namespace OfficeIMO.Word;

public partial class WordImage {
    private const double CropPercentageUnits = 100000D;

    /// <summary>Gets or sets the crop from the top as a length of the original image height, before displayed scaling.</summary>
    /// <remarks>Uses the embedded image's intrinsic dimensions and DPI. Setting a crop preserves its displayed scale, adjusts the frame and floating anchor, and respects rotation and mirrors. Negative values add empty margins. Null removes this edge and restores its area. Values are rounded to DrawingML percentage precision. Reading leaves the document XML unchanged.</remarks>
    /// <exception cref="InvalidOperationException">The embedded source size or floating anchor position cannot be resolved without pagination.</exception>
    /// <exception cref="ArgumentOutOfRangeException">The crop is not finite, removes the entire source area, or overflows the frame or anchor range.</exception>
    public OfficeLength? CropTop {
        get => GetCropLength(CropTopPercentage, horizontal: false);
        set => SetCropCentimeters(value?.ToCentimeters(), horizontal: false, leading: true);
    }

    /// <summary>Gets or sets the crop from the bottom as a length of the original image height, before displayed scaling.</summary>
    /// <remarks>Uses intrinsic image dimensions and DPI, preserves displayed scale and the opposite edge, and adjusts the frame and floating anchor for rotation and mirrors. Negative values add empty margins; null removes this edge. Reading leaves the document XML unchanged.</remarks>
    /// <exception cref="InvalidOperationException">The embedded source size or floating anchor position cannot be resolved without pagination.</exception>
    /// <exception cref="ArgumentOutOfRangeException">The crop is not finite, removes the entire source area, or overflows the frame or anchor range.</exception>
    public OfficeLength? CropBottom {
        get => GetCropLength(CropBottomPercentage, horizontal: false);
        set => SetCropCentimeters(value?.ToCentimeters(), horizontal: false, leading: false);
    }

    /// <summary>Gets or sets the crop from the left as a length of the original image width, before displayed scaling.</summary>
    /// <remarks>Uses intrinsic image dimensions and DPI, preserves displayed scale and the opposite edge, and adjusts the frame and floating anchor for rotation and mirrors. Negative values add empty margins; null removes this edge. Reading leaves the document XML unchanged.</remarks>
    /// <exception cref="InvalidOperationException">The embedded source size or floating anchor position cannot be resolved without pagination.</exception>
    /// <exception cref="ArgumentOutOfRangeException">The crop is not finite, removes the entire source area, or overflows the frame or anchor range.</exception>
    public OfficeLength? CropLeft {
        get => GetCropLength(CropLeftPercentage, horizontal: true);
        set => SetCropCentimeters(value?.ToCentimeters(), horizontal: true, leading: true);
    }

    /// <summary>Gets or sets the crop from the right as a length of the original image width, before displayed scaling.</summary>
    /// <remarks>Uses intrinsic image dimensions and DPI, preserves displayed scale and the opposite edge, and adjusts the frame and floating anchor for rotation and mirrors. Negative values add empty margins; null removes this edge. Reading leaves the document XML unchanged.</remarks>
    /// <exception cref="InvalidOperationException">The embedded source size or floating anchor position cannot be resolved without pagination.</exception>
    /// <exception cref="ArgumentOutOfRangeException">The crop is not finite, removes the entire source area, or overflows the frame or anchor range.</exception>
    public OfficeLength? CropRight {
        get => GetCropLength(CropRightPercentage, horizontal: true);
        set => SetCropCentimeters(value?.ToCentimeters(), horizontal: true, leading: false);
    }

    private OfficeLength? GetCropLength(int? crop, bool horizontal) {
        double? centimeters = GetCropCentimeters(crop, horizontal);
        return centimeters.HasValue ? OfficeLength.Centimeters(centimeters.Value) : null;
    }

    private double? GetCropCentimeters(int? crop, bool horizontal) {
        if (!crop.HasValue) return null;
        var size = GetIntrinsicCropSizeCentimeters();
        return crop.Value / CropPercentageUnits * (horizontal ? size.Width : size.Height);
    }

    private void SetCropCentimeters(double? value, bool horizontal, bool leading) {
        int? first = horizontal ? CropLeftPercentage : CropTopPercentage;
        int? last = horizontal ? CropRightPercentage : CropBottomPercentage;
        if (!value.HasValue && !(leading ? first : last).HasValue) return;
        int? crop = null;
        if (value.HasValue) {
            var size = GetIntrinsicCropSizeCentimeters();
            crop = ConvertCentimetersToCrop(value.Value, horizontal ? size.Width : size.Height, nameof(value));
        }
        if (leading) first = crop;
        else last = crop;
        ApplyPhysicalCrop(horizontal ? first : CropLeftPercentage, horizontal ? CropTopPercentage : first,
            horizontal ? last : CropRightPercentage, horizontal ? CropBottomPercentage : last);
    }

    /// <summary>Applies all four physical source crops together, validating the frame and anchor before changing the image.</summary>
    /// <remarks>Null clears an edge. Negative lengths add source margins. Distances use the original image resolution and retain its displayed scale.</remarks>
    public WordImage SetCrop(OfficeLength? left, OfficeLength? top, OfficeLength? right, OfficeLength? bottom) {
        (double Width, double Height) size = left.HasValue || top.HasValue || right.HasValue || bottom.HasValue
            ? GetIntrinsicCropSizeCentimeters() : default;
        int? cropLeft = left.HasValue ? ConvertCentimetersToCrop(left.Value.ToCentimeters(), size.Width, nameof(left)) : null;
        int? cropRight = right.HasValue ? ConvertCentimetersToCrop(right.Value.ToCentimeters(), size.Width, nameof(right)) : null;
        int? cropTop = top.HasValue ? ConvertCentimetersToCrop(top.Value.ToCentimeters(), size.Height, nameof(top)) : null;
        int? cropBottom = bottom.HasValue ? ConvertCentimetersToCrop(bottom.Value.ToCentimeters(), size.Height, nameof(bottom)) : null;
        ApplyPhysicalCrop(cropLeft, cropTop, cropRight, cropBottom);
        return this;
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
        double currentExtent = (horizontal ? ((double?)Width?.ToEmus()) : ((double?)Height?.ToEmus())) ?? 0D;
        double previousVisible = horizontal ? VisibleCropFraction(CropLeftPercentage, CropRightPercentage) : VisibleCropFraction(CropTopPercentage, CropBottomPercentage);
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
        CropLeftPercentage = left; CropRightPercentage = right; CropTopPercentage = top; CropBottomPercentage = bottom;
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
