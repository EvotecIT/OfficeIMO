namespace OfficeIMO.Drawing;

public static partial class OfficeRasterImageDecoder {
    // Static PNG consumers use IDAT, which may be separate from every APNG frame.
    // Keep the public selected-frame contract unchanged and share its container,
    // resource and cancellation checks before decoding the default image.
    internal static bool TryDecodePngDefault(byte[] bytes, OfficeRasterDecodeOptions options,
        out OfficeRasterImage? image) {
        image = null;
        if (!OfficeRasterContainerInspector.TryInspectForDecode(bytes, options, out var container, out _, out _, out var pngValidation) ||
            container?.Format != OfficeImageFormat.Png ||
            !OfficePngReader.TryDecode(bytes, options.CancellationToken, options.RetainedManagedBytes, out image, pngValidation) ||
            !IsDecodedImageWithinLimit(image, options.MaximumDecodedPixels)) {
            image = null;
            return false;
        }
        return true;
    }
}
