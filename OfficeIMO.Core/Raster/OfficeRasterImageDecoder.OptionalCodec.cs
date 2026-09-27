namespace OfficeIMO.Drawing;

public static partial class OfficeRasterImageDecoder {
    private static bool TryDecodeWithOptionalCodec(
        byte[] bytes,
        OfficeRasterDecodeOptions options,
        OfficeRasterContainerInfo container,
        out OfficeRasterImage? image) {
        image = null;
        options.CancellationToken.ThrowIfCancellationRequested();
        if (options.ImageCodec == null || options.FrameIndex != 0 ||
            !IsWithinPixelLimit(container.CanvasWidth, container.CanvasHeight, options.MaximumDecodedPixels)) return false;

        // Providers may own their input buffer. Preserve the caller's encoded resource.
        byte[] codecBytes = (byte[])bytes.Clone();
        options.CancellationToken.ThrowIfCancellationRequested();
        bool decoded;
        try {
            decoded = options.ImageCodec.TryDecode(codecBytes, OfficeImageInfo.GetMimeType(container.Format), out image);
        } catch (System.Exception exception) when (exception is System.ArgumentException || exception is System.FormatException ||
            exception is System.NotSupportedException || exception is System.IO.IOException ||
            exception is System.InvalidOperationException || exception is System.OverflowException) {
            decoded = false;
            image = null;
        }
        options.CancellationToken.ThrowIfCancellationRequested();
        if (!decoded || image == null || image.Width != container.CanvasWidth || image.Height != container.CanvasHeight ||
            !IsDecodedImageWithinLimit(image, options.MaximumDecodedPixels)) {
            image = null;
            return false;
        }
        return true;
    }
}
