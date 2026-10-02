namespace OfficeIMO.Drawing;

public static partial class OfficeRasterImageDecoder {
    /// <summary>Preserves rejection and encoded limits when a surface supports custom image formats.</summary>
    internal static bool CanUseUninspectedCallerCodec(byte[] bytes, OfficeRasterDecodeOptions options, OfficeRasterDecodeInfo info) =>
        bytes.Length <= options.MaximumEncodedBytes && info.Container == null &&
        info.Format != OfficeImageFormat.Webp && !OfficeImageReader.HasWebpSignature(bytes);

    private static bool TryDecodeWithOptionalCodec(
        byte[] bytes,
        OfficeRasterDecodeOptions options,
        OfficeRasterContainerInfo container,
        out OfficeRasterImage? image) {
        image = null;
        options.CancellationToken.ThrowIfCancellationRequested();
        IOfficeRasterImageCodec? codec = options.ImageCodec;
        while (codec is OfficeRasterImageFallbackCodec fallback) codec = fallback.SourceCodec;
        if (codec == null || options.FrameIndex != 0 ||
            !IsWithinPixelLimit(container.CanvasWidth, container.CanvasHeight, options.MaximumDecodedPixels)) return false;

        // Include the original resource, provider-owned clone and expected RGBA
        // alongside bytes retained by stream conversion before allocating or calling out.
        long remaining = OfficeRasterGuards.MaximumDecodedBytes - options.RetainedManagedBytes;
        long ownedBytes = 2L * bytes.Length + 4L * container.CanvasWidth * container.CanvasHeight;
        if (remaining < ownedBytes) return false;

        // Providers may own their input buffer. Preserve the caller's encoded resource.
        byte[] codecBytes = (byte[])bytes.Clone();
        options.CancellationToken.ThrowIfCancellationRequested();
        bool decoded;
        try {
            decoded = codec.TryDecode(codecBytes, OfficeImageInfo.GetMimeType(container.Format), out image);
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
