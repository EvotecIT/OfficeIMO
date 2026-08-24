using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Dependency-free baseline and progressive JPEG decoder plus JPEG encoder.</summary>
public static partial class OfficeJpegCodec {
    /// <summary>Returns whether the payload starts with the JPEG start-of-image marker.</summary>
    public static bool IsJpeg(byte[]? encodedBytes) =>
        encodedBytes != null && OfficeJpegReader.IsJpeg(encodedBytes);

    /// <summary>Attempts to decode JPEG bytes into an RGBA image.</summary>
    public static bool TryDecode(byte[]? encodedBytes, out OfficeRasterImage? image, OfficeJpegDecodeOptions options = default) {
        return TryDecode(encodedBytes, CancellationToken.None, out image, options);
    }

    internal static bool TryDecode(
        byte[]? encodedBytes,
        CancellationToken cancellationToken,
        out OfficeRasterImage? image,
        OfficeJpegDecodeOptions options = default) =>
        TryDecode(encodedBytes, cancellationToken, retainedManagedBytes: 0L, out image, options);

    internal static bool TryDecode(
        byte[]? encodedBytes,
        CancellationToken cancellationToken,
        long retainedManagedBytes,
        out OfficeRasterImage? image,
        OfficeJpegDecodeOptions options = default) {
        image = null;
        if (!IsJpeg(encodedBytes)) return false;
        try {
            byte[] pixels = OfficeJpegReader.DecodeRgba32(
                encodedBytes!, out int width, out int height, options, cancellationToken, retainedManagedBytes);
            image = OfficeRasterImage.FromOwnedRgba32(width, height, pixels);
            return true;
        } catch (Exception ex) when (ex is FormatException || ex is ArgumentException || ex is IndexOutOfRangeException || ex is OverflowException) {
            return false;
        }
    }

    /// <summary>Decodes JPEG bytes into an RGBA image.</summary>
    public static OfficeRasterImage Decode(byte[] encodedBytes, OfficeJpegDecodeOptions options = default) {
        if (encodedBytes == null) throw new ArgumentNullException(nameof(encodedBytes));
        byte[] pixels = OfficeJpegReader.DecodeRgba32(encodedBytes, out int width, out int height, options);
        return OfficeRasterImage.FromOwnedRgba32(width, height, pixels);
    }

    internal static bool TryDecodeColorComponents(
        byte[]? encodedBytes,
        int? requestedColorTransform,
        bool usePdfColorTransformDefault,
        out byte[] components,
        out int width,
        out int height,
        out int componentCount,
        OfficeJpegDecodeOptions options = default) {
        components = Array.Empty<byte>();
        width = 0;
        height = 0;
        componentCount = 0;
        if (!IsJpeg(encodedBytes)) return false;
        try {
            components = OfficeJpegReader.DecodeColorComponents(
                encodedBytes!,
                out width,
                out height,
                out componentCount,
                requestedColorTransform,
                usePdfColorTransformDefault,
                options);
            return true;
        } catch (Exception ex) when (ex is FormatException || ex is ArgumentException || ex is IndexOutOfRangeException || ex is OverflowException) {
            components = Array.Empty<byte>();
            width = 0;
            height = 0;
            componentCount = 0;
            return false;
        }
    }

    /// <summary>Encodes an RGBA image as JPEG bytes.</summary>
    public static byte[] Encode(OfficeRasterImage image, OfficeJpegEncodeOptions? options = null) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        OfficeJpegEncodeOptions effectiveOptions = options ?? new OfficeJpegEncodeOptions();
        byte[] rgba = image.PixelBuffer;
        if (HasTransparency(rgba)) {
            effectiveOptions = effectiveOptions.WithAdditionalRetainedManagedBytes(rgba.LongLength + 24L);
            rgba = (byte[])rgba.Clone();
            FlattenAlpha(rgba, effectiveOptions.Background);
        }
        return OfficeJpegWriter.WriteRgba(image.Width, image.Height, rgba, checked(image.Width * 4), effectiveOptions);
    }

    private static bool HasTransparency(byte[] rgba) {
        for (int index = 3; index < rgba.Length; index += 4) {
            if (rgba[index] != byte.MaxValue) return true;
        }
        return false;
    }

    private static void FlattenAlpha(byte[] rgba, OfficeColor background) {
        for (int i = 0; i < rgba.Length; i += 4) {
            int alpha = rgba[i + 3];
            if (alpha == 255) continue;
            int inverse = 255 - alpha;
            rgba[i] = Composite(rgba[i], background.R, alpha, inverse);
            rgba[i + 1] = Composite(rgba[i + 1], background.G, alpha, inverse);
            rgba[i + 2] = Composite(rgba[i + 2], background.B, alpha, inverse);
            rgba[i + 3] = 255;
        }
    }

    private static byte Composite(byte foreground, byte background, int alpha, int inverse) =>
        (byte)(((foreground * alpha) + (background * inverse) + 127) / 255);
}
