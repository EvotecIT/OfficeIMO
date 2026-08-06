namespace OfficeIMO.Drawing;

/// <summary>Shared dependency-free conversion helpers for raster formats supported by OfficeIMO.Drawing.</summary>
public static class OfficeImagePngConverter {
    /// <summary>Attempts to convert a Drawing-supported raster payload to PNG bytes.</summary>
    public static bool TryConvertToPng(byte[]? imageBytes, out byte[] pngBytes) {
        return TryConvertToPng(imageBytes, options: null, out pngBytes, out _);
    }

    /// <summary>Attempts to convert a selected static raster frame to PNG with typed loss evidence.</summary>
    public static bool TryConvertToPng(byte[]? imageBytes, OfficeRasterDecodeOptions? options, out byte[] pngBytes, out OfficeRasterDecodeInfo decodeInfo) {
        pngBytes = System.Array.Empty<byte>();
        if (!OfficeRasterImageDecoder.TryDecode(imageBytes, options, out OfficeRasterImage? image, out decodeInfo)) {
            if ((options?.FrameIndex ?? 0) != 0) return false;
            if (!OfficeDibReader.TryDecode(imageBytes, out image)) return false;
            decodeInfo = new OfficeRasterDecodeInfo(OfficeImageFormat.Unknown, 1, 0, succeeded: true, diagnostic: null);
        }

        if (image == null) return false;

        OfficeImageInfo? sourceInfo = null;
        if (imageBytes != null && OfficeImageReader.TryIdentify(imageBytes, null, out OfficeImageInfo identified)) {
            sourceInfo = identified;
        }

        pngBytes = sourceInfo == null
            ? OfficePngWriter.Encode(image)
            : OfficePngWriter.Encode(image, new OfficePngEncodeOptions {
                DpiX = sourceInfo.DpiX,
                DpiY = sourceInfo.DpiY
            });
        return true;
    }

    /// <summary>Attempts to convert an RTF-style raw DIB payload to PNG bytes.</summary>
    public static bool TryConvertDibToPng(byte[]? dibBytes, out byte[] pngBytes) {
        pngBytes = System.Array.Empty<byte>();
        if (!OfficeDibReader.TryDecode(dibBytes, out OfficeRasterImage? image)) return false;
        pngBytes = OfficePngWriter.Encode(image!);
        return true;
    }
}
