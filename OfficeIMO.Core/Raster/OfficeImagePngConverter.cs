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
        var effective = options ?? new OfficeRasterDecodeOptions();
        effective.Validate();
        System.Threading.CancellationToken cancellationToken = effective.CancellationToken;
        cancellationToken.ThrowIfCancellationRequested();
        if (!OfficeRasterImageDecoder.TryDecode(imageBytes, effective, out OfficeRasterImage? image, out decodeInfo)) {
            if (effective.FrameIndex != 0) return false;
            if (!OfficeDibReader.TryDecode(imageBytes, effective, out image)) return false;
            decodeInfo = new OfficeRasterDecodeInfo(OfficeImageFormat.Unknown, 1, 0, succeeded: true, diagnostic: null);
        }

        if (image == null) return false;

        OfficeImageInfo? sourceInfo = null;
        if (imageBytes != null && OfficeImageReader.TryIdentify(imageBytes, null, cancellationToken, out OfficeImageInfo identified)) {
            sourceInfo = identified;
        }

        double? selectedDpiX = decodeInfo.Format == OfficeImageFormat.Tiff
            ? decodeInfo.SelectedFrame?.DpiX
            : null;
        double? selectedDpiY = decodeInfo.Format == OfficeImageFormat.Tiff
            ? decodeInfo.SelectedFrame?.DpiY
            : null;
        var encodeOptions = new OfficePngEncodeOptions();
        if (decodeInfo.Format == OfficeImageFormat.Tiff && selectedDpiX.HasValue && selectedDpiY.HasValue) {
            encodeOptions.DpiX = selectedDpiX.Value;
            encodeOptions.DpiY = selectedDpiY.Value;
        } else if (decodeInfo.Format == OfficeImageFormat.Tiff) {
            encodeOptions.WritePhysicalResolution = false;
        } else if (sourceInfo != null) {
            encodeOptions.DpiX = sourceInfo.DpiX;
            encodeOptions.DpiY = sourceInfo.DpiY;
        } else {
            encodeOptions.WritePhysicalResolution = false;
        }
        using var output = new System.IO.MemoryStream();
        OfficePngWriter.EncodeTo(image, output, encodeOptions, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        pngBytes = CopyOutput(output, cancellationToken);
        return true;
    }

    internal static bool TryConvertToPng(
        byte[]? imageBytes,
        System.Threading.CancellationToken cancellationToken,
        out byte[] pngBytes) {
        return TryConvertToPng(
            imageBytes,
            new OfficeRasterDecodeOptions { CancellationToken = cancellationToken },
            out pngBytes,
            out _);
    }

    private static byte[] CopyOutput(
        System.IO.MemoryStream output,
        System.Threading.CancellationToken cancellationToken) {
        if (!output.TryGetBuffer(out System.ArraySegment<byte> segment)) return output.ToArray();
        var bytes = new byte[output.Length];
        const int chunkSize = 64 * 1024;
        for (int offset = 0; offset < bytes.Length; offset += chunkSize) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = System.Math.Min(chunkSize, bytes.Length - offset);
            System.Buffer.BlockCopy(segment.Array!, segment.Offset + offset, bytes, offset, count);
        }
        return bytes;
    }

    /// <summary>Attempts to convert an RTF-style raw DIB payload to PNG bytes.</summary>
    public static bool TryConvertDibToPng(byte[]? dibBytes, out byte[] pngBytes) {
        pngBytes = System.Array.Empty<byte>();
        if (!OfficeDibReader.TryDecode(dibBytes, out OfficeRasterImage? image)) return false;
        pngBytes = OfficePngWriter.Encode(image!);
        return true;
    }
}
