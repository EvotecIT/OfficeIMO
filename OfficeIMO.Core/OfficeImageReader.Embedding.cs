using System;
using System.IO;

namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    /// <summary>Converts HEIC still images to PNG for Office document embedding, leaving other recognized formats unchanged.</summary>
    /// <remarks>Uses an explicit request codec or the registered platform fallback. The returned array is the
    /// original caller-owned input for other formats, and a new PNG for HEIC. No system codec is redistributed.</remarks>
    /// <exception cref="NotSupportedException">The required system codec is unavailable or the HEIC profile is unsupported.</exception>
    public static byte[] NormalizeForEmbedding(byte[] encodedBytes, OfficeRasterDecodeOptions? options, out OfficeImageInfo info) {
        if (encodedBytes == null) throw new ArgumentNullException(nameof(encodedBytes));
        var effective = options ?? new OfficeRasterDecodeOptions();
        effective.Validate();
        effective.CancellationToken.ThrowIfCancellationRequested();
        if (encodedBytes.Length > effective.MaximumEncodedBytes ||
            !TryIdentifyByContent(encodedBytes, null, effective.CancellationToken, out info))
            throw new InvalidDataException("The image is malformed, unsupported, or exceeds the encoded-byte limit.");
        if (info.Format != OfficeImageFormat.Heic) return encodedBytes;
        if (!OfficeImagePngConverter.TryConvertToPng(encodedBytes, effective, out byte[] png, out var diagnostic)) {
            if (diagnostic.CodecUnavailableReason != OfficeRasterCodecUnavailableReason.None)
                throw new OfficeRasterDecodeException(diagnostic);
            string code = diagnostic.DiagnosticCode ?? "image-decode-failed";
            throw new NotSupportedException(code + " [" + diagnostic.CodecUnavailableReason + "]: " + diagnostic.Diagnostic);
        }
        effective.CancellationToken.ThrowIfCancellationRequested();
        if (!TryIdentifyByContent(png, null, effective.CancellationToken, out info))
            throw new InvalidDataException("The normalized HEIC image is not a valid PNG.");
        return png;
    }
}
