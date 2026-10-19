using System;
using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfExtractedImage {
    private byte[]? _drawingBytes;

    /// <summary>PDF image samples use their placement matrix, never the standalone JPEG container's EXIF orientation.</summary>
    /// <remarks>The private presentation copy retains all compressed samples and other metadata; extraction retains the original payload.</remarks>
    internal byte[] GetDrawingBytes(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        byte[]? cached = Volatile.Read(ref _drawingBytes);
        if (cached != null) return cached;
        byte[] presentation = _bytes;
        if (IsImageFile && _bytes.Length >= 2 && _bytes[0] == 0xFF && _bytes[1] == 0xD8 &&
            OfficeImageOrientationNormalizer.TryRead(_bytes, cancellationToken, out OfficeImageOrientation orientation) &&
            orientation != OfficeImageOrientation.Normal &&
            !OfficeImageOrientationNormalizer.TryCreateOrientationNeutralCopy(_bytes, cancellationToken, out presentation))
            throw new NotSupportedException("The PDF JPEG's EXIF orientation cannot be neutralized safely for drawing projection.");
        return Interlocked.CompareExchange(ref _drawingBytes, presentation, null) ?? presentation;
    }
}
