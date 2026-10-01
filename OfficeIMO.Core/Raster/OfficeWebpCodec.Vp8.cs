using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
    private static bool TryDecodeVp8(byte[]? encodedBytes, CancellationToken cancellationToken,
        long retainedManagedBytes, out OfficeRasterImage? image) {
        image = null;
        cancellationToken.ThrowIfCancellationRequested();
        if (encodedBytes == null || encodedBytes.Length > OfficeRasterGuards.MaximumEncodedBytes ||
            !IsWebp(encodedBytes) || ReadUInt32(encodedBytes, 4) != encodedBytes.Length - 8 ||
            !OfficeImageReader.TryIdentifyByContent(encodedBytes, null, cancellationToken, out OfficeImageInfo info) ||
            info.Format != OfficeImageFormat.Webp ||
            TryFindChunk(encodedBytes, "ANIM", cancellationToken, out _, out _) ||
            !TryFindChunk(encodedBytes, "VP8 ", cancellationToken, out int offset, out int length)) return false;
        if (!OfficeVp8Decoder.TryDecode(new OfficeByteView(encodedBytes).Slice(offset, length), cancellationToken,
                checked(retainedManagedBytes + encodedBytes.LongLength), out byte[] pixels, out int width, out int height) ||
            width != info.Width || height != info.Height) return false;
        if (TryFindChunk(encodedBytes, "ALPH", cancellationToken, out int alphaOffset, out int alphaLength) &&
            !TryApplyVp8Alpha(encodedBytes, alphaOffset, alphaLength, width, height, pixels,
                checked(retainedManagedBytes + encodedBytes.LongLength + pixels.LongLength + 24L), cancellationToken)) return false;
        image = OfficeRasterImage.FromOwnedRgba32(width, height, pixels);
        return true;
    }
}
