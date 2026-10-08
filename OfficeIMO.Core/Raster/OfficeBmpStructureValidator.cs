using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Validates complete BMP storage for lossless metadata operations without requiring a managed pixel decoder.</summary>
internal static class OfficeBmpStructureValidator {
    internal readonly struct Layout {
        internal Layout(int pixelOffset, long pixelLength, int profileOffset, int profileLength, bool embeddedProfile) {
            PixelOffset = pixelOffset; PixelLength = pixelLength; ProfileOffset = profileOffset; ProfileLength = profileLength; EmbeddedProfile = embeddedProfile;
        }
        internal int PixelOffset { get; }
        internal long PixelLength { get; }
        internal int ProfileOffset { get; }
        internal int ProfileLength { get; }
        internal bool EmbeddedProfile { get; }
    }

    internal static bool TryValidate(byte[] bytes, CancellationToken token, out Layout layout, long additionallyRetainedBytes = 0L) {
        layout = default;
        token.ThrowIfCancellationRequested();
        if (!OfficeRasterGuards.IsEncodedPayloadWithinLimits(bytes.Length) ||
            !OfficeImageReader.TryIdentifyByContent(bytes, null, token, out OfficeImageInfo info) || info.Format != OfficeImageFormat.Bmp) return false;
        try {
            if (Read(bytes, 2, 4) != bytes.Length || Read(bytes, 6, 2) != 0 || Read(bytes, 8, 2) != 0) return false;
            int header = checked((int)Read(bytes, 14, 4));
            int pixels = checked((int)Read(bytes, 10, 4));
            bool core = header == 12;
            int bits = (int)Read(bytes, core ? 24 : 28, 2);
            int compression = core ? 0 : checked((int)Read(bytes, 30, 4));
            uint imageSize = core ? 0 : Read(bytes, 34, 4);
            uint colors = bits > 0 && bits <= 8 ? (uint)(1 << bits) : 0;
            if (!core && Read(bytes, 46, 4) != 0) colors = Read(bytes, 46, 4);
            if (bits > 0 && bits <= 8 && colors > (1U << bits)) return false;
            long colorTable = 14L + header;
            if (compression == 3 || compression == 6) {
                int maskCount = compression == 6 || header >= 56 ? 4 : 3;
                if (header != 40 && header < 52 || 54L + maskCount * 4 > bytes.Length) return false;
                uint occupied = 0;
                for (int index = 0; index < maskCount; index++) {
                    uint mask = Read(bytes, 54 + index * 4, 4);
                    if (index < 3 && mask == 0 || bits == 16 && (mask & 0xFFFF0000U) != 0 || (occupied & mask) != 0) return false;
                    if (mask != 0) { uint shifted = mask; while ((shifted & 1) == 0) shifted >>= 1; if ((shifted & (shifted + 1U)) != 0) return false; }
                    occupied |= mask;
                }
                if (header == 40) colorTable += maskCount * 4;
            }
            long metadataEnd = checked(colorTable + colors * (core ? 3L : 4L));
            if (pixels < metadataEnd || pixels >= bytes.Length) return false;
            long pixelLength;
            if (compression == 0 || compression == 3 || compression == 6) {
                pixelLength = checked(((info.Width * (long)bits + 31L) / 32L) * 4L * info.Height);
                if (pixelLength <= 0 || pixelLength > bytes.Length - (long)pixels || imageSize != 0 && (imageSize < pixelLength || imageSize > bytes.Length - (long)pixels)) return false;
                if (bits <= 8 && !ValidateIndices(bytes, pixels, info.Width, info.Height, bits, colors, token)) return false;
            } else if (compression == 1 || compression == 2) {
                long available = bytes.Length - (long)pixels;
                if (header == 124 && Read(bytes, 70, 4) is 0x4C494E4B or 0x4D424544) {
                    long profileStart = 14L + Read(bytes, 126, 4);
                    if (profileStart > pixels && profileStart < bytes.Length) available = profileStart - pixels;
                }
                if (imageSize > available) return false;
                int end = checked(pixels + (int)(imageSize == 0 ? available : imageSize));
                if (!ValidateRle(bytes, pixels, end, info.Width, info.Height, bits, colors, token, out int consumed)) return false;
                pixelLength = imageSize == 0 ? consumed : imageSize;
            } else {
                if (imageSize == 0 || imageSize > bytes.Length - (long)pixels ||
                    !ValidateEmbedded(bytes, pixels, checked((int)imageSize), compression, info, token, additionallyRetainedBytes)) return false;
                pixelLength = imageSize;
            }
            if (!OfficeBitmapV5ProfileValidator.TryValidate(bytes, 14, header, pixels, pixelLength, bytes.Length, out int profileOffset, out int profileLength) ||
                profileLength > OfficeExifProfileCodec.MaximumProfileBytes || profileLength != 0 && profileOffset < metadataEnd) return false;
            layout = new Layout(pixels, pixelLength, profileOffset, profileLength, header == 124 && Read(bytes, 70, 4) == 0x4D424544);
            token.ThrowIfCancellationRequested();
            return true;
        } catch (OverflowException) {
            return false;
        }
    }

    private static bool ValidateIndices(byte[] bytes, int pixels, int width, int height, int bits, uint colors, CancellationToken token) {
        if (colors == (1U << bits)) return true;
        long stride = ((width * (long)bits + 31L) / 32L) * 4L;
        for (int y = 0; y < height; y++) {
            token.ThrowIfCancellationRequested();
            int row = checked((int)(pixels + y * stride));
            for (int x = 0; x < width; x++) {
                if ((x & 4095) == 0) token.ThrowIfCancellationRequested();
                int bit = x * bits; int index = bytes[row + bit / 8] >> (8 - bits - bit % 8) & ((1 << bits) - 1);
                if (index >= colors) return false;
            }
        }
        return true;
    }

    private static bool ValidateRle(byte[] bytes, int start, int end, int width, int height, int bits, uint colors, CancellationToken token, out int consumed) {
        consumed = 0;
        int x = 0, y = 0, at = start, records = 0;
        while (at < end) {
            if ((records++ & 1023) == 0) token.ThrowIfCancellationRequested();
            if (at > end - 2) return false;
            int count = bytes[at++], value = bytes[at++];
            if (count != 0) {
                if (y >= height || count > width - x || (bits == 8 ? value >= colors : (value >> 4) >= colors || count > 1 && (value & 15) >= colors)) return false;
                x += count;
            } else if (value == 0) { x = 0; if (++y > height) return false; }
            else if (value == 1) { consumed = at - start; return true; }
            else if (value == 2) {
                if (at > end - 2 || bytes[at] > width - x || bytes[at + 1] >= height - y) return false;
                x += bytes[at++]; y += bytes[at++];
            } else {
                int payload = bits == 8 ? value : (value + 1) / 2; int padded = payload + (payload & 1);
                if (y >= height || value > width - x || padded > end - at) return false;
                for (int index = 0; index < value; index++) {
                    int color = bits == 8 ? bytes[at + index] : bytes[at + index / 2] >> (index % 2 == 0 ? 4 : 0) & 15;
                    if (color >= colors) return false;
                }
                at += padded; x += value;
            }
        }
        return false;
    }

    private static bool ValidateEmbedded(byte[] bytes, int at, int length, int compression, OfficeImageInfo bitmap, CancellationToken token, long additionallyRetainedBytes) {
        // Existing JPEG/PNG structure owners operate on complete arrays. Reserve the
        // borrowed BMP plus the independent payload and bounded validation workspace.
        if (additionallyRetainedBytes < 0L || checked(bytes.LongLength + length * 2L + 65536L + additionallyRetainedBytes) > OfficeRasterGuards.MaximumDecodedBytes) return false;
        var payload = new byte[length]; Buffer.BlockCopy(bytes, at, payload, 0, length);
        if (!OfficeImageReader.TryIdentifyByContent(payload, null, token, out OfficeImageInfo info) || info.Width != bitmap.Width || info.Height != bitmap.Height) return false;
        return compression == 4
            ? info.Format == OfficeImageFormat.Jpeg && OfficeImageReader.HasCompleteJpegPayload(payload, token, requireManagedFrame: false, validateMetadata: true)
            : info.Format == OfficeImageFormat.Png && OfficePngContainerValidator.TryValidate(payload, token, out _, out _);
    }

    private static uint Read(byte[] bytes, int at, int length) => (uint)OfficeExifProfileCodec.Read(bytes, at, length, true);
}
