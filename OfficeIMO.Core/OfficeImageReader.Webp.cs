using System;
using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    private const int MaximumWebpExifBytes = 1024 * 1024;

    internal static bool TryValidateWebpContainer(byte[] data) => TryReadWebp(data, out _);

    private static bool TryReadWebp(
        byte[] data,
        out OfficeImageInfo info,
        bool validateDecodedAlpha = false,
        OfficeRasterImage? decodedImage = null,
        CancellationToken cancellationToken = default) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 20 ||
            GetAscii(data, 0, 4) != "RIFF" ||
            GetAscii(data, 8, 4) != "WEBP") {
            return false;
        }

        long containerLength = 8L + ReadUInt32LittleEndian(data, 4);
        if (containerLength != data.LongLength) return false;

        int width = 0;
        int height = 0;
        int imageWidth = 0;
        int imageHeight = 0;
        bool extended = false;
        bool hasImage = false;
        bool hasAlpha = false;
        bool alphaSemanticsKnown = false;
        bool seenAnimationControl = false;
        bool hasAnimationFrame = false;
        bool seenAlphaChunk = false;
        bool seenIccProfile = false;
        bool seenXmp = false;
        byte extendedFlags = 0;
        int exifOffset = 0;
        int exifLength = 0;
        int offset = 12;
        while (offset < data.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            if (offset > data.Length - 8) return false;
            string chunkType = GetAscii(data, offset, 4);
            uint declaredChunkSize = ReadUInt32LittleEndian(data, offset + 4);
            if (declaredChunkSize > int.MaxValue) return false;

            int chunkSize = (int)declaredChunkSize;
            int chunkDataOffset = checked(offset + 8);
            long chunkDataEnd = (long)chunkDataOffset + chunkSize;
            long paddedChunkEnd = chunkDataEnd + (chunkSize & 1);
            if (chunkDataEnd > containerLength ||
                paddedChunkEnd > containerLength ||
                (chunkSize & 1) != 0 && data[(int)chunkDataEnd] != 0) {
                return false;
            }

            if (chunkType == "VP8X") {
                if (offset != 12 || extended || chunkSize != 10) return false;
                extended = true;
                extendedFlags = data[chunkDataOffset];
                if ((extendedFlags & 0xC1) != 0 ||
                    data[chunkDataOffset + 1] != 0 ||
                    data[chunkDataOffset + 2] != 0 ||
                    data[chunkDataOffset + 3] != 0) {
                    return false;
                }

                width = 1 + ReadUInt24LittleEndian(data, chunkDataOffset + 4);
                height = 1 + ReadUInt24LittleEndian(data, chunkDataOffset + 7);
            } else if (chunkType == "VP8L") {
                if (hasImage || seenAnimationControl || seenXmp || exifOffset != 0) {
                    return false;
                }
                if (seenAlphaChunk || !TryReadWebpImageHeader(
                        data, chunkDataOffset, chunkSize, "VP8L", out imageWidth, out imageHeight, out bool imageHasAlpha)) {
                    return false;
                }
                hasImage = true;
                hasAlpha = imageHasAlpha;
                if (extended && validateDecodedAlpha && decodedImage != null) {
                    hasAlpha = HasWebpPixelTransparency(decodedImage.PixelBuffer);
                    alphaSemanticsKnown = true;
                }
            } else if (chunkType == "VP8 ") {
                if (hasImage || seenAnimationControl || seenXmp || exifOffset != 0 || !TryReadWebpImageHeader(
                        data, chunkDataOffset, chunkSize, "VP8 ", out imageWidth, out imageHeight, out _)) {
                    return false;
                }
                hasImage = true;
                hasAlpha = seenAlphaChunk;
                alphaSemanticsKnown = true;
            } else if (chunkType == "ICCP") {
                if (!extended || seenIccProfile || hasImage || seenAnimationControl || hasAnimationFrame ||
                    seenAlphaChunk || exifOffset != 0 || seenXmp ||
                    !OfficeIccProfileValidator.TryValidate(data, chunkDataOffset, chunkSize)) {
                    return false;
                }
                seenIccProfile = true;
            } else if (chunkType == "ALPH") {
                if (!extended || seenAlphaChunk || hasImage || seenAnimationControl || hasAnimationFrame ||
                    exifOffset != 0 || seenXmp || !HasValidWebpAlphaHeader(data, chunkDataOffset, chunkSize) ||
                    (extendedFlags & 0x02) != 0) {
                    return false;
                }
                seenAlphaChunk = true;
            } else if (chunkType == "ANIM") {
                if (!extended || seenAnimationControl || hasImage || seenAlphaChunk ||
                    exifOffset != 0 || seenXmp || chunkSize != 6) return false;
                seenAnimationControl = true;
            } else if (chunkType == "ANMF") {
                if (!extended || !seenAnimationControl || hasImage || exifOffset != 0 || seenXmp ||
                    !TryReadWebpAnimationFrame(
                        data, chunkDataOffset, chunkSize, width, height,
                        validateDecodedAlpha,
                        out bool frameHasAlpha,
                        out bool frameAlphaSemanticsKnown,
                        cancellationToken)) {
                    return false;
                }
                hasAlpha |= frameHasAlpha;
                alphaSemanticsKnown = !hasAnimationFrame
                    ? frameAlphaSemanticsKnown
                    : alphaSemanticsKnown && frameAlphaSemanticsKnown;
                hasAnimationFrame = true;
            } else if (chunkType == "EXIF") {
                if (!extended || exifOffset != 0 || (!hasImage && !hasAnimationFrame) || seenXmp ||
                    !HasValidWebpExif(data, chunkDataOffset, chunkSize, cancellationToken)) {
                    return false;
                }
                exifOffset = chunkDataOffset;
                exifLength = chunkSize;
            } else if (chunkType == "XMP ") {
                if (!extended || seenXmp || (!hasImage && !hasAnimationFrame) ||
                    !OfficeXmpPacketValidator.TryValidate(data, chunkDataOffset, chunkSize)) return false;
                seenXmp = true;
            }

            offset = (int)paddedChunkEnd;
        }

        if (offset != data.Length) return false;
        if (extended) {
            bool declaresAnimation = (extendedFlags & 0x02) != 0;
            if (declaresAnimation) {
                if (hasImage || !seenAnimationControl || !hasAnimationFrame) return false;
            } else if (!hasImage || seenAnimationControl || hasAnimationFrame ||
                       width != imageWidth || height != imageHeight) {
                return false;
            }
            if (((extendedFlags & 0x20) != 0) != seenIccProfile ||
                alphaSemanticsKnown && ((extendedFlags & 0x10) != 0) != hasAlpha ||
                ((extendedFlags & 0x08) != 0) != (exifOffset != 0) ||
                ((extendedFlags & 0x04) != 0) != seenXmp) {
                return false;
            }
        } else {
            if (!hasImage) return false;
            width = imageWidth;
            height = imageHeight;
        }

        double dpiX = 96D;
        double dpiY = 96D;
        if (exifOffset != 0 &&
            TryReadWebpExif(
                data, exifOffset, exifLength, width, height,
                out OfficeImageInfo exifInfo,
                cancellationToken)) {
            dpiX = exifInfo.DpiX;
            dpiY = exifInfo.DpiY;
        }

        info = new OfficeImageInfo(OfficeImageFormat.Webp, width, height, dpiX, dpiY);
        return OfficeRasterGuards.TryEnsurePixelCount(width, height, out _);
    }

    private static bool TryReadWebpAnimationFrame(
        byte[] data,
        int offset,
        int length,
        int canvasWidth,
        int canvasHeight,
        bool validateDecodedAlpha,
        out bool hasAlpha,
        out bool alphaSemanticsKnown,
        CancellationToken cancellationToken) {
        hasAlpha = false;
        alphaSemanticsKnown = false;
        if (length < 24) return false;

        int frameX = checked(ReadUInt24LittleEndian(data, offset) * 2);
        int frameY = checked(ReadUInt24LittleEndian(data, offset + 3) * 2);
        int frameWidth = checked(ReadUInt24LittleEndian(data, offset + 6) + 1);
        int frameHeight = checked(ReadUInt24LittleEndian(data, offset + 9) + 1);
        if ((data[offset + 15] & 0xFC) != 0 ||
            (long)frameX + frameWidth > canvasWidth ||
            (long)frameY + frameHeight > canvasHeight) {
            return false;
        }

        bool seenAlpha = false;
        bool seenImage = false;
        int frameEnd = checked(offset + length);
        int chunkOffset = offset + 16;
        while (chunkOffset < frameEnd) {
            cancellationToken.ThrowIfCancellationRequested();
            if (chunkOffset > frameEnd - 8) return false;
            string chunkType = GetAscii(data, chunkOffset, 4);
            uint declaredChunkSize = ReadUInt32LittleEndian(data, chunkOffset + 4);
            if (declaredChunkSize > int.MaxValue) return false;

            int chunkSize = (int)declaredChunkSize;
            int chunkDataOffset = checked(chunkOffset + 8);
            long chunkDataEnd = (long)chunkDataOffset + chunkSize;
            long paddedChunkEnd = chunkDataEnd + (chunkSize & 1);
            if (chunkDataEnd > frameEnd || paddedChunkEnd > frameEnd ||
                (chunkSize & 1) != 0 && data[(int)chunkDataEnd] != 0) {
                return false;
            }

            if (chunkType == "ALPH") {
                if (seenAlpha || seenImage || !HasValidWebpAlphaHeader(data, chunkDataOffset, chunkSize)) return false;
                seenAlpha = true;
            } else if (chunkType == "VP8 " || chunkType == "VP8L") {
                if (seenImage || seenAlpha && chunkType == "VP8L" || !TryReadWebpImageHeader(
                        data, chunkDataOffset, chunkSize, chunkType,
                        out int imageWidth, out int imageHeight, out _) ||
                    imageWidth != frameWidth || imageHeight != frameHeight) {
                    return false;
                }
                seenImage = true;
                if (chunkType == "VP8 ") {
                    hasAlpha = seenAlpha;
                    alphaSemanticsKnown = true;
                } else if (validateDecodedAlpha && TryDecodeStandaloneWebpChunk(
                               data,
                               chunkOffset,
                               (int)paddedChunkEnd - chunkOffset,
                               out OfficeRasterImage? decoded) &&
                           decoded != null) {
                    hasAlpha = HasWebpPixelTransparency(decoded.PixelBuffer);
                    alphaSemanticsKnown = true;
                }
            }

            chunkOffset = (int)paddedChunkEnd;
        }

        return chunkOffset == frameEnd && seenImage;
    }

    private static bool TryDecodeStandaloneWebpChunk(
        byte[] source,
        int chunkOffset,
        int chunkLength,
        out OfficeRasterImage? image) {
        var wrapped = new byte[12 + chunkLength];
        wrapped[0] = (byte)'R';
        wrapped[1] = (byte)'I';
        wrapped[2] = (byte)'F';
        wrapped[3] = (byte)'F';
        int riffLength = wrapped.Length - 8;
        wrapped[4] = (byte)riffLength;
        wrapped[5] = (byte)(riffLength >> 8);
        wrapped[6] = (byte)(riffLength >> 16);
        wrapped[7] = (byte)(riffLength >> 24);
        wrapped[8] = (byte)'W';
        wrapped[9] = (byte)'E';
        wrapped[10] = (byte)'B';
        wrapped[11] = (byte)'P';
        Buffer.BlockCopy(source, chunkOffset, wrapped, 12, chunkLength);
        return OfficeWebpCodec.TryDecode(wrapped, out image);
    }

    private static bool HasWebpPixelTransparency(byte[] pixels) {
        for (int offset = 3; offset < pixels.Length; offset += 4) {
            if (pixels[offset] != byte.MaxValue) return true;
        }

        return false;
    }

    private static bool HasValidWebpAlphaHeader(byte[] data, int offset, int length) {
        if (length < 2) return false;
        byte control = data[offset];
        // Compression method 0 is the only defined value, preprocessing values 2-3 are reserved,
        // and the two high bits are reserved for future use.
        return (control & 0xC3) == 0 && (control & 0x30) <= 0x10;
    }

    private static bool TryReadWebpImageHeader(
        byte[] data,
        int offset,
        int length,
        string chunkType,
        out int width,
        out int height,
        out bool hasAlpha) {
        width = 0;
        height = 0;
        hasAlpha = false;
        if (chunkType == "VP8L") {
            if (length < 5 || data[offset] != 0x2F || (data[offset + 4] & 0xE0) != 0) return false;
            width = 1 + data[offset + 1] + ((data[offset + 2] & 0x3F) << 8);
            height = 1 + ((data[offset + 2] & 0xC0) >> 6) +
                     (data[offset + 3] << 2) +
                     ((data[offset + 4] & 0x0F) << 10);
            hasAlpha = (data[offset + 4] & 0x10) != 0;
            return true;
        }

        if (chunkType != "VP8 " || length < 10 ||
            data[offset + 3] != 0x9D || data[offset + 4] != 0x01 || data[offset + 5] != 0x2A) {
            return false;
        }
        width = ReadUInt16LittleEndian(data, offset + 6) & 0x3FFF;
        height = ReadUInt16LittleEndian(data, offset + 8) & 0x3FFF;
        return width > 0 && height > 0;
    }

    private static bool TryReadWebpExif(
        byte[] data,
        int offset,
        int length,
        int expectedWidth,
        int expectedHeight,
        out OfficeImageInfo info,
        CancellationToken cancellationToken) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (length < 8 || length > MaximumWebpExifBytes) return false;
        if (length >= 14 &&
            GetAscii(data, offset, 4) == "Exif" &&
            data[offset + 4] == 0 &&
            data[offset + 5] == 0) {
            offset += 6;
            length -= 6;
        }

        cancellationToken.ThrowIfCancellationRequested();
        byte[] tiff = new byte[length];
        Buffer.BlockCopy(data, offset, tiff, 0, length);
        return TryReadTiff(tiff, cancellationToken, out info) &&
               info.Width == expectedWidth &&
               info.Height == expectedHeight;
    }

    private static bool HasValidWebpExif(
        byte[] data,
        int offset,
        int length,
        CancellationToken cancellationToken) {
        if (length < 8 || length > MaximumWebpExifBytes) return false;
        if (length >= 14 &&
            GetAscii(data, offset, 4) == "Exif" &&
            data[offset + 4] == 0 &&
            data[offset + 5] == 0) {
            offset += 6;
            length -= 6;
        }

        return OfficeTiffStructureValidator.TryValidateExif(data, offset, length, cancellationToken);
    }

    private static uint ReadUInt32LittleEndian(byte[] data, int offset) =>
        offset + 4 <= data.Length
            ? (uint)(data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24))
            : 0U;
}
