using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    private static readonly byte[] IconPngSignature = { 137, 80, 78, 71, 13, 10, 26, 10 };

    private static bool HasCompleteIconPayload(byte[] data) {
        if (!TryReadIcon(data, out _)) return false;
        int count = ReadUInt16LittleEndian(data, 4);
        var validatedPayloads = new Dictionary<(int Offset, int Length), IconPayloadValidation>();
        long aggregateValidationBytes = 0;
        for (int index = 0; index < count; index++) {
            int entryOffset = 6 + index * 16;
            int declaredWidth = data[entryOffset] == 0 ? 256 : data[entryOffset];
            int declaredHeight = data[entryOffset + 1] == 0 ? 256 : data[entryOffset + 1];
            int imageLength = checked((int)ReadUInt32LittleEndian(data, entryOffset + 8));
            int imageOffset = checked((int)ReadUInt32LittleEndian(data, entryOffset + 12));
            var key = (imageOffset, imageLength);
            if (!validatedPayloads.TryGetValue(key, out IconPayloadValidation validation)) {
                if (aggregateValidationBytes > OfficeRasterGuards.MaximumEncodedBytes - imageLength) return false;
                aggregateValidationBytes += imageLength;
                var payload = new byte[imageLength];
                Buffer.BlockCopy(data, imageOffset, payload, 0, imageLength);
                if (HasIconPngSignature(payload)) {
                    bool valid = TryReadPng(payload, out OfficeImageInfo pngInfo) &&
                                 OfficePngReader.TryValidateDecodedPayload(payload);
                    validation = new IconPayloadValidation(valid, pngInfo.Width, pngInfo.Height);
                } else {
                    bool valid = TryReadCompleteIconDibPayload(payload, out int width, out int height);
                    validation = new IconPayloadValidation(valid, width, height);
                }
                validatedPayloads.Add(key, validation);
            }
            if (!validation.IsValid || validation.Width != declaredWidth || validation.Height != declaredHeight) {
                return false;
            }
        }
        return true;
    }

    private static bool HasIconPngSignature(byte[] payload) {
        if (payload.Length < IconPngSignature.Length) return false;
        for (int index = 0; index < IconPngSignature.Length; index++) {
            if (payload[index] != IconPngSignature[index]) return false;
        }
        return true;
    }

    private static bool TryReadCompleteIconDibPayload(byte[] payload, out int width, out int height) {
        width = 0;
        height = 0;
        if (payload.Length < 12) return false;
        int headerSize = ReadInt32LittleEndian(payload, 0);
        int storedHeight;
        int planes;
        int bitsPerPixel;
        int compression = 0;
        int imageSize = 0;
        int paletteEntrySize;
        long paletteEntries;
        long pixelOffset;

        if (headerSize == 12) {
            width = ReadUInt16LittleEndian(payload, 4);
            storedHeight = ReadUInt16LittleEndian(payload, 6);
            planes = ReadUInt16LittleEndian(payload, 8);
            bitsPerPixel = ReadUInt16LittleEndian(payload, 10);
            paletteEntrySize = 3;
            paletteEntries = bitsPerPixel <= 8 ? 1L << bitsPerPixel : 0L;
            pixelOffset = headerSize + paletteEntries * paletteEntrySize;
        } else {
            if (!OfficeDibHeaderLayout.IsSupportedWindowsInfoHeaderSize(headerSize) || headerSize > payload.Length) {
                return false;
            }
            width = ReadInt32LittleEndian(payload, 4);
            storedHeight = ReadInt32LittleEndian(payload, 8);
            planes = ReadUInt16LittleEndian(payload, 12);
            bitsPerPixel = ReadUInt16LittleEndian(payload, 14);
            compression = ReadInt32LittleEndian(payload, 16);
            imageSize = ReadInt32LittleEndian(payload, 20);
            uint colorsUsed = ReadUInt32LittleEndian(payload, 32);
            paletteEntrySize = 4;
            long maximumPaletteEntries = bitsPerPixel <= 8 ? 1L << bitsPerPixel : 0L;
            if (colorsUsed > maximumPaletteEntries && maximumPaletteEntries != 0) return false;
            paletteEntries = colorsUsed != 0 ? colorsUsed : maximumPaletteEntries;
            int externalMaskBytes = headerSize == 40 && compression == 3 ? 12 :
                headerSize == 40 && compression == 6 ? 16 : 0;
            pixelOffset = (long)headerSize + externalMaskBytes + paletteEntries * paletteEntrySize;
        }

        if (width <= 0 || storedHeight <= 0 || (storedHeight & 1) != 0 || planes != 1 ||
            (bitsPerPixel != 1 && bitsPerPixel != 4 && bitsPerPixel != 8 &&
             bitsPerPixel != 16 && bitsPerPixel != 24 && bitsPerPixel != 32) ||
            (compression != 0 && compression != 3 && compression != 6) ||
            ((compression == 3 || compression == 6) && bitsPerPixel != 16 && bitsPerPixel != 32)) {
            return false;
        }
        if (!HasValidIconBitfieldMasks(payload, headerSize, compression, bitsPerPixel)) return false;

        height = storedHeight / 2;
        if (!OfficeRasterGuards.TryEnsurePixelCount(width, height, out _) ||
            pixelOffset < headerSize || pixelOffset > payload.LongLength) {
            return false;
        }

        long xorStride = (((long)width * bitsPerPixel + 31L) / 32L) * 4L;
        long maskStride = (((long)width + 31L) / 32L) * 4L;
        long xorBytes = xorStride * height;
        long maskBytes = maskStride * height;
        bool hasMask = HasExactIconDibLayout(payload, headerSize, pixelOffset, xorBytes + maskBytes);
        bool hasOmittedMask = bitsPerPixel == 32 &&
                              HasExactIconDibLayout(payload, headerSize, pixelOffset, xorBytes);
        if ((!hasMask && !hasOmittedMask) || imageSize < 0 ||
            (imageSize != 0 && imageSize != xorBytes && !(hasMask && imageSize == xorBytes + maskBytes))) {
            return false;
        }
        return bitsPerPixel > 8 || HasValidIndexedIconPixels(
                   payload,
                   (int)pixelOffset,
                   width,
                   height,
                   bitsPerPixel,
                   (int)xorStride,
                   (int)paletteEntries);
    }

    private static bool HasExactIconDibLayout(
        byte[] payload,
        int headerSize,
        long pixelOffset,
        long pixelLength) {
        long pixelEnd = pixelOffset + pixelLength;
        if (pixelEnd > payload.LongLength) return false;
        if (!OfficeBitmapV5ProfileValidator.TryValidate(
                payload,
                0,
                headerSize,
                pixelOffset,
                pixelLength,
                payload.Length,
                out int profileOffset,
                out int profileSize)) return false;
        if (profileSize == 0) return pixelEnd == payload.LongLength;
        if (profileOffset < pixelEnd || profileOffset + profileSize != payload.Length) return false;
        for (long offset = pixelEnd; offset < profileOffset; offset++) {
            if (payload[(int)offset] != 0) return false;
        }
        return true;
    }

    private static bool HasValidIconBitfieldMasks(
        byte[] payload,
        int headerSize,
        int compression,
        int bitsPerPixel) {
        if (compression != 3 && compression != 6) return true;
        if (compression == 6 && headerSize == 52) return false;

        bool hasAlphaMask = compression == 6 || headerSize >= 56;
        int maskCount = hasAlphaMask ? 4 : 3;
        const int maskOffset = 40;
        if (payload.Length < maskOffset + maskCount * 4) return false;

        uint red = ReadUInt32LittleEndian(payload, maskOffset);
        uint green = ReadUInt32LittleEndian(payload, maskOffset + 4);
        uint blue = ReadUInt32LittleEndian(payload, maskOffset + 8);
        uint alpha = hasAlphaMask ? ReadUInt32LittleEndian(payload, maskOffset + 12) : 0;
        if (!IsValidIconBitfieldMask(red, bitsPerPixel) ||
            !IsValidIconBitfieldMask(green, bitsPerPixel) ||
            !IsValidIconBitfieldMask(blue, bitsPerPixel) ||
            (compression == 6 || alpha != 0) && !IsValidIconBitfieldMask(alpha, bitsPerPixel)) {
            return false;
        }
        return (red & green) == 0 && (red & blue) == 0 && (green & blue) == 0 &&
               (alpha == 0 || ((alpha & red) == 0 && (alpha & green) == 0 && (alpha & blue) == 0));
    }

    private static bool IsValidIconBitfieldMask(uint mask, int bitsPerPixel) {
        if (mask == 0 || bitsPerPixel < 32 && (mask >> bitsPerPixel) != 0) return false;
        while ((mask & 1) == 0) mask >>= 1;
        return (mask & (mask + 1)) == 0;
    }

    private static bool HasValidIndexedIconPixels(
        byte[] payload,
        int pixelOffset,
        int width,
        int height,
        int bitsPerPixel,
        int rowStride,
        int paletteEntries) {
        for (int y = 0; y < height; y++) {
            int rowOffset = pixelOffset + y * rowStride;
            for (int x = 0; x < width; x++) {
                int paletteIndex;
                if (bitsPerPixel == 8) {
                    paletteIndex = payload[rowOffset + x];
                } else if (bitsPerPixel == 4) {
                    byte packed = payload[rowOffset + x / 2];
                    paletteIndex = (x & 1) == 0 ? packed >> 4 : packed & 0x0F;
                } else {
                    byte packed = payload[rowOffset + x / 8];
                    paletteIndex = packed >> (7 - (x & 7)) & 0x01;
                }

                if (paletteIndex >= paletteEntries) return false;
            }
        }

        return true;
    }

    private readonly struct IconPayloadValidation {
        internal IconPayloadValidation(bool isValid, int width, int height) {
            IsValid = isValid;
            Width = width;
            Height = height;
        }

        internal bool IsValid { get; }
        internal int Width { get; }
        internal int Height { get; }
    }
}
