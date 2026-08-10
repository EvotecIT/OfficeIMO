using System;
using System.IO;
using System.Text;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Drawing;

/// <summary>Validates PNG chunk framing, ordering, and CRC integrity without decoding pixels.</summary>
internal static class OfficePngContainerValidator {
    private static readonly byte[] Signature = { 137, 80, 78, 71, 13, 10, 26, 10 };

    internal static bool TryValidate(byte[]? bytes, out int frameCount, out string? failureReason) {
        frameCount = 0;
        failureReason = null;
        if (bytes == null || bytes.Length < 33 || !HasSignature(bytes)) {
            failureReason = "PNG bytes are missing the PNG signature or required chunks.";
            return false;
        }

        try {
            OfficeRasterGuards.EnsurePayloadWithinLimits(bytes.Length, "PNG payload exceeds size limits.");
            bool seenHeader = false;
            bool seenImageData = false;
            bool imageDataEnded = false;
            bool seenAnimationControl = false;
            bool seenPalette = false;
            bool seenTransparency = false;
            bool seenBackground = false;
            bool seenHistogram = false;
            bool seenPhysicalDimensions = false;
            bool seenGamma = false;
            bool seenChromaticities = false;
            bool seenSignificantBits = false;
            bool seenStandardRgb = false;
            bool seenIccProfile = false;
            int bitDepth = 0;
            int colorType = 0;
            int paletteEntries = 0;
            int declaredFrameCount = 1;
            int offset = Signature.Length;
            while (offset + 12 <= bytes.Length) {
                int length = ReadBigEndianInt32(bytes, offset);
                long chunkEnd = (long)offset + 12L + length;
                if (length < 0 || chunkEnd > bytes.Length) {
                    failureReason = "PNG chunk length exceeds the available image bytes.";
                    return false;
                }

                string type = Encoding.ASCII.GetString(bytes, offset + 4, 4);
                if (!IsValidChunkType(bytes, offset + 4)) {
                    failureReason = "PNG bytes contain an invalid chunk type.";
                    return false;
                }
                if (!seenHeader && (type != "IHDR" || length != 13)) {
                    failureReason = "PNG bytes must start with an IHDR chunk.";
                    return false;
                }

                uint expectedCrc = ReadBigEndianUInt32(bytes, offset + 8 + length);
                uint actualCrc = ComputeCrc(bytes, offset + 4, 4 + length);
                if (actualCrc != expectedCrc) {
                    failureReason = "PNG chunk '" + type + "' has an invalid CRC.";
                    return false;
                }

                int dataOffset = offset + 8;
                switch (type) {
                    case "IHDR":
                        if (seenHeader || length != 13) {
                            failureReason = "PNG bytes contain an invalid or repeated IHDR chunk.";
                            return false;
                        }
                        int width = ReadBigEndianInt32(bytes, dataOffset);
                        int height = ReadBigEndianInt32(bytes, dataOffset + 4);
                        bitDepth = bytes[dataOffset + 8];
                        colorType = bytes[dataOffset + 9];
                        if (width <= 0 || height <= 0 ||
                            !IsValidColorLayout(colorType, bitDepth) ||
                            bytes[dataOffset + 10] != 0 ||
                            bytes[dataOffset + 11] != 0 ||
                            bytes[dataOffset + 12] > 1) {
                            failureReason = "PNG IHDR fields are invalid or unsupported.";
                            return false;
                        }
                        seenHeader = true;
                        break;
                    case "PLTE":
                        if (!seenHeader || seenImageData || seenPalette || seenTransparency ||
                            seenBackground || seenHistogram ||
                            length < 3 || length > 768 || length % 3 != 0 ||
                            colorType == 0 || colorType == 4) {
                            failureReason = "PNG bytes contain an invalid or misplaced PLTE chunk.";
                            return false;
                        }
                        paletteEntries = length / 3;
                        if (colorType == 3 && paletteEntries > 1 << bitDepth) {
                            failureReason = "PNG palette has more entries than its bit depth permits.";
                            return false;
                        }
                        seenPalette = true;
                        break;
                    case "tRNS":
                        if (!seenHeader || seenImageData || seenTransparency ||
                            (colorType == 0 && length != 2) ||
                            (colorType == 2 && length != 6) ||
                            (colorType == 3 && (!seenPalette || length == 0 || length > paletteEntries)) ||
                            colorType == 4 || colorType == 6 ||
                            !HasValidTransparencySamples(bytes, dataOffset, colorType, bitDepth)) {
                            failureReason = "PNG bytes contain an invalid or misplaced tRNS chunk.";
                            return false;
                        }
                        seenTransparency = true;
                        break;
                    case "bKGD":
                        if (!seenHeader || seenImageData || seenBackground ||
                            !HasValidBackground(bytes, dataOffset, length, colorType, bitDepth, seenPalette, paletteEntries)) {
                            failureReason = "PNG bytes contain an invalid or misplaced bKGD chunk.";
                            return false;
                        }
                        seenBackground = true;
                        break;
                    case "hIST":
                        if (!seenHeader || seenImageData || seenHistogram || colorType != 3 ||
                            !seenPalette || length != paletteEntries * 2) {
                            failureReason = "PNG bytes contain an invalid or misplaced hIST chunk.";
                            return false;
                        }
                        seenHistogram = true;
                        break;
                    case "acTL":
                        if (!seenHeader || seenImageData || seenAnimationControl || length != 8) {
                            failureReason = "PNG bytes contain an invalid APNG animation-control chunk.";
                            return false;
                        }
                        int candidate = ReadBigEndianInt32(bytes, dataOffset);
                        if (candidate <= 0) {
                            failureReason = "PNG animation frame count must be positive.";
                            return false;
                        }
                        declaredFrameCount = candidate;
                        seenAnimationControl = true;
                        break;
                    case "pHYs":
                        if (!seenHeader || seenImageData || seenPhysicalDimensions || length != 9 ||
                            bytes[dataOffset + 8] > 1) {
                            failureReason = "PNG bytes contain an invalid or misplaced pHYs chunk.";
                            return false;
                        }
                        seenPhysicalDimensions = true;
                        break;
                    case "gAMA":
                        if (!seenHeader || seenPalette || seenImageData || seenGamma || length != 4 ||
                            ReadBigEndianUInt32(bytes, dataOffset) == 0) {
                            failureReason = "PNG bytes contain an invalid or misplaced gAMA chunk.";
                            return false;
                        }
                        seenGamma = true;
                        break;
                    case "cHRM":
                        if (!seenHeader || seenPalette || seenImageData || seenChromaticities || length != 32) {
                            failureReason = "PNG bytes contain an invalid or misplaced cHRM chunk.";
                            return false;
                        }
                        seenChromaticities = true;
                        break;
                    case "sBIT":
                        if (!seenHeader || seenPalette || seenImageData || seenSignificantBits ||
                            !HasValidSignificantBits(bytes, dataOffset, length, colorType, bitDepth)) {
                            failureReason = "PNG bytes contain an invalid or misplaced sBIT chunk.";
                            return false;
                        }
                        seenSignificantBits = true;
                        break;
                    case "sRGB":
                        if (!seenHeader || seenPalette || seenImageData || seenStandardRgb || seenIccProfile || length != 1 ||
                            bytes[dataOffset] > 3) {
                            failureReason = "PNG bytes contain an invalid or misplaced sRGB chunk.";
                            return false;
                        }
                        seenStandardRgb = true;
                        break;
                    case "iCCP":
                        if (!seenHeader || seenPalette || seenImageData || seenIccProfile || seenStandardRgb ||
                            !HasValidIccProfile(bytes, dataOffset, length)) {
                            failureReason = "PNG bytes contain an invalid or misplaced iCCP chunk.";
                            return false;
                        }
                        seenIccProfile = true;
                        break;
                    case "IDAT":
                        if (!seenHeader || imageDataEnded || (colorType == 3 && !seenPalette)) {
                            failureReason = "PNG image data is misplaced or its required palette is missing.";
                            return false;
                        }
                        seenImageData = true;
                        break;
                    case "IEND":
                        if (length != 0 || !seenImageData) {
                            failureReason = "PNG bytes contain an invalid IEND chunk or no image data.";
                            return false;
                        }
                        offset = (int)chunkEnd;
                        if (offset != bytes.Length) {
                            failureReason = "PNG bytes contain trailing data after IEND.";
                            return false;
                        }
                        frameCount = declaredFrameCount;
                        return true;
                    default:
                        if (IsCriticalChunk(bytes[offset + 4])) {
                            failureReason = "PNG bytes contain the unknown critical chunk '" + type + "'.";
                            return false;
                        }
                        if (seenImageData) imageDataEnded = true;
                        break;
                }

                offset = (int)chunkEnd;
            }
        } catch (Exception exception) when (exception is FormatException || exception is OverflowException) {
            failureReason = exception.Message;
            return false;
        }

        failureReason = "PNG bytes do not contain a complete IEND chunk.";
        return false;
    }

    private static bool HasValidTransparencySamples(byte[] bytes, int offset, int colorType, int bitDepth) {
        if (colorType == 3 || bitDepth == 16) return true;
        int maximumSample = (1 << bitDepth) - 1;
        int sampleCount = colorType == 0 ? 1 : 3;
        for (int index = 0; index < sampleCount; index++) {
            int sampleOffset = offset + index * 2;
            int sample = bytes[sampleOffset] << 8 | bytes[sampleOffset + 1];
            if (sample > maximumSample) return false;
        }
        return true;
    }

    private static bool HasValidBackground(
        byte[] bytes,
        int offset,
        int length,
        int colorType,
        int bitDepth,
        bool seenPalette,
        int paletteEntries) {
        if (colorType == 3) {
            return seenPalette && length == 1 && bytes[offset] < paletteEntries;
        }
        int expectedSamples = colorType == 0 || colorType == 4 ? 1 : 3;
        if (length != expectedSamples * 2 || bitDepth == 16) return length == expectedSamples * 2;
        int maximumSample = (1 << bitDepth) - 1;
        for (int index = 0; index < expectedSamples; index++) {
            int sampleOffset = offset + index * 2;
            int sample = bytes[sampleOffset] << 8 | bytes[sampleOffset + 1];
            if (sample > maximumSample) return false;
        }
        return true;
    }

    private static bool HasValidSignificantBits(
        byte[] bytes,
        int offset,
        int length,
        int colorType,
        int bitDepth) {
        int expectedLength;
        switch (colorType) {
            case 0: expectedLength = 1; break;
            case 2:
            case 3: expectedLength = 3; break;
            case 4: expectedLength = 2; break;
            case 6: expectedLength = 4; break;
            default: return false;
        }
        if (length != expectedLength) return false;
        int maximum = colorType == 3 ? 8 : bitDepth;
        for (int index = 0; index < length; index++) {
            if (bytes[offset + index] == 0 || bytes[offset + index] > maximum) return false;
        }
        return true;
    }

    private static bool HasValidIccProfile(byte[] bytes, int offset, int length) {
        if (length < 9) return false;
        int nameLength = 0;
        while (nameLength < length && nameLength < 80 && bytes[offset + nameLength] != 0) {
            byte value = bytes[offset + nameLength];
            if (!IsValidKeywordByte(value) ||
                (value == (byte)' ' && (nameLength == 0 || bytes[offset + nameLength - 1] == (byte)' '))) {
                return false;
            }
            nameLength++;
        }
        if (nameLength < 1 || nameLength > 79 || nameLength >= length ||
            bytes[offset + nameLength - 1] == (byte)' ') {
            return false;
        }

        int compressionMethodOffset = offset + nameLength + 1;
        if (compressionMethodOffset >= offset + length || bytes[compressionMethodOffset] != 0) return false;
        int compressedOffset = compressionMethodOffset + 1;
        int compressedLength = offset + length - compressedOffset;
        if (compressedLength < 6) return false;
        var compressed = new byte[compressedLength];
        Buffer.BlockCopy(bytes, compressedOffset, compressed, 0, compressedLength);
        try {
            return OfficeZlibCodec.Decompress(
                compressed,
                OfficeRasterGuards.MaximumEncodedBytes).Length > 0;
        } catch (Exception exception) when (
            exception is ArgumentException ||
            exception is FormatException ||
            exception is InvalidDataException ||
            exception is NotSupportedException ||
            exception is OverflowException) {
            return false;
        }
    }

    private static bool IsValidKeywordByte(byte value) =>
        value >= 32 && value <= 126 || value >= 161;

    private static bool IsValidColorLayout(int colorType, int bitDepth) {
        switch (colorType) {
            case 0:
                return bitDepth == 1 || bitDepth == 2 || bitDepth == 4 || bitDepth == 8 || bitDepth == 16;
            case 2:
            case 4:
            case 6:
                return bitDepth == 8 || bitDepth == 16;
            case 3:
                return bitDepth == 1 || bitDepth == 2 || bitDepth == 4 || bitDepth == 8;
            default:
                return false;
        }
    }

    private static bool IsValidChunkType(byte[] bytes, int offset) {
        for (int index = 0; index < 4; index++) {
            byte value = bytes[offset + index];
            if (!((value >= (byte)'A' && value <= (byte)'Z') ||
                  (value >= (byte)'a' && value <= (byte)'z'))) return false;
        }
        // The third chunk-type byte is reserved by PNG and must be uppercase.
        return bytes[offset + 2] >= (byte)'A' && bytes[offset + 2] <= (byte)'Z';
    }

    private static bool IsCriticalChunk(byte firstTypeByte) =>
        firstTypeByte >= (byte)'A' && firstTypeByte <= (byte)'Z';

    private static bool HasSignature(byte[] bytes) {
        for (int index = 0; index < Signature.Length; index++) {
            if (bytes[index] != Signature[index]) return false;
        }
        return true;
    }

    private static uint ComputeCrc(byte[] bytes, int offset, int count) {
        uint crc = 0xFFFFFFFFU;
        for (int index = 0; index < count; index++) {
            crc ^= bytes[offset + index];
            for (int bit = 0; bit < 8; bit++) {
                crc = (crc & 1U) != 0 ? 0xEDB88320U ^ (crc >> 1) : crc >> 1;
            }
        }
        return crc ^ 0xFFFFFFFFU;
    }

    private static int ReadBigEndianInt32(byte[] bytes, int offset) =>
        (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

    private static uint ReadBigEndianUInt32(byte[] bytes, int offset) =>
        ((uint)bytes[offset] << 24) | ((uint)bytes[offset + 1] << 16) | ((uint)bytes[offset + 2] << 8) | bytes[offset + 3];
}
