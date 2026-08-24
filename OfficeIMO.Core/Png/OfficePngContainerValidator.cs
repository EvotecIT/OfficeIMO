using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Drawing;

/// <summary>Validates PNG chunk framing, ordering, and CRC integrity without decoding pixels.</summary>
internal static class OfficePngContainerValidator {
    private const int MaximumPngTextBytes = 1024 * 1024;
    internal const long MaximumSuggestedPaletteMetadataBytes = 4L * 1024L * 1024L;
    private const long SuggestedPaletteEntryOverheadBytes = 256L;
    private static readonly byte[] Signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
    private static readonly UTF8Encoding StrictUtf8 = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

    internal static bool TryValidate(byte[]? bytes, out int frameCount, out string? failureReason) =>
        TryValidate(bytes, CancellationToken.None, out frameCount, out failureReason);

    internal static bool TryValidate(
        byte[]? bytes,
        CancellationToken cancellationToken,
        out int frameCount,
        out string? failureReason) {
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
            bool seenExif = false;
            bool seenModificationTime = false;
            var suggestedPaletteNames = new HashSet<string>(StringComparer.Ordinal);
            long suggestedPaletteMetadataBytes = 0;
            int bitDepth = 0;
            int colorType = 0;
            int paletteEntries = 0;
            int declaredFrameCount = 1;
            uint gammaValue = 0;
            bool hasStandardRgbChromaticities = false;
            long decodedTextBytes = 0;
            int offset = Signature.Length;
            while (offset + 12 <= bytes.Length) {
                cancellationToken.ThrowIfCancellationRequested();
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
                uint actualCrc = ComputeCrc(bytes, offset + 4, 4 + length, cancellationToken);
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
                        uint candidateGamma = length == 4 ? ReadBigEndianUInt32(bytes, dataOffset) : 0;
                        if (!seenHeader || seenPalette || seenImageData || seenGamma || length != 4 ||
                            candidateGamma is 0 or > int.MaxValue ||
                            seenStandardRgb && candidateGamma != 45455U) {
                            failureReason = "PNG bytes contain an invalid or misplaced gAMA chunk.";
                            return false;
                        }
                        gammaValue = candidateGamma;
                        seenGamma = true;
                        break;
                    case "cHRM":
                        if (!seenHeader || seenPalette || seenImageData || seenChromaticities || length != 32 ||
                            !HasValidChromaticities(bytes, dataOffset) ||
                            seenStandardRgb && !HasStandardRgbChromaticities(bytes, dataOffset)) {
                            failureReason = "PNG bytes contain an invalid or misplaced cHRM chunk.";
                            return false;
                        }
                        hasStandardRgbChromaticities = HasStandardRgbChromaticities(bytes, dataOffset);
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
                            bytes[dataOffset] > 3 ||
                            seenGamma && gammaValue != 45455U ||
                            seenChromaticities && !hasStandardRgbChromaticities) {
                            failureReason = "PNG bytes contain an invalid or misplaced sRGB chunk.";
                            return false;
                        }
                        seenStandardRgb = true;
                        break;
                    case "iCCP":
                        if (!seenHeader || seenPalette || seenImageData || seenIccProfile || seenStandardRgb ||
                            !HasValidIccProfile(bytes, dataOffset, length, cancellationToken)) {
                            failureReason = "PNG bytes contain an invalid or misplaced iCCP chunk.";
                            return false;
                        }
                        seenIccProfile = true;
                        break;
                    case "tEXt":
                        if (!seenHeader || !HasValidLatinText(
                                bytes, dataOffset, length, ref decodedTextBytes, cancellationToken)) {
                            failureReason = "PNG bytes contain an invalid tEXt chunk.";
                            return false;
                        }
                        if (seenImageData) imageDataEnded = true;
                        break;
                    case "zTXt":
                        if (!seenHeader || !HasValidCompressedText(
                                bytes, dataOffset, length, ref decodedTextBytes, cancellationToken)) {
                            failureReason = "PNG bytes contain an invalid zTXt chunk.";
                            return false;
                        }
                        if (seenImageData) imageDataEnded = true;
                        break;
                    case "iTXt":
                        if (!seenHeader || !HasValidInternationalText(
                                bytes, dataOffset, length, ref decodedTextBytes, cancellationToken)) {
                            failureReason = "PNG bytes contain an invalid iTXt chunk.";
                            return false;
                        }
                        if (seenImageData) imageDataEnded = true;
                        break;
                    case "eXIf":
                        if (!seenHeader || seenExif ||
                            !OfficeTiffStructureValidator.TryValidateExif(bytes, dataOffset, length)) {
                            failureReason = "PNG bytes contain an invalid or repeated eXIf chunk.";
                            return false;
                        }
                        seenExif = true;
                        if (seenImageData) imageDataEnded = true;
                        break;
                    case "tIME":
                        if (!seenHeader || seenModificationTime ||
                            !HasValidModificationTime(bytes, dataOffset, length)) {
                            failureReason = "PNG bytes contain an invalid or repeated tIME chunk.";
                            return false;
                        }
                        seenModificationTime = true;
                        if (seenImageData) imageDataEnded = true;
                        break;
                    case "sPLT":
                        if (!seenHeader || !HasValidSuggestedPalette(
                                bytes,
                                dataOffset,
                                length,
                                suggestedPaletteNames,
                                bytes.LongLength,
                                ref suggestedPaletteMetadataBytes)) {
                            failureReason = "PNG bytes contain an invalid or repeated sPLT chunk.";
                            return false;
                        }
                        if (seenImageData) imageDataEnded = true;
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

    private static bool HasValidSuggestedPalette(
        byte[] bytes,
        int offset,
        int length,
        HashSet<string> names,
        long encodedBytes,
        ref long metadataBytes) {
        if (!TryReadKeyword(bytes, offset, length, out int keywordEnd)) return false;
        int nameLength = keywordEnd - offset;
        int sampleDepthOffset = offset + nameLength + 1;
        if (sampleDepthOffset >= offset + length) return false;
        int sampleDepth = bytes[sampleDepthOffset];
        int entrySize = sampleDepth == 8 ? 6 : sampleDepth == 16 ? 10 : 0;
        int entriesLength = length - nameLength - 2;
        if (entrySize == 0 || entriesLength < entrySize || entriesLength % entrySize != 0 ||
            !TryReserveSuggestedPaletteName(encodedBytes, nameLength, ref metadataBytes)) return false;

        var nameCharacters = new char[nameLength];
        for (int index = 0; index < nameLength; index++) nameCharacters[index] = (char)bytes[offset + index];
        return names.Add(new string(nameCharacters));
    }

    internal static bool TryReserveSuggestedPaletteName(
        long encodedBytes,
        int nameLength,
        ref long metadataBytes) {
        if (encodedBytes < 0L || nameLength < 1 || nameLength > 79 || metadataBytes < 0L) return false;
        try {
            // Covers the retained UTF-16 string, HashSet slots/buckets, and transient resize slack.
            long reservation = checked(SuggestedPaletteEntryOverheadBytes + nameLength * 4L);
            long updatedMetadataBytes = checked(metadataBytes + reservation);
            if (updatedMetadataBytes > MaximumSuggestedPaletteMetadataBytes ||
                checked(encodedBytes + updatedMetadataBytes + 64L * 1024L) > OfficeRasterGuards.MaximumDecodedBytes) {
                return false;
            }
            metadataBytes = updatedMetadataBytes;
            return true;
        } catch (OverflowException) {
            return false;
        }
    }

    private static bool HasValidModificationTime(byte[] bytes, int offset, int length) {
        if (length != 7) return false;
        int year = bytes[offset] << 8 | bytes[offset + 1];
        int month = bytes[offset + 2];
        int day = bytes[offset + 3];
        int hour = bytes[offset + 4];
        int minute = bytes[offset + 5];
        int second = bytes[offset + 6];
        if (year == 0 || month < 1 || month > 12 || hour > 23 || minute > 59 || second > 60) {
            return false;
        }

        int daysInMonth;
        switch (month) {
            case 2:
                bool leapYear = year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);
                daysInMonth = leapYear ? 29 : 28;
                break;
            case 4:
            case 6:
            case 9:
            case 11:
                daysInMonth = 30;
                break;
            default:
                daysInMonth = 31;
                break;
        }
        return day >= 1 && day <= daysInMonth;
    }

    private static bool HasValidChromaticities(byte[] bytes, int offset) {
        var coordinates = new uint[8];
        for (int index = 0; index < coordinates.Length; index++) {
            uint value = ReadBigEndianUInt32(bytes, offset + index * 4);
            if (value > int.MaxValue) return false;
            coordinates[index] = value;
        }

        if (coordinates[1] == 0) return false;
        for (int index = 0; index < coordinates.Length; index += 2) {
            uint x = coordinates[index];
            uint y = coordinates[index + 1];
            if (x > 100000U || y > 100000U - x) return false;
        }
        return true;
    }

    private static bool HasStandardRgbChromaticities(byte[] bytes, int offset) {
        int[] expected = { 31270, 32900, 64000, 33000, 30000, 60000, 15000, 6000 };
        for (int index = 0; index < expected.Length; index++) {
            if (ReadBigEndianUInt32(bytes, offset + index * 4) != expected[index]) return false;
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

    private static bool HasValidIccProfile(
        byte[] bytes,
        int offset,
        int length,
        CancellationToken cancellationToken) {
        if (length < 9 || !TryReadKeyword(bytes, offset, length, out int keywordEnd)) return false;
        int compressionMethodOffset = keywordEnd + 1;
        if (compressionMethodOffset >= offset + length || bytes[compressionMethodOffset] != 0) return false;
        int compressedOffset = compressionMethodOffset + 1;
        int compressedLength = offset + length - compressedOffset;
        if (compressedLength < 6 || !TryGetCompressedMetadataOutputLimit(
                bytes.LongLength,
                compressedLength,
                OfficeRasterGuards.MaximumEncodedBytes,
                out int maximumProfileBytes)) return false;
        var compressed = new byte[compressedLength];
        Buffer.BlockCopy(bytes, compressedOffset, compressed, 0, compressedLength);
        try {
            byte[] profile = OfficeZlibCodec.Decompress(
                compressed,
                maximumProfileBytes,
                cancellationToken: cancellationToken);
            return OfficeIccProfileValidator.TryValidate(profile, 0, profile.Length);
        } catch (Exception exception) when (
            exception is ArgumentException ||
            exception is FormatException ||
            exception is OfficeDecompressionSizeLimitException ||
            exception is InvalidDataException ||
            exception is NotSupportedException ||
            exception is OverflowException) {
            return false;
        }
    }

    private static bool HasValidLatinText(
        byte[] bytes,
        int offset,
        int length,
        ref long decodedTextBytes,
        CancellationToken cancellationToken) {
        if (!TryReadKeyword(bytes, offset, length, out int keywordEnd)) return false;
        int textOffset = keywordEnd + 1;
        int textLength = offset + length - textOffset;
        for (int index = 0; index < textLength; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (bytes[textOffset + index] == 0) return false;
        }
        return TryAddDecodedTextBytes(ref decodedTextBytes, textLength);
    }

    private static bool HasValidCompressedText(
        byte[] bytes,
        int offset,
        int length,
        ref long decodedTextBytes,
        CancellationToken cancellationToken) {
        if (!TryReadKeyword(bytes, offset, length, out int keywordEnd)) return false;
        int methodOffset = keywordEnd + 1;
        if (methodOffset >= offset + length || bytes[methodOffset] != 0) return false;
        return TryInflateText(
            bytes,
            methodOffset + 1,
            offset + length - methodOffset - 1,
            ref decodedTextBytes,
            requireUtf8: false,
            cancellationToken);
    }

    private static bool HasValidInternationalText(
        byte[] bytes,
        int offset,
        int length,
        ref long decodedTextBytes,
        CancellationToken cancellationToken) {
        if (!TryReadKeyword(bytes, offset, length, out int keywordEnd)) return false;
        int end = offset + length;
        int flagOffset = keywordEnd + 1;
        if (flagOffset > end - 2 || bytes[flagOffset] > 1 || bytes[flagOffset + 1] != 0) return false;

        int languageOffset = flagOffset + 2;
        int languageEnd = FindNull(bytes, languageOffset, end);
        if (languageEnd < 0 || !HasValidLanguageTag(bytes, languageOffset, languageEnd - languageOffset)) return false;
        int translatedOffset = languageEnd + 1;
        int translatedEnd = FindNull(bytes, translatedOffset, end);
        int translatedLength = translatedEnd - translatedOffset;
        if (translatedEnd < 0 ||
            !TryAddDecodedTextBytes(ref decodedTextBytes, translatedLength) ||
            !HasValidUtf8(bytes, translatedOffset, translatedLength)) return false;
        int textOffset = translatedEnd + 1;
        int textLength = end - textOffset;
        if (bytes[flagOffset] == 1) {
            return TryInflateText(
                bytes, textOffset, textLength, ref decodedTextBytes, requireUtf8: true, cancellationToken);
        }
        return HasValidUtf8(bytes, textOffset, textLength) &&
               TryAddDecodedTextBytes(ref decodedTextBytes, textLength);
    }

    private static bool TryInflateText(
        byte[] bytes,
        int offset,
        int length,
        ref long decodedTextBytes,
        bool requireUtf8,
        CancellationToken cancellationToken) {
        if (length < 6 || !TryGetCompressedMetadataOutputLimit(
                bytes.LongLength,
                length,
                MaximumPngTextBytes,
                out int maximumTextBytes)) return false;
        var compressed = new byte[length];
        Buffer.BlockCopy(bytes, offset, compressed, 0, length);
        try {
            byte[] text = OfficeZlibCodec.Decompress(
                compressed, maximumTextBytes, cancellationToken: cancellationToken);
            return (!requireUtf8 || HasValidUtf8(text, 0, text.Length)) &&
                   TryAddDecodedTextBytes(ref decodedTextBytes, text.Length);
        } catch (Exception exception) when (
            exception is ArgumentException ||
            exception is FormatException ||
            exception is OfficeDecompressionSizeLimitException ||
            exception is InvalidDataException ||
            exception is NotSupportedException ||
            exception is OverflowException) {
            return false;
        }
    }

    internal static bool TryGetCompressedMetadataOutputLimit(
        long encodedBytes,
        long compressedBytes,
        int requestedMaximumOutputBytes,
        out int maximumOutputBytes) {
        maximumOutputBytes = 0;
        if (encodedBytes < 0L || compressedBytes < 0L || requestedMaximumOutputBytes < 1) return false;
        try {
            long availableBytes = checked(
                OfficeRasterGuards.MaximumDecodedBytes - encodedBytes - compressedBytes - 64L * 1024L);
            // Bounded zlib output can retain MemoryStream growth capacity and its final ToArray copy together.
            long boundedOutputBytes = Math.Min(requestedMaximumOutputBytes, availableBytes / 3L);
            if (boundedOutputBytes < 1L) return false;
            maximumOutputBytes = (int)boundedOutputBytes;
            return true;
        } catch (OverflowException) {
            return false;
        }
    }

    private static bool TryReadKeyword(byte[] bytes, int offset, int length, out int keywordEnd) {
        keywordEnd = offset;
        int end = offset + length;
        while (keywordEnd < end && keywordEnd - offset < 80 && bytes[keywordEnd] != 0) {
            byte value = bytes[keywordEnd];
            if (!IsValidKeywordByte(value) ||
                (value == (byte)' ' && (keywordEnd == offset || bytes[keywordEnd - 1] == (byte)' '))) {
                return false;
            }
            keywordEnd++;
        }
        int keywordLength = keywordEnd - offset;
        return keywordLength >= 1 && keywordLength <= 79 && keywordEnd < end &&
               bytes[keywordEnd] == 0 && bytes[keywordEnd - 1] != (byte)' ';
    }

    private static int FindNull(byte[] bytes, int offset, int end) {
        for (int index = offset; index < end; index++) {
            if (bytes[index] == 0) return index;
        }
        return -1;
    }

    private static bool HasValidLanguageTag(byte[] bytes, int offset, int length) {
        for (int index = 0; index < length; index++) {
            byte value = bytes[offset + index];
            if (!((value >= (byte)'A' && value <= (byte)'Z') ||
                  (value >= (byte)'a' && value <= (byte)'z') ||
                  (value >= (byte)'0' && value <= (byte)'9') ||
                  value == (byte)'-')) return false;
        }
        return true;
    }

    private static bool HasValidUtf8(byte[] bytes, int offset, int length) {
        try {
            _ = StrictUtf8.GetCharCount(bytes, offset, length);
            return true;
        } catch (DecoderFallbackException) {
            return false;
        }
    }

    private static bool TryAddDecodedTextBytes(ref long total, int count) {
        total += count;
        return total <= MaximumPngTextBytes;
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

    private static uint ComputeCrc(
        byte[] bytes,
        int offset,
        int count,
        CancellationToken cancellationToken) {
        uint crc = 0xFFFFFFFFU;
        for (int index = 0; index < count; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
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
