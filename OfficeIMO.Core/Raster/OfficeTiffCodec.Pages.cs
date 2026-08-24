using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    internal static bool TryInspectPages(
        byte[] encodedBytes,
        OfficeRasterDecodeOptions options,
        out OfficeRasterContainerInfo? container) =>
        TryInspectPages(encodedBytes, options, validatePayloads: false,
            OfficeRasterGuards.MaximumDecodedBytes, out container);

    internal static bool TryValidateAllPages(byte[] encodedBytes) {
        var options = new OfficeRasterDecodeOptions();
        return TryInspectPages(encodedBytes, options, validatePayloads: true,
            OfficeRasterGuards.MaximumDecodedBytes, out _);
    }

    internal static bool TryValidateAllPages(
        byte[] encodedBytes,
        OfficeRasterDecodeOptions options) =>
        TryInspectPages(encodedBytes, options, validatePayloads: true,
            OfficeRasterGuards.MaximumDecodedBytes, out _);

    internal static bool TryValidateAllPages(
        byte[] encodedBytes,
        OfficeRasterDecodeOptions options,
        long maximumValidationWorkBytes) =>
        TryInspectPages(encodedBytes, options, validatePayloads: true,
            maximumValidationWorkBytes, out _);

    private static bool TryInspectPages(
        byte[] encodedBytes,
        OfficeRasterDecodeOptions options,
        bool validatePayloads,
        long maximumValidationWorkBytes,
        out OfficeRasterContainerInfo? container) {
        container = null;
        if (!IsTiff(encodedBytes) || encodedBytes.Length > options.MaximumEncodedBytes ||
            !OfficeTiffStructureValidator.TryValidate(
                encodedBytes, 0, encodedBytes.Length, options.CancellationToken)) {
            return false;
        }

        try {
            bool littleEndian = encodedBytes[0] == (byte)'I';
            int ifdOffset = ReadOffset(encodedBytes, 4, littleEndian);
            var visitedIfds = new HashSet<int>();
            var frames = new List<OfficeRasterFrameInfo>();
            long validatedPixels = 0L;
            var validationBudget = validatePayloads
                ? new TiffValidationBudget(maximumValidationWorkBytes)
                : null;
            while (ifdOffset != 0) {
                options.CancellationToken.ThrowIfCancellationRequested();
                if (frames.Count >= MaximumIfdCount || !visitedIfds.Add(ifdOffset) ||
                    !HasBytes(encodedBytes, ifdOffset, 2)) return false;
                int entryCount = ReadUInt16(encodedBytes, ifdOffset, littleEndian);
                if (entryCount <= 0 || !HasBytes(encodedBytes, ifdOffset + 2, checked(entryCount * 12 + 4))) return false;
                var entries = ReadEntries(
                    encodedBytes, ifdOffset, entryCount, littleEndian, options.CancellationToken);
                if (entries == null ||
                    !TryReadScalar(encodedBytes, entries, 256, littleEndian, out int width) ||
                    !TryReadScalar(encodedBytes, entries, 257, littleEndian, out int height) ||
                    width < 1 || height < 1 ||
                    validatePayloads && !TryReserveAndValidatePage(
                        encodedBytes, entries, littleEndian, width, height, options,
                        validationBudget!, ref validatedPixels)) return false;

                frames.Add(new OfficeRasterFrameInfo(
                    frames.Count,
                    OfficeRasterFrameKind.Page,
                    width,
                    height,
                    0,
                    0,
                    TimeSpan.Zero,
                    OfficeRasterFrameDisposal.None,
                    OfficeRasterFrameBlend.Source,
                    frames.Count == 0));
                int nextIfdPointerOffset = checked(ifdOffset + 2 + entryCount * 12);
                ifdOffset = ReadOffset(encodedBytes, nextIfdPointerOffset, littleEndian);
            }
            if (frames.Count == 0) return false;
            OfficeRasterFrameInfo first = frames[0];
            container = new OfficeRasterContainerInfo(
                OfficeImageFormat.Tiff,
                first.Width,
                first.Height,
                frames.ToArray(),
                1,
                OfficeColor.Transparent);
            return true;
        } catch (ArgumentException) {
            return false;
        } catch (FormatException) {
            return false;
        } catch (OverflowException) {
            return false;
        }
    }

    private static Dictionary<int, TiffEntry>? ReadEntries(
        byte[] encodedBytes,
        int ifdOffset,
        int entryCount,
        bool littleEndian,
        System.Threading.CancellationToken cancellationToken) {
        var entries = new Dictionary<int, TiffEntry>();
        int entryOffset = ifdOffset + 2;
        for (int index = 0; index < entryCount; index++, entryOffset += 12) {
            if ((index & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
            int tag = ReadUInt16(encodedBytes, entryOffset, littleEndian);
            int type = ReadUInt16(encodedBytes, entryOffset + 2, littleEndian);
            uint count = ReadUInt32(encodedBytes, entryOffset + 4, littleEndian);
            if (count == 0 || count > int.MaxValue || entries.ContainsKey(tag) ||
                !HasValidEntryValueRange(encodedBytes, type, (int)count, entryOffset + 8, littleEndian)) {
                return null;
            }
            entries.Add(tag, new TiffEntry(type, (int)count, entryOffset + 8));
        }
        return entries;
    }

    private static bool TryReserveAndValidatePage(
        byte[] encodedBytes,
        IReadOnlyDictionary<int, TiffEntry> entries,
        bool littleEndian,
        int width,
        int height,
        OfficeRasterDecodeOptions options,
        TiffValidationBudget validationBudget,
        ref long validatedPixels) {
        long pagePixels = (long)width * height;
        if (pagePixels <= 0L || pagePixels > options.MaximumDecodedPixels - validatedPixels) return false;
        validatedPixels += pagePixels;
        return TryValidateStripPage(encodedBytes, entries, littleEndian, width, height,
            options, validationBudget);
    }

    private static bool TryValidateStripPage(
        byte[] encodedBytes,
        IReadOnlyDictionary<int, TiffEntry> entries,
        bool littleEndian,
        int width,
        int height,
        OfficeRasterDecodeOptions options,
        TiffValidationBudget validationBudget) {
        if (!TryReadScalarOrDefault(encodedBytes, entries, 259, littleEndian, 1, out int compression) ||
            !TryReadScalarOrDefault(encodedBytes, entries, 262, littleEndian, 2, out int photometric) ||
            !TryReadScalarOrDefault(encodedBytes, entries, 274, littleEndian, 1, out int orientation) ||
            !TryReadScalarOrDefault(encodedBytes, entries, 284, littleEndian, 1, out int planarConfiguration) ||
            !TryReadScalarOrDefault(encodedBytes, entries, 317, littleEndian, 1, out int predictor) ||
            !TryGetBaseSampleCount(photometric, out int baseSamples) ||
            !TryReadScalarOrDefault(encodedBytes, entries, 277, littleEndian, baseSamples, out int samples) ||
            (planarConfiguration != 1 && planarConfiguration != 2) ||
            (predictor != 1 && predictor != 2) ||
            (samples != baseSamples && samples != baseSamples + 1) ||
            orientation < 1 || orientation > 8 ||
            (compression != (int)OfficeTiffCompression.None &&
             compression != (int)OfficeTiffCompression.Lzw &&
             compression != (int)OfficeTiffCompression.PackBits &&
             compression != (int)OfficeTiffCompression.Deflate &&
             compression != 32946)) return false;

        if (!TryReadValues(encodedBytes, entries, 258, littleEndian, samples, out int[] bitsPerSample) ||
            Array.Exists(bitsPerSample, value => value != 8)) return false;

        if (photometric == 5 &&
            (!TryReadScalarOrDefault(encodedBytes, entries, 332, littleEndian, 1, out int inkSet) || inkSet != 1)) {
            return false;
        }
        if (photometric == 3 && !TryReadValues(encodedBytes, entries, 320, littleEndian, 768, out _)) {
            return false;
        }
        if (samples == baseSamples + 1 &&
            (!TryReadValues(encodedBytes, entries, 338, littleEndian, 1, out int[] extraSamples) ||
             (extraSamples[0] != 1 && extraSamples[0] != 2))) {
            return false;
        }

        return TryDecodePixelSegments(encodedBytes, entries, littleEndian, width, height, samples,
            compression, planarConfiguration, predictor, options, validationBudget,
            retainPixels: false, out _);
    }

    private sealed class TiffValidationBudget {
        private readonly long _maximumWorkBytes;
        private long _reservedWorkBytes;

        internal TiffValidationBudget(long maximumWorkBytes) {
            if (maximumWorkBytes < 1L || maximumWorkBytes > OfficeRasterGuards.MaximumDecodedBytes) {
                throw new ArgumentOutOfRangeException(nameof(maximumWorkBytes));
            }
            _maximumWorkBytes = maximumWorkBytes;
        }

        internal bool TryReserve(int compressedBytes, int decodedBytes) {
            if (compressedBytes < 0 || decodedBytes < 0) return false;
            long workBytes = (long)compressedBytes + decodedBytes;
            if (workBytes > _maximumWorkBytes - _reservedWorkBytes) return false;
            _reservedWorkBytes += workBytes;
            return true;
        }
    }

}
