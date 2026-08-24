using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    internal static bool TryInspectPages(
        byte[] encodedBytes,
        OfficeRasterDecodeOptions options,
        out OfficeRasterContainerInfo? container) =>
        TryInspectPages(encodedBytes, options, enforceAllPagePixelLimits: true, out container);

    internal static bool TryInspectPages(
        byte[] encodedBytes,
        OfficeRasterDecodeOptions options,
        bool enforceAllPagePixelLimits,
        out OfficeRasterContainerInfo? container) =>
        TryInspectPages(encodedBytes, options, validatePayloads: enforceAllPagePixelLimits, enforceAllPagePixelLimits,
            OfficeRasterGuards.MaximumDecodedBytes - options.RetainedManagedBytes, out container);

    internal static bool TryValidateAllPages(byte[] encodedBytes) {
        var options = new OfficeRasterDecodeOptions();
        return TryInspectPages(encodedBytes, options, validatePayloads: true, enforceAllPagePixelLimits: true,
            OfficeRasterGuards.MaximumDecodedBytes, out _);
    }

    internal static bool TryValidateAllPages(
        byte[] encodedBytes,
        OfficeRasterDecodeOptions options) =>
        TryInspectPages(encodedBytes, options, validatePayloads: true, enforceAllPagePixelLimits: true,
            OfficeRasterGuards.MaximumDecodedBytes - options.RetainedManagedBytes, out _);

    internal static bool TryValidateAllPages(
        byte[] encodedBytes,
        OfficeRasterDecodeOptions options,
        long maximumValidationWorkBytes) =>
        TryInspectPages(encodedBytes, options, validatePayloads: true, enforceAllPagePixelLimits: true,
            maximumValidationWorkBytes, out _);

    private static bool TryInspectPages(
        byte[] encodedBytes,
        OfficeRasterDecodeOptions options,
        bool validatePayloads,
        bool enforceAllPagePixelLimits,
        long maximumValidationWorkBytes,
        out OfficeRasterContainerInfo? container) {
        container = null;
        if (maximumValidationWorkBytes < 1L ||
            !IsTiff(encodedBytes) || encodedBytes.Length > options.MaximumEncodedBytes ||
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
            int firstOrientation = 1;
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
                    enforceAllPagePixelLimits &&
                    (long)width * height > options.MaximumDecodedPixels ||
                    !TryReadPageDpi(encodedBytes, entries, littleEndian, out double? dpiX, out double? dpiY) ||
                    validatePayloads && !TryReserveAndValidatePage(
                        encodedBytes, entries, littleEndian, width, height, options,
                        validationBudget!, ref validatedPixels)) return false;

                bool hasValidOrientation =
                    TryReadScalarOrDefault(encodedBytes, entries, 274, littleEndian, 1, out int orientation) &&
                    orientation >= 1 && orientation <= 8;
                // Full inventory may not advertise malformed page orientation. Selected-page decoding
                // still skips unsupported metadata on pages that the caller did not select.
                if (!hasValidOrientation &&
                    (enforceAllPagePixelLimits || frames.Count == options.FrameIndex)) return false;
                if (!hasValidOrientation) orientation = 1;

                if (frames.Count == 0) firstOrientation = orientation;
                if (orientation >= 5) (dpiX, dpiY) = (dpiY, dpiX);

                if (encodedBytes.LongLength + options.RetainedManagedBytes +
                    checked((frames.Count + 1L) * 128L) + 64L * 1024L >
                    OfficeRasterGuards.MaximumDecodedBytes) return false;
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
                    frames.Count == 0,
                    dpiX,
                    dpiY));
                int nextIfdPointerOffset = checked(ifdOffset + 2 + entryCount * 12);
                ifdOffset = ReadOffset(encodedBytes, nextIfdPointerOffset, littleEndian);
            }
            if (frames.Count == 0) return false;
            OfficeRasterFrameInfo first = frames[0];
            bool firstPageSwapsAxes = firstOrientation >= 5;
            container = new OfficeRasterContainerInfo(
                OfficeImageFormat.Tiff,
                firstPageSwapsAxes ? first.Height : first.Width,
                firstPageSwapsAxes ? first.Width : first.Height,
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

    private static bool TryReadPageDpi(
        byte[] encodedBytes,
        IReadOnlyDictionary<int, TiffEntry> entries,
        bool littleEndian,
        out double? dpiX,
        out double? dpiY) {
        dpiX = null;
        dpiY = null;
        bool hasX = entries.ContainsKey(282);
        bool hasY = entries.ContainsKey(283);
        if (!hasX && !hasY) return true;
        if (!TryReadScalarOrDefault(encodedBytes, entries, 296, littleEndian, 2, out int unit) ||
            unit < 1 || unit > 3) return false;
        if (unit == 1 || !hasX || !hasY) return true;
        if (!TryReadPositiveRational(encodedBytes, entries, 282, littleEndian, out double x) ||
            !TryReadPositiveRational(encodedBytes, entries, 283, littleEndian, out double y)) return false;
        double scale = unit == 3 ? 2.54D : 1D;
        double physicalDpiX = x * scale;
        double physicalDpiY = y * scale;
        if (double.IsNaN(physicalDpiX) || double.IsInfinity(physicalDpiX) ||
            double.IsNaN(physicalDpiY) || double.IsInfinity(physicalDpiY)) return false;
        dpiX = physicalDpiX;
        dpiY = physicalDpiY;
        return true;
    }

    private static bool TryReadPositiveRational(
        byte[] data,
        IReadOnlyDictionary<int, TiffEntry> entries,
        int tag,
        bool littleEndian,
        out double value) {
        value = 0D;
        if (!entries.TryGetValue(tag, out TiffEntry entry) || entry.Type != 5 || entry.Count != 1) {
            return false;
        }
        int offset = ReadOffset(data, entry.ValueFieldOffset, littleEndian);
        if (!HasBytes(data, offset, 8)) return false;
        uint numerator = ReadUInt32(data, offset, littleEndian);
        uint denominator = ReadUInt32(data, offset + 4, littleEndian);
        if (numerator == 0 || denominator == 0) return false;
        value = numerator / (double)denominator;
        return value > 0D;
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
