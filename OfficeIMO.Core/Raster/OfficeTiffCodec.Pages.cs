using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    internal static bool TryInspectPages(
        byte[] encodedBytes,
        OfficeRasterDecodeOptions options,
        out OfficeRasterContainerInfo? container) {
        container = null;
        if (!IsTiff(encodedBytes) || encodedBytes.Length > options.MaximumEncodedBytes ||
            !OfficeTiffStructureValidator.TryValidate(encodedBytes, 0, encodedBytes.Length)) {
            return false;
        }

        try {
            bool littleEndian = encodedBytes[0] == (byte)'I';
            int ifdOffset = ReadOffset(encodedBytes, 4, littleEndian);
            var visitedIfds = new HashSet<int>();
            var frames = new List<OfficeRasterFrameInfo>();
            while (ifdOffset != 0) {
                options.CancellationToken.ThrowIfCancellationRequested();
                if (frames.Count >= MaximumIfdCount || !visitedIfds.Add(ifdOffset) ||
                    !HasBytes(encodedBytes, ifdOffset, 2)) return false;
                int entryCount = ReadUInt16(encodedBytes, ifdOffset, littleEndian);
                if (entryCount <= 0 || !HasBytes(encodedBytes, ifdOffset + 2, checked(entryCount * 12 + 4))) return false;
                var entries = ReadEntries(encodedBytes, ifdOffset, entryCount, littleEndian);
                if (entries == null ||
                    !TryReadScalar(encodedBytes, entries, 256, littleEndian, out int width) ||
                    !TryReadScalar(encodedBytes, entries, 257, littleEndian, out int height) ||
                    !OfficeRasterImageDecoder.IsWithinPixelLimit(width, height, options.MaximumDecodedPixels) ||
                    !TryReadScalarOrDefault(encodedBytes, entries, 274, littleEndian, 1, out int orientation) ||
                    orientation < 1 || orientation > 8 ||
                    !TryValidateStripPage(encodedBytes, entries, littleEndian, width, height, options)) return false;

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
        bool littleEndian) {
        var entries = new Dictionary<int, TiffEntry>();
        int entryOffset = ifdOffset + 2;
        for (int index = 0; index < entryCount; index++, entryOffset += 12) {
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

    private static bool TryValidateStripPage(
        byte[] encodedBytes,
        IReadOnlyDictionary<int, TiffEntry> entries,
        bool littleEndian,
        int width,
        int height,
        OfficeRasterDecodeOptions options) {
        if (!TryReadScalarOrDefault(encodedBytes, entries, 259, littleEndian, 1, out int compression) ||
            !TryReadScalarOrDefault(encodedBytes, entries, 262, littleEndian, 2, out int photometric) ||
            !TryReadScalarOrDefault(encodedBytes, entries, 284, littleEndian, 1, out int planarConfiguration) ||
            !TryReadScalarOrDefault(encodedBytes, entries, 317, littleEndian, 1, out int predictor) ||
            !TryGetBaseSampleCount(photometric, out int baseSamples) ||
            !TryReadScalarOrDefault(encodedBytes, entries, 277, littleEndian, baseSamples, out int samples) ||
            (planarConfiguration != 1 && planarConfiguration != 2) ||
            (predictor != 1 && predictor != 2) ||
            (samples != baseSamples && samples != baseSamples + 1) ||
            (compression != (int)OfficeTiffCompression.None &&
             compression != (int)OfficeTiffCompression.Lzw &&
             compression != (int)OfficeTiffCompression.PackBits &&
             compression != (int)OfficeTiffCompression.Deflate &&
             compression != 32946)) return false;

        if (!TryReadValues(encodedBytes, entries, 258, littleEndian, samples, out int[] bitsPerSample) ||
            Array.Exists(bitsPerSample, value => value != 8)) return false;

        return TryDecodePixelSegments(encodedBytes, entries, littleEndian, width, height, samples,
            compression, planarConfiguration, predictor, options, retainPixels: false, out _);
    }
}
