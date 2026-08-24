using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    private static bool TryDecodePixelSegments(
        byte[] encodedBytes,
        IReadOnlyDictionary<int, TiffEntry> entries,
        bool littleEndian,
        int width,
        int height,
        int samples,
        int compression,
        int planarConfiguration,
        int predictor,
        OfficeRasterDecodeOptions options,
        TiffValidationBudget? validationBudget,
        bool retainPixels,
        out byte[] source) {
        source = Array.Empty<byte>();
        if (planarConfiguration != 1 && planarConfiguration != 2 ||
            predictor != 1 && predictor != 2 || samples < 1) return false;

        bool hasStrips = entries.ContainsKey(273) || entries.ContainsKey(279);
        bool hasTiles = entries.ContainsKey(324) || entries.ContainsKey(325) ||
                        entries.ContainsKey(322) || entries.ContainsKey(323);
        if (hasStrips == hasTiles) return false;

        int sourceLength = OfficeRasterGuards.EnsureByteCount(
            (long)width * height * samples,
            "TIFF decoded source pixels exceed the managed limit.");

        if (hasStrips) {
            if (!TryReadScalarOrDefault(encodedBytes, entries, 278, littleEndian, height, out int rowsPerStrip) ||
                rowsPerStrip < 1) return false;
            int segmentsPerPlane = checked((height + rowsPerStrip - 1) / rowsPerStrip);
            int segmentCount = checked(segmentsPerPlane * (planarConfiguration == 2 ? samples : 1));
            int segmentSamples = planarConfiguration == 2 ? 1 : samples;
            int maximumRows = Math.Min(rowsPerStrip, height);
            int maximumDecodedSegment = checked(maximumRows * width * segmentSamples);
            int scratchLength = retainPixels && planarConfiguration == 1 ? 0 : maximumDecodedSegment;
            long segmentMetadataBytes = checked((long)segmentCount * 2L * sizeof(int));
            int finalRgbaLength = retainPixels
                ? OfficeRasterGuards.EnsureByteCount(
                    (long)width * height * 4L,
                    "TIFF RGBA output exceeds the managed limit.")
                : 0;
            if (!IsTiffDecodeWorkingSetWithinLimit(
                    encodedBytes.LongLength,
                    sourceLength,
                    scratchLength,
                    finalRgbaLength,
                    segmentMetadataBytes,
                    retainPixels,
                    compression,
                    maximumCompressedSegmentLength: 0,
                    maximumDecodedSegment,
                    options.RetainedManagedBytes)) return false;
            if (!TryReadValues(encodedBytes, entries, 273, littleEndian, segmentCount,
                    options.CancellationToken, out int[] offsets) ||
                !TryReadValues(encodedBytes, entries, 279, littleEndian, segmentCount,
                    options.CancellationToken, out int[] byteCounts)) return false;

            int maximumCompressedSegment = 0;
            for (int segment = 0; segment < segmentCount; segment++) {
                options.CancellationToken.ThrowIfCancellationRequested();
                if (!HasSegment(encodedBytes, offsets[segment], byteCounts[segment])) return false;
                int strip = planarConfiguration == 2 ? segment % segmentsPerPlane : segment;
                int rowStart = checked(strip * rowsPerStrip);
                if (rowStart >= height) return false;
                int rows = Math.Min(rowsPerStrip, height - rowStart);
                int expected = checked(rows * width * segmentSamples);
                if (validationBudget != null &&
                    !validationBudget.TryReserve(byteCounts[segment], expected)) return false;
                maximumCompressedSegment = Math.Max(maximumCompressedSegment, byteCounts[segment]);
            }
            if (!IsTiffDecodeWorkingSetWithinLimit(
                    encodedBytes.LongLength,
                    sourceLength,
                    scratchLength,
                    finalRgbaLength,
                    segmentMetadataBytes,
                    retainPixels,
                    compression,
                    maximumCompressedSegment,
                    maximumDecodedSegment,
                    options.RetainedManagedBytes)) return false;
            if (retainPixels) source = new byte[sourceLength];
            byte[]? scratch = scratchLength > 0 ? new byte[scratchLength] : null;

            for (int segment = 0; segment < segmentCount; segment++) {
                options.CancellationToken.ThrowIfCancellationRequested();
                int plane = planarConfiguration == 2 ? segment / segmentsPerPlane : 0;
                int strip = planarConfiguration == 2 ? segment % segmentsPerPlane : segment;
                int rowStart = checked(strip * rowsPerStrip);
                int rows = Math.Min(rowsPerStrip, height - rowStart);
                int expected = checked(rows * width * segmentSamples);
                byte[] decoded = retainPixels && planarConfiguration == 1
                    ? source
                    : scratch!;
                int decodedOffset = retainPixels && planarConfiguration == 1
                    ? checked(rowStart * width * samples)
                    : 0;
                if (!TryDecodeStrip(encodedBytes, offsets[segment], byteCounts[segment], compression,
                        decoded, decodedOffset, expected, options.CancellationToken)) return false;
                if (predictor == 2) ReverseHorizontalPredictor(decoded, decodedOffset, rows, width,
                    segmentSamples, options.CancellationToken);
                if (retainPixels && planarConfiguration == 2) {
                    CopyPlanarRows(decoded, source, plane, samples, width, rowStart, rows, options);
                }
            }
            return true;
        }

        if (!TryReadScalar(encodedBytes, entries, 322, littleEndian, out int tileWidth) ||
            !TryReadScalar(encodedBytes, entries, 323, littleEndian, out int tileHeight) ||
            tileWidth < 1 || tileHeight < 1) return false;
        int tilesAcross = checked((width + tileWidth - 1) / tileWidth);
        int tilesDown = checked((height + tileHeight - 1) / tileHeight);
        int segmentsPerTilePlane = checked(tilesAcross * tilesDown);
        int tileSegmentCount = checked(segmentsPerTilePlane * (planarConfiguration == 2 ? samples : 1));
        int tileSamples = planarConfiguration == 2 ? 1 : samples;
        int tileByteLength = OfficeRasterGuards.EnsureByteCount(
            (long)tileWidth * tileHeight * tileSamples,
            "TIFF decoded tile exceeds the managed limit.");
        long tileMetadataBytes = checked((long)tileSegmentCount * 2L * sizeof(int));
        int tileFinalRgbaLength = retainPixels
            ? OfficeRasterGuards.EnsureByteCount(
                (long)width * height * 4L,
                "TIFF RGBA output exceeds the managed limit.")
            : 0;
        if (!IsTiffDecodeWorkingSetWithinLimit(
                encodedBytes.LongLength,
                sourceLength,
                tileByteLength,
                tileFinalRgbaLength,
                tileMetadataBytes,
                retainPixels,
                compression,
                maximumCompressedSegmentLength: 0,
                tileByteLength,
                options.RetainedManagedBytes)) return false;
        if (!TryReadValues(encodedBytes, entries, 324, littleEndian, tileSegmentCount,
                options.CancellationToken, out int[] tileOffsets) ||
            !TryReadValues(encodedBytes, entries, 325, littleEndian, tileSegmentCount,
                options.CancellationToken, out int[] tileByteCounts)) return false;

        int maximumCompressedTile = 0;
        for (int segment = 0; segment < tileSegmentCount; segment++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (!HasSegment(encodedBytes, tileOffsets[segment], tileByteCounts[segment])) return false;
            if (validationBudget != null &&
                !validationBudget.TryReserve(tileByteCounts[segment], tileByteLength)) return false;
            maximumCompressedTile = Math.Max(maximumCompressedTile, tileByteCounts[segment]);
        }
        if (!IsTiffDecodeWorkingSetWithinLimit(
                encodedBytes.LongLength,
                sourceLength,
                tileByteLength,
                tileFinalRgbaLength,
                tileMetadataBytes,
                retainPixels,
                compression,
                maximumCompressedTile,
                tileByteLength,
                options.RetainedManagedBytes)) return false;
        if (retainPixels) source = new byte[sourceLength];
        var tileDecoded = new byte[tileByteLength];
        for (int segment = 0; segment < tileSegmentCount; segment++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            int plane = planarConfiguration == 2 ? segment / segmentsPerTilePlane : 0;
            int tile = planarConfiguration == 2 ? segment % segmentsPerTilePlane : segment;
            int tileX = checked((tile % tilesAcross) * tileWidth);
            int tileY = checked((tile / tilesAcross) * tileHeight);
            if (!TryDecodeStrip(encodedBytes, tileOffsets[segment], tileByteCounts[segment], compression,
                    tileDecoded, 0, tileByteLength, options.CancellationToken)) return false;
            if (predictor == 2) ReverseHorizontalPredictor(tileDecoded, 0, tileHeight, tileWidth,
                tileSamples, options.CancellationToken);
            if (retainPixels) {
                CopyTile(tileDecoded, source, plane, planarConfiguration, samples, width, height,
                    tileX, tileY, tileWidth, tileHeight, options);
            }
        }
        return true;
    }

    internal static bool IsTiffDecodeWorkingSetWithinLimit(
        long encodedLength,
        int sourceLength,
        int scratchLength,
        int finalRgbaLength,
        long segmentMetadataBytes,
        bool retainPixels,
        int compression,
        int maximumCompressedSegmentLength,
        int maximumDecodedSegmentLength,
        long retainedManagedBytes = 0L) {
        if (encodedLength < 0L || sourceLength < 0 || scratchLength < 0 || finalRgbaLength < 0 ||
            segmentMetadataBytes < 0L || maximumCompressedSegmentLength < 0 ||
            maximumDecodedSegmentLength < 0 || retainedManagedBytes < 0L) return false;
        try {
            long retainedSourceBytes = retainPixels ? sourceLength : 0L;
            bool usesDeflateTemporaries = compression == (int)OfficeTiffCompression.Deflate || compression == 32946;
            long codecTemporaryBytes = usesDeflateTemporaries
                ? checked((long)maximumCompressedSegmentLength + maximumDecodedSegmentLength)
                : 0L;
            long segmentPeak = checked(
                encodedLength + retainedSourceBytes + scratchLength + segmentMetadataBytes +
                codecTemporaryBytes + retainedManagedBytes + 64L * 1024L);
            long conversionPeak = retainPixels
                ? checked(encodedLength + sourceLength + finalRgbaLength + retainedManagedBytes + 64L * 1024L)
                : 0L;
            return Math.Max(segmentPeak, conversionPeak) <= OfficeRasterGuards.MaximumDecodedBytes;
        } catch (OverflowException) {
            return false;
        }
    }

    private static bool HasSegment(byte[] encodedBytes, int offset, int count) =>
        offset >= 0 && count > 0 && offset <= encodedBytes.Length - count;

    private static void CopyPlanarRows(
        byte[] planeBytes,
        byte[] interleaved,
        int plane,
        int samples,
        int width,
        int rowStart,
        int rows,
        OfficeRasterDecodeOptions options) {
        for (int row = 0; row < rows; row++) {
            if ((row & 31) == 0) options.CancellationToken.ThrowIfCancellationRequested();
            int sourceRow = row * width;
            int targetRow = checked((rowStart + row) * width * samples + plane);
            for (int x = 0; x < width; x++) {
                if ((x & 0xFFF) == 0) options.CancellationToken.ThrowIfCancellationRequested();
                interleaved[targetRow + x * samples] = planeBytes[sourceRow + x];
            }
        }
    }

    private static void CopyTile(
        byte[] tile,
        byte[] interleaved,
        int plane,
        int planarConfiguration,
        int samples,
        int width,
        int height,
        int tileX,
        int tileY,
        int tileWidth,
        int tileHeight,
        OfficeRasterDecodeOptions options) {
        int copyWidth = Math.Min(tileWidth, width - tileX);
        int copyHeight = Math.Min(tileHeight, height - tileY);
        int tileSamples = planarConfiguration == 2 ? 1 : samples;
        for (int row = 0; row < copyHeight; row++) {
            if ((row & 31) == 0) options.CancellationToken.ThrowIfCancellationRequested();
            int sourceRow = row * tileWidth * tileSamples;
            int targetRow = checked(((tileY + row) * width + tileX) * samples);
            if (planarConfiguration == 1) {
                Buffer.BlockCopy(tile, sourceRow, interleaved, targetRow, copyWidth * samples);
            } else {
                for (int x = 0; x < copyWidth; x++) {
                    if ((x & 0xFFF) == 0) options.CancellationToken.ThrowIfCancellationRequested();
                    interleaved[targetRow + x * samples + plane] = tile[sourceRow + x];
                }
            }
        }
    }
}
