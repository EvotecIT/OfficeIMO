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
        if (retainPixels) source = new byte[sourceLength];

        if (hasStrips) {
            if (!TryReadScalarOrDefault(encodedBytes, entries, 278, littleEndian, height, out int rowsPerStrip) ||
                rowsPerStrip < 1) return false;
            int segmentsPerPlane = checked((height + rowsPerStrip - 1) / rowsPerStrip);
            int segmentCount = checked(segmentsPerPlane * (planarConfiguration == 2 ? samples : 1));
            if (!TryReadValues(encodedBytes, entries, 273, littleEndian, segmentCount, out int[] offsets) ||
                !TryReadValues(encodedBytes, entries, 279, littleEndian, segmentCount, out int[] byteCounts)) return false;

            for (int segment = 0; segment < segmentCount; segment++) {
                options.CancellationToken.ThrowIfCancellationRequested();
                if (!HasSegment(encodedBytes, offsets[segment], byteCounts[segment])) return false;
                int plane = planarConfiguration == 2 ? segment / segmentsPerPlane : 0;
                int strip = planarConfiguration == 2 ? segment % segmentsPerPlane : segment;
                int rowStart = checked(strip * rowsPerStrip);
                if (rowStart >= height) return false;
                int rows = Math.Min(rowsPerStrip, height - rowStart);
                int segmentSamples = planarConfiguration == 2 ? 1 : samples;
                int expected = checked(rows * width * segmentSamples);
                if (!retainPixels) continue;
                byte[] decoded = retainPixels && planarConfiguration == 1
                    ? source
                    : new byte[expected];
                int decodedOffset = retainPixels && planarConfiguration == 1
                    ? checked(rowStart * width * samples)
                    : 0;
                if (!TryDecodeStrip(encodedBytes, offsets[segment], byteCounts[segment], compression,
                        decoded, decodedOffset, expected)) return false;
                if (predictor == 2) ReverseHorizontalPredictor(decoded, decodedOffset, rows, width, segmentSamples);
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
        if (!TryReadValues(encodedBytes, entries, 324, littleEndian, tileSegmentCount, out int[] tileOffsets) ||
            !TryReadValues(encodedBytes, entries, 325, littleEndian, tileSegmentCount, out int[] tileByteCounts)) return false;

        int tileSamples = planarConfiguration == 2 ? 1 : samples;
        int tileByteLength = OfficeRasterGuards.EnsureByteCount(
            (long)tileWidth * tileHeight * tileSamples,
            "TIFF decoded tile exceeds the managed limit.");
        for (int segment = 0; segment < tileSegmentCount; segment++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (!HasSegment(encodedBytes, tileOffsets[segment], tileByteCounts[segment])) return false;
            if (!retainPixels) continue;
            int plane = planarConfiguration == 2 ? segment / segmentsPerTilePlane : 0;
            int tile = planarConfiguration == 2 ? segment % segmentsPerTilePlane : segment;
            int tileX = checked((tile % tilesAcross) * tileWidth);
            int tileY = checked((tile / tilesAcross) * tileHeight);
            var decoded = new byte[tileByteLength];
            if (!TryDecodeStrip(encodedBytes, tileOffsets[segment], tileByteCounts[segment], compression,
                    decoded, 0, tileByteLength)) return false;
            if (predictor == 2) ReverseHorizontalPredictor(decoded, 0, tileHeight, tileWidth, tileSamples);
            if (retainPixels) {
                CopyTile(decoded, source, plane, planarConfiguration, samples, width, height,
                    tileX, tileY, tileWidth, tileHeight, options);
            }
        }
        return true;
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
            for (int x = 0; x < width; x++) interleaved[targetRow + x * samples] = planeBytes[sourceRow + x];
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
                for (int x = 0; x < copyWidth; x++) interleaved[targetRow + x * samples + plane] = tile[sourceRow + x];
            }
        }
    }
}
