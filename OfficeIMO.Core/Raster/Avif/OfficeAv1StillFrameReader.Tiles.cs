using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

internal static partial class OfficeAv1StillFrameReader {
    private static void ReadTileInfo(OfficeAv1Bits bits, OfficeAv1StillSequence sequence, OfficeAv1StillFrame frame) {
        int shift = sequence.Use128Superblock ? 5 : 4;
        int sbCols = (frame.MiCols + (1 << shift) - 1) >> shift;
        int sbRows = (frame.MiRows + (1 << shift) - 1) >> shift;
        int maxWidth = 4096 >> (shift + 2);
        int maxArea = (4096 * 2304) >> (2 * (shift + 2));
        int minColsLog2 = TileLog2(maxWidth, sbCols);
        int maxColsLog2 = TileLog2(1, Math.Min(sbCols, 64));
        int maxRowsLog2 = TileLog2(1, Math.Min(sbRows, 64));
        int minTilesLog2 = Math.Max(minColsLog2, TileLog2(maxArea, sbCols * sbRows));
        var columns = new List<int>(65);
        var rows = new List<int>(65);
        if (bits.Flag()) {
            frame.TileColsLog2 = ReadTileLog2(bits, minColsLog2, maxColsLog2);
            int width = (sbCols + (1 << frame.TileColsLog2) - 1) >> frame.TileColsLog2;
            for (int start = 0; start < sbCols; start += width) columns.Add(start << shift);
            frame.TileRowsLog2 = ReadTileLog2(bits, Math.Max(minTilesLog2 - frame.TileColsLog2, 0), maxRowsLog2);
            int height = (sbRows + (1 << frame.TileRowsLog2) - 1) >> frame.TileRowsLog2;
            for (int start = 0; start < sbRows; start += height) rows.Add(start << shift);
        } else {
            int widest = 0;
            for (int start = 0; start < sbCols;) {
                Require(columns.Count < 64);
                columns.Add(start << shift);
                int width = bits.NonSymmetric(Math.Min(sbCols - start, maxWidth)) + 1;
                widest = Math.Max(widest, width);
                start += width;
            }
            frame.TileColsLog2 = TileLog2(1, columns.Count);
            int area = minTilesLog2 > 0 ? (sbCols * sbRows) >> (minTilesLog2 + 1) : sbCols * sbRows;
            int maxHeight = Math.Max(area / widest, 1);
            for (int start = 0; start < sbRows;) {
                Require(rows.Count < 64);
                rows.Add(start << shift);
                start += bits.NonSymmetric(Math.Min(sbRows - start, maxHeight)) + 1;
            }
            frame.TileRowsLog2 = TileLog2(1, rows.Count);
        }
        Require(columns.Count > 0 && columns.Count <= 64 && rows.Count > 0 && rows.Count <= 64);
        columns.Add(frame.MiCols);
        rows.Add(frame.MiRows);
        frame.MiColStarts = columns.ToArray();
        frame.MiRowStarts = rows.ToArray();
        if (frame.TileColsLog2 + frame.TileRowsLog2 > 0) {
            frame.ContextUpdateTileId = bits.Read(frame.TileColsLog2 + frame.TileRowsLog2);
            frame.TileSizeBytes = bits.Read(2) + 1;
            Require(frame.ContextUpdateTileId < (columns.Count - 1) * (rows.Count - 1));
        }
    }

    private static int TileLog2(int blockSize, int target) {
        int result = 0;
        while ((blockSize << result) < target) result++;
        return result;
    }

    private static int ReadTileLog2(OfficeAv1Bits bits, int minimum, int maximum) {
        Require(minimum <= maximum);
        int value = minimum;
        while (value < maximum && bits.Flag()) value++;
        return value;
    }

    private static void ReadTileGroup(byte[] bytes, OfficeAv1Bits bits, OfficeAv1StillSequence sequence,
        OfficeAv1StillFrame frame, OfficeRasterDecodeOptions options) {
        int cols = frame.MiColStarts.Length - 1, rows = frame.MiRowStarts.Length - 1;
        int count = cols * rows;
        // OBU_FRAME requires tile_start_and_end_present_flag=0, even for an explicit full range.
        Require(count == 1 || !bits.Flag());
        bits.AlignZero();
        int p = bits.ByteOffset, end = sequence.FrameOffset + sequence.FrameLength;
        frame.TileGroupHeaderBytes = p - sequence.FrameOffset - frame.HeaderBytes;
        var tiles = new OfficeAv1Tile[count];
        for (int i = 0; i < count; i++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            int length;
            if (i == count - 1) {
                length = end - p;
            } else {
                Require(frame.TileSizeBytes >= 1 && frame.TileSizeBytes <= 4 && p <= end - frame.TileSizeBytes);
                uint sizeMinusOne = 0;
                for (int j = 0; j < frame.TileSizeBytes; j++) sizeMinusOne |= (uint)bytes[p++] << (j * 8);
                Require(sizeMinusOne < int.MaxValue);
                length = (int)sizeMinusOne + 1;
            }
            Require(length > 0 && p <= end - length);
            int row = i / cols, col = i % cols;
            tiles[i] = new OfficeAv1Tile(p, length, frame.MiRowStarts[row], frame.MiRowStarts[row + 1],
                frame.MiColStarts[col], frame.MiColStarts[col + 1]);
            p += length;
        }
        Require(p == end);
        frame.Tiles = tiles;
    }
}
