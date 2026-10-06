using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>Bounded AAT normal-track evaluator. Values remain in design units until the paint size is known.</summary>
internal sealed class OfficeOpenTypeTracking {
    private const int MaximumTracks = 64;
    private const int MaximumSizes = 256;
    private readonly double[] _sizes;
    private readonly double[] _adjustments;

    private OfficeOpenTypeTracking(double[] sizes, double[] adjustments) {
        _sizes = sizes;
        _adjustments = adjustments;
    }

    internal static OfficeOpenTypeTracking? Parse(byte[] data, int offset, int length) {
        if (offset < 0 || length < 12 || offset > data.Length - length)
            throw Invalid("header is truncated");
        if (U32(data, offset) != 0x00010000 || U16(data, offset + 4) != 0 || U16(data, offset + 10) != 0)
            throw Invalid("header is unsupported");
        int horizontal = U16(data, offset + 6);
        int vertical = U16(data, offset + 8);
        // Validate both directories; this evaluator currently applies horizontal normal tracking only.
        OfficeOpenTypeTracking? result = horizontal == 0 ? null : ParseDirection(data, offset, length, horizontal);
        if (vertical != 0) ParseDirection(data, offset, length, vertical);
        return result;
    }

    private static OfficeOpenTypeTracking ParseDirection(byte[] data, int start, int length, int relative) {
        if (relative < 12 || (relative & 3) != 0 || relative > length - 8)
            throw Invalid("TrackData offset is invalid");
        int at = start + relative;
        int tracks = U16(data, at);
        int sizes = U16(data, at + 2);
        if (tracks < 1 || tracks > MaximumTracks || sizes < 1 || sizes > MaximumSizes
            || relative + 8 + tracks * 8 > length)
            throw Invalid("TrackData exceeds its bounds");
        uint sizesRelative = U32(data, at + 4);
        int recordsEnd = relative + 8 + tracks * 8;
        if (sizes * 4 > length || sizesRelative < recordsEnd || (sizesRelative & 3) != 0 || sizesRelative > (uint)(length - sizes * 4))
            throw Invalid("size table offset is invalid");
        var sizeValues = new double[sizes];
        for (int index = 0; index < sizes; index++) {
            double size = Fixed(data, start + (int)sizesRelative + index * 4);
            if (size <= 0 || index > 0 && size <= sizeValues[index - 1])
                throw Invalid("sizes are not positive and strictly sorted");
            sizeValues[index] = size;
        }
        var trackValues = new double[tracks];
        var valueOffsets = new int[tracks];
        for (int index = 0; index < tracks; index++) {
            int record = at + 8 + index * 8;
            double track = Fixed(data, record);
            int name = U16(data, record + 4);
            int values = U16(data, record + 6);
            if (index > 0 && track <= trackValues[index - 1] || name <= 255 || name >= 32768
                || values < recordsEnd || (values & 1) != 0 || values > length - sizes * 2
                || values < sizesRelative + sizes * 4 && sizesRelative < values + sizes * 2)
                throw Invalid("track entry is invalid");
            trackValues[index] = track;
            valueOffsets[index] = start + values;
        }
        int lower = Interval(trackValues, 0D);
        int upper = Math.Min(lower + 1, tracks - 1);
        double weight = upper == lower ? 0D : -trackValues[lower] / (trackValues[upper] - trackValues[lower]);
        var adjustments = new double[sizes];
        for (int index = 0; index < sizes; index++) {
            short first = unchecked((short)U16(data, valueOffsets[lower] + index * 2));
            short second = unchecked((short)U16(data, valueOffsets[upper] + index * 2));
            adjustments[index] = first + (second - first) * weight;
        }
        return new OfficeOpenTypeTracking(sizeValues, adjustments);
    }

    internal double GetAdjustment(double fontSize) {
        if (fontSize <= 0 || double.IsNaN(fontSize) || double.IsInfinity(fontSize)) return 0D;
        int lower = Interval(_sizes, fontSize);
        int upper = Math.Min(lower + 1, _sizes.Length - 1);
        if (upper == lower) return _adjustments[lower];
        double weight = (fontSize - _sizes[lower]) / (_sizes[upper] - _sizes[lower]);
        return _adjustments[lower] + (_adjustments[upper] - _adjustments[lower]) * weight;
    }

    // Place one advance at the edge of each shaped grapheme, including a ligature's single cluster.
    internal static bool[] GetBoundaries(string text, IReadOnlyList<int> indexes, bool negative = false) {
        int[] starts = StringInfo.ParseCombiningCharacters(text);
        var clusters = new int[indexes.Count];
        for (int index = 0; index < indexes.Count; index++) {
            int cluster = Array.BinarySearch(starts, indexes[index]);
            clusters[index] = cluster < 0 ? ~cluster - 1 : cluster;
        }
        var boundaries = new bool[indexes.Count];
        for (int index = 0; index < indexes.Count; index++) {
            int adjacent = negative ? index - 1 : index + 1;
            boundaries[index] = adjacent < 0 || adjacent >= indexes.Count || clusters[index] != clusters[adjacent];
        }
        return boundaries;
    }

    private static int Interval(double[] values, double value) {
        if (values.Length == 1) return 0;
        int index = Array.BinarySearch(values, value);
        if (index < 0) index = ~index - 1;
        return Math.Max(0, Math.Min(index, values.Length - 2));
    }

    private static ushort U16(byte[] data, int offset) => (ushort)((data[offset] << 8) | data[offset + 1]);
    private static uint U32(byte[] data, int offset) =>
        ((uint)U16(data, offset) << 16) | U16(data, offset + 2);
    private static double Fixed(byte[] data, int offset) => unchecked((int)U32(data, offset)) / 65536D;
    private static InvalidDataException Invalid(string reason) => new("The AAT trak " + reason + ".");
}
