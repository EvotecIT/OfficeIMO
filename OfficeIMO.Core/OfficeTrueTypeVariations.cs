using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>First-party evaluator for TrueType gvar outlines and HVAR advance widths.</summary>
internal sealed class OfficeTrueTypeVariations {
    private const ushort GvarLongOffsets = 0x0001;
    private const ushort SharedPointNumbers = 0x8000;
    private const ushort EmbeddedPeakTuple = 0x8000;
    private const ushort IntermediateRegion = 0x4000;
    private const ushort PrivatePointNumbers = 0x2000;
    private const int MaximumTupleCount = 4095;
    private const int MaximumPointCount = 1_000_000;
    private const long MaximumGlyphTupleCells = 4_000_000;
    private const long MaximumOperationTupleCells = 16_000_000;

    private readonly OfficeOpenTypeReader _reader;
    private readonly OfficeFontVariationModel _model;
    private readonly int _gvar;
    private readonly int _gvarEnd;
    private readonly int _glyphData;
    private readonly uint[] _glyphOffsets;
    private readonly double[][] _sharedTuples;
    private readonly OfficeOpenTypeHvarMetrics? _hvar;

    private OfficeTrueTypeVariations(
        OfficeOpenTypeReader reader,
        OfficeFontVariationModel model,
        int gvar,
        int gvarEnd,
        int glyphData,
        uint[] glyphOffsets,
        double[][] sharedTuples,
        OfficeOpenTypeHvarMetrics? hvar) {
        _reader = reader;
        _model = model;
        _gvar = gvar;
        _gvarEnd = gvarEnd;
        _glyphData = glyphData;
        _glyphOffsets = glyphOffsets;
        _sharedTuples = sharedTuples;
        _hvar = hvar;
    }

    internal static OfficeTrueTypeVariations Parse(
        byte[] data,
        IReadOnlyDictionary<string, int> tables,
        OfficeFontVariationModel model,
        int glyphCount) {
        OfficeOpenTypeReader reader = OfficeOpenTypeReader.TryCreate(data)
            ?? throw new InvalidDataException("The variable TrueType sfnt directory is invalid.");
        if (!reader.TryGetTable("gvar", out int gvar, out int length)) {
            throw new NotSupportedException("The variable TrueType font does not contain a gvar outline table.");
        }
        int end = checked(gvar + length);
        if (length < 20 || reader.ReadUInt16(gvar) != 1 || reader.ReadUInt16(gvar + 2) != 0) {
            throw new InvalidDataException("The gvar table header is invalid.");
        }
        int axisCount = reader.ReadUInt16(gvar + 4);
        int sharedTupleCount = reader.ReadUInt16(gvar + 6);
        uint sharedTupleRelative = reader.ReadUInt32(gvar + 8);
        int gvarGlyphCount = reader.ReadUInt16(gvar + 12);
        int flags = reader.ReadUInt16(gvar + 14);
        uint glyphDataRelative = reader.ReadUInt32(gvar + 16);
        if (axisCount != model.AxisCount || gvarGlyphCount != glyphCount || sharedTupleCount > MaximumTupleCount
            || sharedTupleRelative > int.MaxValue || glyphDataRelative > int.MaxValue) {
            throw new InvalidDataException("The gvar table dimensions are invalid.");
        }
        int sharedTupleOffset = checked(gvar + (int)sharedTupleRelative);
        int glyphData = checked(gvar + (int)glyphDataRelative);
        bool longOffsets = (flags & GvarLongOffsets) != 0;
        int offsetSize = longOffsets ? 4 : 2;
        int offsetDirectoryEnd = checked(gvar + 20 + (glyphCount + 1) * offsetSize);
        int sharedTupleBytes = checked(sharedTupleCount * axisCount * 2);
        if (offsetDirectoryEnd > end ||
            sharedTupleCount > 0 && (sharedTupleOffset < offsetDirectoryEnd || sharedTupleOffset > end - sharedTupleBytes) ||
            glyphData < offsetDirectoryEnd || glyphData > end) {
            throw new InvalidDataException("The gvar table offsets are invalid.");
        }
        var offsets = new uint[glyphCount + 1];
        int cursor = gvar + 20;
        uint previous = 0;
        for (int index = 0; index < offsets.Length; index++) {
            uint value = longOffsets ? reader.ReadUInt32(cursor) : (uint)reader.ReadUInt16(cursor) * 2U;
            cursor += offsetSize;
            if (value < previous || value > end - glyphData) throw new InvalidDataException("A gvar glyph offset is invalid.");
            offsets[index] = value;
            previous = value;
        }
        long glyphDataEnd = (long)glyphData + offsets[offsets.Length - 1];
        long sharedTupleEnd = (long)sharedTupleOffset + sharedTupleBytes;
        if (glyphDataEnd > end || sharedTupleOffset < glyphDataEnd && glyphData < sharedTupleEnd) {
            throw new InvalidDataException("The gvar table data regions overlap.");
        }
        var sharedTuples = new double[sharedTupleCount][];
        cursor = sharedTupleOffset;
        for (int tuple = 0; tuple < sharedTupleCount; tuple++) {
            sharedTuples[tuple] = ReadTuple(reader, ref cursor, axisCount, end);
        }
        OfficeOpenTypeHvarMetrics? hvar = OfficeOpenTypeHvarMetrics.TryParse(reader, model);
        return new OfficeTrueTypeVariations(reader, model, gvar, end, glyphData, offsets, sharedTuples, hvar);
    }

    internal int AdvanceWidthDelta(ushort glyphId) => _hvar?.AdvanceWidthDelta(glyphId) ?? 0;

    internal bool HasHvar => _hvar != null;

    internal WorkBudget CreateWorkBudget() => new(Math.Min(
        MaximumOperationTupleCells,
        Math.Max(64L, 64L * _reader.Data.Length)));

    internal int GvarAdvanceWidthDelta(
        ushort glyphId,
        double[] xs,
        double[] ys,
        ushort[] contourEnds,
        int baseAdvanceWidth,
        int leftSideBearing,
        int xMin,
        WorkBudget workBudget,
        CancellationToken cancellationToken) {
        if (_hvar != null) return _hvar.AdvanceWidthDelta(glyphId);
        if (xs.Length != ys.Length || xs.Length > MaximumPointCount) {
            throw new InvalidDataException("A variable glyph metric point array is invalid.");
        }
        int outlinePointCount = xs.Length;
        var originalsX = new double[outlinePointCount + 4];
        var originalsY = new double[outlinePointCount + 4];
        Array.Copy(xs, originalsX, outlinePointCount);
        Array.Copy(ys, originalsY, outlinePointCount);
        double leftPhantom = xMin - leftSideBearing;
        originalsX[outlinePointCount] = leftPhantom;
        originalsX[outlinePointCount + 1] = leftPhantom + baseAdvanceWidth;
        var deltasX = new double[originalsX.Length];
        var deltasY = new double[originalsY.Length];
        ApplyGlyphTuples(
            glyphId,
            originalsX,
            originalsY,
            contourEnds,
            deltasX,
            deltasY,
            workBudget,
            cancellationToken);
        return checked((int)Math.Round(
            deltasX[outlinePointCount + 1] - deltasX[outlinePointCount],
            MidpointRounding.ToEven));
    }

    internal void ApplySimpleGlyph(
        ushort glyphId,
        double[] xs,
        double[] ys,
        ushort[] contourEnds,
        WorkBudget workBudget,
        CancellationToken cancellationToken) {
        if (xs.Length != ys.Length || xs.Length > MaximumPointCount) throw new InvalidDataException("A variable glyph point array is invalid.");
        int outlinePointCount = xs.Length;
        var originalsX = new double[outlinePointCount + 4];
        var originalsY = new double[outlinePointCount + 4];
        Array.Copy(xs, originalsX, outlinePointCount);
        Array.Copy(ys, originalsY, outlinePointCount);
        double[] deltasX = new double[originalsX.Length];
        double[] deltasY = new double[originalsY.Length];
        ApplyGlyphTuples(
            glyphId,
            originalsX,
            originalsY,
            contourEnds,
            deltasX,
            deltasY,
            workBudget,
            cancellationToken);
        for (int index = 0; index < outlinePointCount; index++) {
            xs[index] += deltasX[index];
            ys[index] += deltasY[index];
        }
    }

    internal void ApplyCompositeGlyph(
        ushort glyphId,
        double[] xs,
        double[] ys,
        WorkBudget workBudget,
        CancellationToken cancellationToken) {
        if (xs.Length != ys.Length || xs.Length > MaximumPointCount) throw new InvalidDataException("A variable composite glyph point array is invalid.");
        int componentCount = xs.Length;
        var originalsX = new double[componentCount + 4];
        var originalsY = new double[componentCount + 4];
        Array.Copy(xs, originalsX, componentCount);
        Array.Copy(ys, originalsY, componentCount);
        double[] deltasX = new double[originalsX.Length];
        double[] deltasY = new double[originalsY.Length];
        ApplyGlyphTuples(
            glyphId,
            originalsX,
            originalsY,
            Array.Empty<ushort>(),
            deltasX,
            deltasY,
            workBudget,
            cancellationToken);
        for (int index = 0; index < componentCount; index++) {
            xs[index] += deltasX[index];
            ys[index] += deltasY[index];
        }
    }

    private void ApplyGlyphTuples(
        ushort glyphId,
        double[] originalsX,
        double[] originalsY,
        ushort[] contourEnds,
        double[] accumulatedX,
        double[] accumulatedY,
        WorkBudget workBudget,
        CancellationToken cancellationToken) {
        uint startRelative = _glyphOffsets[glyphId];
        uint endRelative = _glyphOffsets[glyphId + 1];
        if (startRelative == endRelative) return;
        int start = checked(_glyphData + (int)startRelative);
        int end = checked(_glyphData + (int)endRelative);
        if (start < _glyphData || end > _gvarEnd || start > end - 4) throw new InvalidDataException("A gvar glyph variation record is invalid.");
        int tupleCountValue = _reader.ReadUInt16(start);
        int tupleCount = tupleCountValue & 0x0FFF;
        int serializedDataOffset = _reader.ReadUInt16(start + 2);
        if (tupleCount > MaximumTupleCount) throw new InvalidDataException("A gvar glyph has too many variation tuples.");
        long expandedCells = checked((long)tupleCount * originalsX.Length);
        long serializedBudget = Math.Max(64L, 64L * (end - start));
        if (expandedCells > Math.Min(MaximumGlyphTupleCells, serializedBudget)) {
            throw new InvalidDataException("A gvar glyph exceeds the bounded tuple-work budget.");
        }
        workBudget.Consume(expandedCells);
        int headerCursor = start + 4;
        int dataCursor = checked(start + serializedDataOffset);
        if (dataCursor < headerCursor || dataCursor > end) throw new InvalidDataException("A gvar serialized-data offset is invalid.");

        int[]? sharedPoints = null;
        if ((tupleCountValue & SharedPointNumbers) != 0) {
            sharedPoints = ReadPackedPoints(ref dataCursor, end, originalsX.Length, cancellationToken);
        }
        var tuples = new TupleHeader[tupleCount];
        for (int tupleIndex = 0; tupleIndex < tupleCount; tupleIndex++) {
            if (headerCursor > dataCursor - 4) throw new InvalidDataException("A gvar tuple header is truncated.");
            int dataSize = _reader.ReadUInt16(headerCursor);
            int tupleIndexValue = _reader.ReadUInt16(headerCursor + 2);
            headerCursor += 4;
            double[] peak;
            if ((tupleIndexValue & EmbeddedPeakTuple) != 0) peak = ReadTuple(_reader, ref headerCursor, _model.AxisCount, dataCursor);
            else {
                int sharedIndex = tupleIndexValue & 0x0FFF;
                if (sharedIndex >= _sharedTuples.Length) throw new InvalidDataException("A gvar shared-tuple index is invalid.");
                peak = _sharedTuples[sharedIndex];
            }
            double[]? intermediateStart = null;
            double[]? intermediateEnd = null;
            if ((tupleIndexValue & IntermediateRegion) != 0) {
                intermediateStart = ReadTuple(_reader, ref headerCursor, _model.AxisCount, dataCursor);
                intermediateEnd = ReadTuple(_reader, ref headerCursor, _model.AxisCount, dataCursor);
            }
            tuples[tupleIndex] = new TupleHeader(dataSize, tupleIndexValue, peak, intermediateStart, intermediateEnd);
        }
        if (headerCursor > dataCursor) throw new InvalidDataException("The gvar tuple headers overlap serialized data.");

        var xDeltas = new int[originalsX.Length];
        var yDeltas = new int[originalsX.Length];
        var tupleX = new double[originalsX.Length];
        var tupleY = new double[originalsY.Length];
        var touched = new bool[originalsX.Length];
        var touchedIndexes = new int[originalsX.Length];
        foreach (TupleHeader tuple in tuples) {
            cancellationToken.ThrowIfCancellationRequested();
            int tupleEnd = checked(dataCursor + tuple.DataSize);
            if (tupleEnd > end) throw new InvalidDataException("A gvar tuple data block is truncated.");
            double scalar = CalculateTupleScalar(tuple);
            int[]? points = (tuple.TupleIndex & PrivatePointNumbers) != 0
                ? ReadPackedPoints(ref dataCursor, tupleEnd, originalsX.Length, cancellationToken)
                : sharedPoints;
            int deltaCount = points?.Length ?? originalsX.Length;
            ReadPackedDeltas(ref dataCursor, tupleEnd, deltaCount, xDeltas, cancellationToken);
            ReadPackedDeltas(ref dataCursor, tupleEnd, deltaCount, yDeltas, cancellationToken);
            if (dataCursor != tupleEnd) throw new InvalidDataException("A gvar tuple contains trailing serialized data.");
            if (scalar != 0D) {
                Array.Clear(tupleX, 0, tupleX.Length);
                Array.Clear(tupleY, 0, tupleY.Length);
                Array.Clear(touched, 0, touched.Length);
                if (points == null) {
                    for (int index = 0; index < originalsX.Length; index++) {
                        if ((index & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                        tupleX[index] = xDeltas[index];
                        tupleY[index] = yDeltas[index];
                        touched[index] = true;
                    }
                } else {
                    for (int index = 0; index < points.Length; index++) {
                        if ((index & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                        int point = points[index];
                        tupleX[point] = xDeltas[index];
                        tupleY[point] = yDeltas[index];
                        touched[point] = true;
                    }
                    InterpolateUntouched(originalsX, tupleX, touched, contourEnds, touchedIndexes, cancellationToken);
                    InterpolateUntouched(originalsY, tupleY, touched, contourEnds, touchedIndexes, cancellationToken);
                }
                for (int index = 0; index < originalsX.Length; index++) {
                    if ((index & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                    accumulatedX[index] += tupleX[index] * scalar;
                    accumulatedY[index] += tupleY[index] * scalar;
                }
            }
            dataCursor = tupleEnd;
        }
    }

    private double CalculateTupleScalar(TupleHeader tuple) {
        double scalar = 1D;
        IReadOnlyList<double> coordinates = _model.NormalizedCoordinates;
        for (int axis = 0; axis < coordinates.Count; axis++) {
            scalar *= OfficeOpenTypeVariationRegion.CalculateTupleScalar(
                coordinates[axis],
                tuple.Peak[axis],
                tuple.IntermediateStart?[axis],
                tuple.IntermediateEnd?[axis]);
            if (scalar == 0D) return 0D;
        }
        return scalar;
    }

    private int[]? ReadPackedPoints(
        ref int cursor,
        int end,
        int totalPointCount,
        CancellationToken cancellationToken) {
        if (cursor >= end) throw new InvalidDataException("A gvar packed-point count is truncated.");
        int count = _reader.Data[cursor++];
        if ((count & 0x80) != 0) {
            if (cursor >= end) throw new InvalidDataException("A gvar packed-point count is truncated.");
            count = ((count & 0x7F) << 8) | _reader.Data[cursor++];
        }
        if (count == 0) return null;
        if (count > totalPointCount) throw new InvalidDataException("A gvar packed-point count exceeds the glyph point count.");
        var points = new int[count];
        int point = 0;
        int output = 0;
        while (output < count) {
            if ((output & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (cursor >= end) throw new InvalidDataException("A gvar packed-point run is truncated.");
            int control = _reader.Data[cursor++];
            int runCount = (control & 0x7F) + 1;
            bool words = (control & 0x80) != 0;
            if (runCount > count - output || cursor > end - runCount * (words ? 2 : 1)) {
                throw new InvalidDataException("A gvar packed-point run is invalid.");
            }
            for (int index = 0; index < runCount; index++) {
                int delta = words ? _reader.ReadUInt16(cursor) : _reader.Data[cursor];
                cursor += words ? 2 : 1;
                point = checked(point + delta);
                if (point >= totalPointCount || output > 0 && point <= points[output - 1]) {
                    throw new InvalidDataException("A gvar packed-point index is invalid.");
                }
                points[output++] = point;
            }
        }
        return points;
    }

    private void ReadPackedDeltas(
        ref int cursor,
        int end,
        int count,
        int[] deltas,
        CancellationToken cancellationToken) {
        if (count < 0 || count > deltas.Length) throw new InvalidDataException("A gvar packed-delta count is invalid.");
        Array.Clear(deltas, 0, count);
        int output = 0;
        while (output < count) {
            if ((output & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (cursor >= end) throw new InvalidDataException("A gvar packed-delta run is truncated.");
            int control = _reader.Data[cursor++];
            int runCount = (control & 0x3F) + 1;
            if (runCount > count - output) throw new InvalidDataException("A gvar packed-delta run is too long.");
            bool zero = (control & 0x80) != 0;
            bool words = (control & 0x40) != 0;
            if (zero) {
                output += runCount;
                continue;
            }
            int size = words ? 2 : 1;
            if (cursor > end - runCount * size) throw new InvalidDataException("A gvar packed-delta run is truncated.");
            for (int index = 0; index < runCount; index++) {
                deltas[output++] = words ? _reader.ReadInt16(cursor) : unchecked((sbyte)_reader.Data[cursor]);
                cursor += size;
            }
        }
    }

    private static void InterpolateUntouched(
        double[] coordinates,
        double[] deltas,
        bool[] touched,
        ushort[] contourEnds,
        int[] touchedIndexes,
        CancellationToken cancellationToken) {
        int contourStart = 0;
        for (int contour = 0; contour < contourEnds.Length; contour++) {
            cancellationToken.ThrowIfCancellationRequested();
            int contourEnd = contourEnds[contour];
            if (contourEnd < contourStart || contourEnd >= coordinates.Length) throw new InvalidDataException("A gvar contour endpoint is invalid.");
            int touchedCount = 0;
            for (int point = contourStart; point <= contourEnd; point++) {
                if ((point & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (touched[point]) touchedIndexes[touchedCount++] = point;
            }
            if (touchedCount == 1) {
                double delta = deltas[touchedIndexes[0]];
                for (int point = contourStart; point <= contourEnd; point++) if (!touched[point]) deltas[point] = delta;
            } else if (touchedCount > 1) {
                for (int index = 0; index < touchedCount; index++) {
                    int first = touchedIndexes[index];
                    int second = touchedIndexes[(index + 1) % touchedCount];
                    int point = first == contourEnd ? contourStart : first + 1;
                    while (point != second) {
                        if ((point & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                        deltas[point] = Interpolate(coordinates[point], coordinates[first], deltas[first], coordinates[second], deltas[second]);
                        point = point == contourEnd ? contourStart : point + 1;
                    }
                }
            }
            contourStart = contourEnd + 1;
        }
    }

    private static double Interpolate(double value, double first, double firstDelta, double second, double secondDelta) {
        if (first == second) return firstDelta == secondDelta ? firstDelta : 0D;
        if (first > second) {
            Swap(ref first, ref second);
            Swap(ref firstDelta, ref secondDelta);
        }
        if (value <= first) return firstDelta;
        if (value >= second) return secondDelta;
        return firstDelta + ((value - first) * (secondDelta - firstDelta) / (second - first));
    }

    private static void Swap(ref double left, ref double right) {
        double value = left;
        left = right;
        right = value;
    }

    private static double[] ReadTuple(OfficeOpenTypeReader reader, ref int cursor, int axisCount, int end) {
        if (cursor > end - checked(axisCount * 2)) throw new InvalidDataException("A gvar tuple coordinate array is truncated.");
        var tuple = new double[axisCount];
        for (int axis = 0; axis < axisCount; axis++) {
            tuple[axis] = reader.ReadF2Dot14(cursor);
            cursor += 2;
        }
        return tuple;
    }

    private readonly struct TupleHeader {
        internal TupleHeader(int dataSize, int tupleIndex, double[] peak, double[]? intermediateStart, double[]? intermediateEnd) {
            DataSize = dataSize;
            TupleIndex = tupleIndex;
            Peak = peak;
            IntermediateStart = intermediateStart;
            IntermediateEnd = intermediateEnd;
        }
        internal int DataSize { get; }
        internal int TupleIndex { get; }
        internal double[] Peak { get; }
        internal double[]? IntermediateStart { get; }
        internal double[]? IntermediateEnd { get; }
    }

    internal sealed class WorkBudget {
        private long _remainingCells;

        internal WorkBudget(long remainingCells) => _remainingCells = remainingCells;

        internal void Consume(long cells) {
            if (cells < 0 || cells > _remainingCells) {
                throw new InvalidDataException("Variable-font outline expansion exceeds the operation work budget.");
            }
            _remainingCells -= cells;
        }
    }

}
