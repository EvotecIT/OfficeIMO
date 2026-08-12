using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeIccColorProfile {
    private static bool TryReadMabTransform(
        byte[] bytes,
        TagRange range,
        int expectedInputChannels,
        bool pcsIsLab,
        out MabTransform transform) {
        transform = null!;
        if (range.Length < 32 ||
            ReadUInt32(bytes, range.Offset) != LutAToBTypeSignature ||
            !AreZero(bytes, range.Offset + 4, 4) ||
            !AreZero(bytes, range.Offset + 10, 2)) {
            return false;
        }

        int inputChannels = bytes[range.Offset + 8];
        int outputChannels = bytes[range.Offset + 9];
        if (inputChannels != expectedInputChannels || inputChannels is < 3 or > 4 || outputChannels != 3) return false;

        int bOffset = ReadRelativeOffset(bytes, range.Offset + 12);
        int matrixOffset = ReadRelativeOffset(bytes, range.Offset + 16);
        int mOffset = ReadRelativeOffset(bytes, range.Offset + 20);
        int clutOffset = ReadRelativeOffset(bytes, range.Offset + 24);
        int aOffset = ReadRelativeOffset(bytes, range.Offset + 28);
        if (bOffset <= 0 || matrixOffset < 0 || mOffset < 0 || clutOffset < 0 || aOffset < 0 ||
            (matrixOffset == 0) != (mOffset == 0) ||
            (clutOffset == 0) != (aOffset == 0) ||
            (clutOffset == 0 && inputChannels != outputChannels)) {
            return false;
        }

        var regions = new ElementRegion[12];
        int regionCount = 0;
        if (!TryReadEmbeddedCurveSet(bytes, range, bOffset, outputChannels, out ToneCurve[] bCurves, out ElementRange[] bRegions)) {
            return false;
        }
        AddCurveRegions(regions, ref regionCount, bRegions);

        ToneCurve[]? mCurves = null;
        double[]? matrix = null;
        if (matrixOffset != 0) {
            if (!TryReadMabMatrix(bytes, range, matrixOffset, out matrix, out ElementRange matrixRegion) ||
                !TryReadEmbeddedCurveSet(bytes, range, mOffset, outputChannels, out mCurves, out ElementRange[] mRegions)) {
                return false;
            }
            regions[regionCount++] = new ElementRegion(matrixRegion, isCurve: false);
            AddCurveRegions(regions, ref regionCount, mRegions);
        }

        ToneCurve[]? aCurves = null;
        MabClut? clut = null;
        if (clutOffset != 0) {
            if (!TryReadMabClut(bytes, range, clutOffset, inputChannels, outputChannels, out clut, out ElementRange clutRegion) ||
                !TryReadEmbeddedCurveSet(bytes, range, aOffset, inputChannels, out aCurves, out ElementRange[] aRegions)) {
                return false;
            }
            regions[regionCount++] = new ElementRegion(clutRegion, isCurve: false);
            AddCurveRegions(regions, ref regionCount, aRegions);
        }

        if (!HasValidMabElementLayout(bytes, range, regions, regionCount)) return false;
        transform = new MabTransform(inputChannels, aCurves, clut, mCurves, matrix, bCurves, pcsIsLab);
        return true;
    }

    private static bool TryReadMbaTransform(
        byte[] bytes,
        TagRange range,
        int expectedOutputChannels,
        bool pcsIsLab,
        out MbaTransform transform) {
        transform = null!;
        if (range.Length < 32 ||
            ReadUInt32(bytes, range.Offset) != LutBToATypeSignature ||
            !AreZero(bytes, range.Offset + 4, 4) ||
            !AreZero(bytes, range.Offset + 10, 2)) {
            return false;
        }

        int inputChannels = bytes[range.Offset + 8];
        int outputChannels = bytes[range.Offset + 9];
        if (inputChannels != 3 || outputChannels != expectedOutputChannels || outputChannels is < 3 or > 4) return false;

        int bOffset = ReadRelativeOffset(bytes, range.Offset + 12);
        int matrixOffset = ReadRelativeOffset(bytes, range.Offset + 16);
        int mOffset = ReadRelativeOffset(bytes, range.Offset + 20);
        int clutOffset = ReadRelativeOffset(bytes, range.Offset + 24);
        int aOffset = ReadRelativeOffset(bytes, range.Offset + 28);
        if (bOffset <= 0 || matrixOffset < 0 || mOffset < 0 || clutOffset < 0 || aOffset < 0 ||
            (matrixOffset == 0) != (mOffset == 0) ||
            (clutOffset == 0) != (aOffset == 0) ||
            (clutOffset == 0 && inputChannels != outputChannels)) {
            return false;
        }

        var regions = new ElementRegion[12];
        int regionCount = 0;
        if (!TryReadEmbeddedCurveSet(bytes, range, bOffset, inputChannels, out ToneCurve[] bCurves, out ElementRange[] bRegions)) {
            return false;
        }
        AddCurveRegions(regions, ref regionCount, bRegions);

        ToneCurve[]? mCurves = null;
        double[]? matrix = null;
        if (matrixOffset != 0) {
            if (!TryReadMabMatrix(bytes, range, matrixOffset, out matrix, out ElementRange matrixRegion) ||
                !TryReadEmbeddedCurveSet(bytes, range, mOffset, inputChannels, out mCurves, out ElementRange[] mRegions)) {
                return false;
            }
            regions[regionCount++] = new ElementRegion(matrixRegion, isCurve: false);
            AddCurveRegions(regions, ref regionCount, mRegions);
        }

        ToneCurve[]? aCurves = null;
        MabClut? clut = null;
        if (clutOffset != 0) {
            if (!TryReadMabClut(bytes, range, clutOffset, inputChannels, outputChannels, out clut, out ElementRange clutRegion) ||
                !TryReadEmbeddedCurveSet(bytes, range, aOffset, outputChannels, out aCurves, out ElementRange[] aRegions)) {
                return false;
            }
            regions[regionCount++] = new ElementRegion(clutRegion, isCurve: false);
            AddCurveRegions(regions, ref regionCount, aRegions);
        }

        if (!HasValidMabElementLayout(bytes, range, regions, regionCount)) return false;
        transform = new MbaTransform(outputChannels, bCurves, matrix, mCurves, clut, aCurves, pcsIsLab);
        return true;
    }

    private static bool TryReadEmbeddedCurveSet(
        byte[] bytes,
        TagRange tagRange,
        int relativeOffset,
        int count,
        out ToneCurve[] curves,
        out ElementRange[] regions) {
        curves = Array.Empty<ToneCurve>();
        regions = Array.Empty<ElementRange>();
        if (!TryGetElementStart(tagRange, relativeOffset, out int cursor)) return false;
        var parsed = new ToneCurve[count];
        var parsedRegions = new ElementRange[count];
        for (int index = 0; index < parsed.Length; index++) {
            int start = cursor;
            if (!TryReadEmbeddedCurve(bytes, tagRange, cursor, out parsed[index], out cursor)) return false;
            parsedRegions[index] = new ElementRange(start, cursor);
        }
        curves = parsed;
        regions = parsedRegions;
        return true;
    }

    private static void AddCurveRegions(ElementRegion[] target, ref int count, ElementRange[] curves) {
        for (int index = 0; index < curves.Length; index++) {
            target[count++] = new ElementRegion(curves[index], isCurve: true);
        }
    }

    private static bool TryReadEmbeddedCurve(
        byte[] bytes,
        TagRange tagRange,
        int offset,
        out ToneCurve curve,
        out int nextOffset) {
        curve = ToneCurve.Identity;
        nextOffset = 0;
        int tagEnd = checked(tagRange.Offset + tagRange.Length);
        if (offset < tagRange.Offset + 32 || offset > tagEnd - 12 || !AreZero(bytes, offset + 4, 4)) return false;

        uint type = ReadUInt32(bytes, offset);
        int size;
        if (type == CurveTypeSignature) {
            uint declaredCount = ReadUInt32(bytes, offset + 8);
            if (declaredCount > MaximumCurveEntries) return false;
            long required = 12L + declaredCount * 2L;
            if (required > int.MaxValue) return false;
            size = (int)required;
            if (offset > tagEnd - size || !TryReadSampledCurve(bytes, new TagRange(offset, size), out curve)) return false;
        } else if (type == ParametricCurveTypeSignature) {
            int functionType = ReadUInt16(bytes, offset + 8);
            int parameterCount = functionType switch { 0 => 1, 1 => 3, 2 => 4, 3 => 5, 4 => 7, _ => 0 };
            if (parameterCount == 0) return false;
            size = checked(12 + parameterCount * 4);
            if (offset > tagEnd - size || !TryReadParametricCurve(bytes, new TagRange(offset, size), out curve)) return false;
        } else {
            return false;
        }

        int paddedSize = Align4(size);
        if (offset > tagEnd - paddedSize || !AreZero(bytes, offset + size, paddedSize - size)) return false;
        nextOffset = checked(offset + paddedSize);
        return true;
    }

    private static bool TryReadMabMatrix(
        byte[] bytes,
        TagRange tagRange,
        int relativeOffset,
        out double[] matrix,
        out ElementRange region) {
        matrix = Array.Empty<double>();
        region = default;
        if (!TryGetElementStart(tagRange, relativeOffset, out int start)) return false;
        int tagEnd = checked(tagRange.Offset + tagRange.Length);
        const int matrixLength = 48;
        if (start > tagEnd - matrixLength) return false;
        var values = new double[12];
        for (int index = 0; index < values.Length; index++) {
            values[index] = ReadS15Fixed16(bytes, start + index * 4);
            if (!IsFinite(values[index])) return false;
        }
        matrix = values;
        region = new ElementRange(start, start + matrixLength);
        return true;
    }

    private static bool TryReadMabClut(
        byte[] bytes,
        TagRange tagRange,
        int relativeOffset,
        int inputChannels,
        int outputChannels,
        out MabClut clut,
        out ElementRange region) {
        clut = null!;
        region = default;
        if (!TryGetElementStart(tagRange, relativeOffset, out int start)) return false;
        int tagEnd = checked(tagRange.Offset + tagRange.Length);
        if (start > tagEnd - 20) return false;

        int grid0 = bytes[start];
        int grid1 = bytes[start + 1];
        int grid2 = bytes[start + 2];
        int grid3 = inputChannels == 4 ? bytes[start + 3] : 1;
        if (grid0 is < 2 or > MaximumMabClutGridPoints ||
            grid1 is < 2 or > MaximumMabClutGridPoints ||
            grid2 is < 2 or > MaximumMabClutGridPoints ||
            (inputChannels == 4 && grid3 is < 2 or > MaximumMabClutGridPoints)) return false;
        for (int index = inputChannels; index < 16; index++) {
            if (bytes[start + index] != 0) return false;
        }

        int precision = bytes[start + 16];
        if (precision is not (1 or 2) || !AreZero(bytes, start + 17, 3)) return false;
        long sampleCount = (long)grid0 * grid1 * grid2 * grid3;
        long dataLength = sampleCount * outputChannels * precision;
        long elementLength = 20L + dataLength;
        if (elementLength > int.MaxValue) return false;
        int paddedLength = Align4((int)elementLength);
        if (start > tagEnd - paddedLength || !AreZero(bytes, start + (int)elementLength, paddedLength - (int)elementLength)) {
            return false;
        }

        var payload = new byte[(int)dataLength];
        Buffer.BlockCopy(bytes, start + 20, payload, 0, payload.Length);
        clut = new MabClut(payload, inputChannels, outputChannels, grid0, grid1, grid2, grid3, precision);
        region = new ElementRange(start, start + paddedLength);
        return true;
    }

    private static bool HasValidMabElementLayout(
        byte[] bytes,
        TagRange tagRange,
        ElementRegion[] regions,
        int regionCount) {
        for (int left = 0; left < regionCount; left++) {
            for (int right = left + 1; right < regionCount; right++) {
                bool overlaps = regions[left].Range.Start < regions[right].Range.End &&
                    regions[right].Range.Start < regions[left].Range.End;
                bool exactlyShared = regions[left].IsCurve && regions[right].IsCurve &&
                    regions[left].Range.Start == regions[right].Range.Start &&
                    regions[left].Range.End == regions[right].Range.End;
                if (overlaps && !exactlyShared) return false;
            }
        }

        for (int index = 1; index < regionCount; index++) {
            ElementRegion current = regions[index];
            int insertion = index;
            while (insertion > 0 && regions[insertion - 1].Range.Start > current.Range.Start) {
                regions[insertion] = regions[insertion - 1];
                insertion--;
            }
            regions[insertion] = current;
        }

        int cursor = tagRange.Offset + 32;
        for (int index = 0; index < regionCount; index++) {
            ElementRange region = regions[index].Range;
            if (region.Start > cursor && !AreZero(bytes, cursor, region.Start - cursor)) return false;
            if (region.End > cursor) cursor = region.End;
        }

        int tagEnd = checked(tagRange.Offset + tagRange.Length);
        return cursor <= tagEnd && AreZero(bytes, cursor, tagEnd - cursor);
    }

    private static int ReadRelativeOffset(byte[] bytes, int offset) {
        uint value = ReadUInt32(bytes, offset);
        return value > int.MaxValue ? -1 : (int)value;
    }

    private static bool TryGetElementStart(TagRange range, int relativeOffset, out int start) {
        start = 0;
        if (relativeOffset < 32 || (relativeOffset & 3) != 0 || relativeOffset > range.Length - 4) return false;
        start = checked(range.Offset + relativeOffset);
        return true;
    }

    private static bool AreZero(byte[] bytes, int offset, int count) {
        for (int index = 0; index < count; index++) {
            if (bytes[offset + index] != 0) return false;
        }
        return true;
    }

    private static int Align4(int value) => checked((value + 3) & ~3);

    private readonly struct ElementRange {
        internal ElementRange(int start, int end) {
            Start = start;
            End = end;
        }

        internal int Start { get; }
        internal int End { get; }
    }

    private readonly struct ElementRegion {
        internal ElementRegion(ElementRange range, bool isCurve) {
            Range = range;
            IsCurve = isCurve;
        }

        internal ElementRange Range { get; }
        internal bool IsCurve { get; }
    }

    private sealed class MabTransform : IDeviceToPcsTransform {
        private const double PcsXyzScale = 65535D / 32768D;
        private readonly int _inputChannels;
        private readonly ToneCurve[]? _aCurves;
        private readonly MabClut? _clut;
        private readonly ToneCurve[]? _mCurves;
        private readonly double[]? _matrix;
        private readonly ToneCurve[] _bCurves;
        private readonly bool _pcsIsLab;

        internal MabTransform(
            int inputChannels,
            ToneCurve[]? aCurves,
            MabClut? clut,
            ToneCurve[]? mCurves,
            double[]? matrix,
            ToneCurve[] bCurves,
            bool pcsIsLab) {
            _inputChannels = inputChannels;
            _aCurves = aCurves;
            _clut = clut;
            _mCurves = mCurves;
            _matrix = matrix;
            _bCurves = bCurves;
            _pcsIsLab = pcsIsLab;
        }

        public bool TryTransform(IReadOnlyList<double> components, XyzValue whitePoint, out XyzValue pcsXyz) {
            pcsXyz = default;
            if (components.Count < _inputChannels) return false;
            return TryTransform(
                components[0],
                components[1],
                components[2],
                _inputChannels == 4 ? components[3] : 0D,
                whitePoint,
                out pcsXyz);
        }

        public bool TryTransform(DeviceComponentValues components, XyzValue whitePoint, out XyzValue pcsXyz) {
            pcsXyz = default;
            if (components.Count < _inputChannels) return false;
            return TryTransform(
                components[0],
                components[1],
                components[2],
                _inputChannels == 4 ? components[3] : 0D,
                whitePoint,
                out pcsXyz);
        }

        private bool TryTransform(
            double component0,
            double component1,
            double component2,
            double component3,
            XyzValue whitePoint,
            out XyzValue pcsXyz) {
            double value0 = Evaluate(_aCurves, 0, component0);
            double value1 = Evaluate(_aCurves, 1, component1);
            double value2 = Evaluate(_aCurves, 2, component2);
            double value3 = _inputChannels == 4 ? Evaluate(_aCurves, 3, component3) : 0D;

            if (_clut != null) {
                _clut.Interpolate(value0, value1, value2, value3, out value0, out value1, out value2, out _);
            }
            value0 = Evaluate(_mCurves, 0, value0);
            value1 = Evaluate(_mCurves, 1, value1);
            value2 = Evaluate(_mCurves, 2, value2);
            if (_matrix != null) {
                double matrix0 = Clamp01((_matrix[0] * value0) + (_matrix[1] * value1) + (_matrix[2] * value2) + _matrix[9]);
                double matrix1 = Clamp01((_matrix[3] * value0) + (_matrix[4] * value1) + (_matrix[5] * value2) + _matrix[10]);
                double matrix2 = Clamp01((_matrix[6] * value0) + (_matrix[7] * value1) + (_matrix[8] * value2) + _matrix[11]);
                value0 = matrix0;
                value1 = matrix1;
                value2 = matrix2;
            }
            value0 = _bCurves[0].Evaluate(Clamp01(value0));
            value1 = _bCurves[1].Evaluate(Clamp01(value1));
            value2 = _bCurves[2].Evaluate(Clamp01(value2));

            if (_pcsIsLab) {
                OfficeColorSpaceConverter.ConvertLabToXyz(
                    value0 * 100D,
                    value1 * 255D - 128D,
                    value2 * 255D - 128D,
                    whitePoint.X,
                    whitePoint.Y,
                    whitePoint.Z,
                    out double x,
                    out double y,
                    out double z);
                pcsXyz = new XyzValue(x, y, z);
            } else {
                pcsXyz = new XyzValue(
                    value0 * PcsXyzScale,
                    value1 * PcsXyzScale,
                    value2 * PcsXyzScale);
            }
            return true;
        }

        private static double Evaluate(ToneCurve[]? curves, int channel, double value) =>
            curves == null ? Clamp01(value) : curves[channel].Evaluate(Clamp01(value));
    }

    private sealed class MbaTransform : IPcsToDeviceTransform {
        private const double PcsXyzScale = 65535D / 32768D;
        private readonly int _outputChannels;
        private readonly ToneCurve[] _bCurves;
        private readonly double[]? _matrix;
        private readonly ToneCurve[]? _mCurves;
        private readonly MabClut? _clut;
        private readonly ToneCurve[]? _aCurves;
        private readonly bool _pcsIsLab;

        internal MbaTransform(
            int outputChannels,
            ToneCurve[] bCurves,
            double[]? matrix,
            ToneCurve[]? mCurves,
            MabClut? clut,
            ToneCurve[]? aCurves,
            bool pcsIsLab) {
            _outputChannels = outputChannels;
            _bCurves = bCurves;
            _matrix = matrix;
            _mCurves = mCurves;
            _clut = clut;
            _aCurves = aCurves;
            _pcsIsLab = pcsIsLab;
        }

        public bool TryTransform(XyzValue pcsXyz, XyzValue whitePoint, out DeviceComponentValues components) {
            components = default;
            double value0;
            double value1;
            double value2;
            if (_pcsIsLab) {
                ConvertXyzToLab(pcsXyz, whitePoint, out double lightness, out double a, out double b);
                value0 = lightness / 100D;
                value1 = (a + 128D) / 255D;
                value2 = (b + 128D) / 255D;
            } else {
                value0 = pcsXyz.X / PcsXyzScale;
                value1 = pcsXyz.Y / PcsXyzScale;
                value2 = pcsXyz.Z / PcsXyzScale;
            }

            value0 = _bCurves[0].Evaluate(Clamp01(value0));
            value1 = _bCurves[1].Evaluate(Clamp01(value1));
            value2 = _bCurves[2].Evaluate(Clamp01(value2));
            if (_matrix != null) {
                double matrix0 = Clamp01((_matrix[0] * value0) + (_matrix[1] * value1) + (_matrix[2] * value2) + _matrix[9]);
                double matrix1 = Clamp01((_matrix[3] * value0) + (_matrix[4] * value1) + (_matrix[5] * value2) + _matrix[10]);
                double matrix2 = Clamp01((_matrix[6] * value0) + (_matrix[7] * value1) + (_matrix[8] * value2) + _matrix[11]);
                value0 = matrix0;
                value1 = matrix1;
                value2 = matrix2;
            }
            value0 = Evaluate(_mCurves, 0, value0);
            value1 = Evaluate(_mCurves, 1, value1);
            value2 = Evaluate(_mCurves, 2, value2);

            double value3 = 0D;
            if (_clut != null) {
                _clut.Interpolate(value0, value1, value2, 0D, out value0, out value1, out value2, out value3);
            }
            value0 = Evaluate(_aCurves, 0, value0);
            value1 = Evaluate(_aCurves, 1, value1);
            value2 = Evaluate(_aCurves, 2, value2);
            value3 = _outputChannels == 4 ? Evaluate(_aCurves, 3, value3) : 0D;
            components = new DeviceComponentValues(_outputChannels, value0, value1, value2, value3);
            return true;
        }

        private static void ConvertXyzToLab(
            XyzValue xyz,
            XyzValue whitePoint,
            out double lightness,
            out double a,
            out double b) {
            double fx = LabFunction(xyz.X / whitePoint.X);
            double fy = LabFunction(xyz.Y / whitePoint.Y);
            double fz = LabFunction(xyz.Z / whitePoint.Z);
            lightness = Clamp(116D * fy - 16D, 0D, 100D);
            a = Clamp(500D * (fx - fy), -128D, 127D);
            b = Clamp(200D * (fy - fz), -128D, 127D);
        }

        private static double LabFunction(double value) => value > 216D / 24389D
            ? Math.Pow(Math.Max(0D, value), 1D / 3D)
            : (841D / 108D * value) + (4D / 29D);

        private static double Evaluate(ToneCurve[]? curves, int channel, double value) =>
            curves == null ? Clamp01(value) : curves[channel].Evaluate(Clamp01(value));

        private static double Clamp(double value, double minimum, double maximum) =>
            value < minimum ? minimum : value > maximum ? maximum : value;
    }

    private sealed class MabClut {
        private readonly byte[] _payload;
        private readonly int _inputChannels;
        private readonly int _outputChannels;
        private readonly int _grid0;
        private readonly int _grid1;
        private readonly int _grid2;
        private readonly int _grid3;
        private readonly int _precision;

        internal MabClut(
            byte[] payload,
            int inputChannels,
            int outputChannels,
            int grid0,
            int grid1,
            int grid2,
            int grid3,
            int precision) {
            _payload = payload;
            _inputChannels = inputChannels;
            _outputChannels = outputChannels;
            _grid0 = grid0;
            _grid1 = grid1;
            _grid2 = grid2;
            _grid3 = grid3;
            _precision = precision;
        }

        internal void Interpolate(
            double input0,
            double input1,
            double input2,
            double input3,
            out double output0,
            out double output1,
            out double output2,
            out double output3) {
            GetGridPosition(input0, _grid0, out int lower0, out double fraction0);
            GetGridPosition(input1, _grid1, out int lower1, out double fraction1);
            GetGridPosition(input2, _grid2, out int lower2, out double fraction2);
            GetGridPosition(input3, _grid3, out int lower3, out double fraction3);
            output0 = 0D;
            output1 = 0D;
            output2 = 0D;
            output3 = 0D;
            int cornerCount = 1 << _inputChannels;
            for (int corner = 0; corner < cornerCount; corner++) {
                double weight = 1D;
                int gridIndex = 0;
                for (int channel = 0; channel < _inputChannels; channel++) {
                    bool upper = (corner & (1 << channel)) != 0;
                    int lower = channel == 0 ? lower0 : channel == 1 ? lower1 : channel == 2 ? lower2 : lower3;
                    double fraction = channel == 0 ? fraction0 : channel == 1 ? fraction1 : channel == 2 ? fraction2 : fraction3;
                    int grid = channel == 0 ? _grid0 : channel == 1 ? _grid1 : channel == 2 ? _grid2 : _grid3;
                    weight *= upper ? fraction : 1D - fraction;
                    gridIndex = gridIndex * grid + lower + (upper ? 1 : 0);
                }
                if (weight == 0D) continue;
                int offset = gridIndex * _outputChannels * _precision;
                output0 += ReadNormalized(offset) * weight;
                output1 += ReadNormalized(offset + _precision) * weight;
                output2 += ReadNormalized(offset + 2 * _precision) * weight;
                if (_outputChannels == 4) output3 += ReadNormalized(offset + 3 * _precision) * weight;
            }
        }

        private static void GetGridPosition(double input, int grid, out int lower, out double fraction) {
            double position = Clamp01(input) * (grid - 1);
            lower = Math.Min((int)Math.Floor(position), grid - 2);
            fraction = position - lower;
        }

        private double ReadNormalized(int offset) => _precision == 1
            ? _payload[offset] / 255D
            : ReadUInt16(_payload, offset) / 65535D;
    }
}
