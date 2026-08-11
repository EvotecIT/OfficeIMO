using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Provides bounded, dependency-free conversion for matrix/TRC ICC color profiles.</summary>
/// <remarks>
/// Supports RGB and Gray input profiles whose profile connection space is XYZ. Profiles that
/// require multidimensional lookup tables are rejected so callers can select an explicit fallback.
/// </remarks>
public sealed class OfficeIccColorProfile {
    private const uint InputDeviceClassSignature = 0x73636E72U;
    private const uint DisplayDeviceClassSignature = 0x6D6E7472U;
    private const uint GraySignature = 0x47524159U;
    private const uint RgbSignature = 0x52474220U;
    private const uint XyzSignature = 0x58595A20U;
    private const uint CurveTypeSignature = 0x63757276U;
    private const uint ParametricCurveTypeSignature = 0x70617261U;
    private const uint AToB0TagSignature = 0x41324230U;
    private const uint AToB1TagSignature = 0x41324231U;
    private const uint AToB2TagSignature = 0x41324232U;
    private const uint DToB0TagSignature = 0x44324230U;
    private const uint DToB1TagSignature = 0x44324231U;
    private const uint DToB2TagSignature = 0x44324232U;
    private const uint DToB3TagSignature = 0x44324233U;
    private const int HeaderLength = 128;
    private const int TagTableHeaderLength = 4;
    private const int TagEntryLength = 12;
    private const int MaximumCurveEntries = 65536;

    private readonly ToneCurve _redCurve;
    private readonly ToneCurve _greenCurve;
    private readonly ToneCurve _blueCurve;
    private readonly XyzValue _redColumn;
    private readonly XyzValue _greenColumn;
    private readonly XyzValue _blueColumn;
    private readonly XyzValue _whitePoint;

    private OfficeIccColorProfile(
        int componentCount,
        ToneCurve redCurve,
        ToneCurve greenCurve,
        ToneCurve blueCurve,
        XyzValue redColumn,
        XyzValue greenColumn,
        XyzValue blueColumn,
        XyzValue whitePoint) {
        ComponentCount = componentCount;
        _redCurve = redCurve;
        _greenCurve = greenCurve;
        _blueCurve = blueCurve;
        _redColumn = redColumn;
        _greenColumn = greenColumn;
        _blueColumn = blueColumn;
        _whitePoint = whitePoint;
    }

    /// <summary>Gets the number of device components accepted by this profile.</summary>
    public int ComponentCount { get; }

    /// <summary>Attempts to parse a bounded RGB or Gray matrix/TRC ICC profile.</summary>
    public static bool TryCreate(byte[] profileBytes, out OfficeIccColorProfile? profile) {
        profile = null;
        if (profileBytes == null || !OfficeIccProfileValidator.TryValidate(profileBytes, 0, profileBytes.Length) ||
            !IsSupportedProfileClass(ReadUInt32(profileBytes, 12)) ||
            ReadUInt32(profileBytes, 20) != XyzSignature ||
            !TryReadXyz(profileBytes, 68, profileBytes.Length - 68, requireTypeHeader: false, out XyzValue whitePoint) ||
            !whitePoint.IsPositive) {
            return false;
        }

        if (!TryReadTagTable(profileBytes, out Dictionary<uint, TagRange> tags) ||
            HasUnsupportedDeviceToPcsTransform(tags)) return false;
        uint deviceColorSpace = ReadUInt32(profileBytes, 16);
        if (deviceColorSpace == GraySignature) {
            if (!TryReadToneCurve(profileBytes, tags, 0x6B545243U, out ToneCurve grayCurve) || // kTRC
                !TryReadXyzTag(profileBytes, tags, 0x77747074U, out XyzValue mediaWhite) || // wtpt
                !mediaWhite.IsPositive) return false;
            whitePoint = mediaWhite;
            profile = new OfficeIccColorProfile(
                1,
                grayCurve,
                ToneCurve.Identity,
                ToneCurve.Identity,
                whitePoint,
                default,
                default,
                whitePoint);
            return true;
        }

        if (deviceColorSpace != RgbSignature ||
            !TryReadToneCurve(profileBytes, tags, 0x72545243U, out ToneCurve redCurve) || // rTRC
            !TryReadToneCurve(profileBytes, tags, 0x67545243U, out ToneCurve greenCurve) || // gTRC
            !TryReadToneCurve(profileBytes, tags, 0x62545243U, out ToneCurve blueCurve) || // bTRC
            !TryReadXyzTag(profileBytes, tags, 0x7258595AU, out XyzValue redColumn) || // rXYZ
            !TryReadXyzTag(profileBytes, tags, 0x6758595AU, out XyzValue greenColumn) || // gXYZ
            !TryReadXyzTag(profileBytes, tags, 0x6258595AU, out XyzValue blueColumn)) { // bXYZ
            return false;
        }

        profile = new OfficeIccColorProfile(
            3,
            redCurve,
            greenCurve,
            blueCurve,
            redColumn,
            greenColumn,
            blueColumn,
            whitePoint);
        return true;
    }

    private static bool IsSupportedProfileClass(uint signature) =>
        signature == InputDeviceClassSignature || signature == DisplayDeviceClassSignature;

    /// <summary>Attempts to convert device components through the ICC profile to sRGB.</summary>
    public bool TryConvert(IReadOnlyList<double> components, out OfficeColor color) {
        color = OfficeColor.Black;
        if (components == null || components.Count < ComponentCount) return false;
        for (int index = 0; index < ComponentCount; index++) {
            if (!IsFinite(components[index])) return false;
        }

        if (ComponentCount == 1) {
            double level = _redCurve.Evaluate(Clamp01(components[0]));
            color = OfficeColorSpaceConverter.FromXyz(
                _redColumn.X * level,
                _redColumn.Y * level,
                _redColumn.Z * level,
                _whitePoint.X,
                _whitePoint.Y,
                _whitePoint.Z);
            return true;
        }

        double red = _redCurve.Evaluate(Clamp01(components[0]));
        double green = _greenCurve.Evaluate(Clamp01(components[1]));
        double blue = _blueCurve.Evaluate(Clamp01(components[2]));
        color = OfficeColorSpaceConverter.FromXyz(
            (_redColumn.X * red) + (_greenColumn.X * green) + (_blueColumn.X * blue),
            (_redColumn.Y * red) + (_greenColumn.Y * green) + (_blueColumn.Y * blue),
            (_redColumn.Z * red) + (_greenColumn.Z * green) + (_blueColumn.Z * blue),
            _whitePoint.X,
            _whitePoint.Y,
            _whitePoint.Z);
        return true;
    }

    private static bool TryReadTagTable(byte[] bytes, out Dictionary<uint, TagRange> tags) {
        tags = new Dictionary<uint, TagRange>();
        uint declaredCount = ReadUInt32(bytes, HeaderLength);
        if (declaredCount > int.MaxValue) return false;
        int count = (int)declaredCount;
        for (int index = 0; index < count; index++) {
            int entry = HeaderLength + TagTableHeaderLength + index * TagEntryLength;
            uint signature = ReadUInt32(bytes, entry);
            int offset = checked((int)ReadUInt32(bytes, entry + 4));
            int length = checked((int)ReadUInt32(bytes, entry + 8));
            tags[signature] = new TagRange(offset, length);
        }
        return true;
    }

    private static bool HasUnsupportedDeviceToPcsTransform(Dictionary<uint, TagRange> tags) =>
        tags.ContainsKey(AToB0TagSignature) ||
        tags.ContainsKey(AToB1TagSignature) ||
        tags.ContainsKey(AToB2TagSignature) ||
        tags.ContainsKey(DToB0TagSignature) ||
        tags.ContainsKey(DToB1TagSignature) ||
        tags.ContainsKey(DToB2TagSignature) ||
        tags.ContainsKey(DToB3TagSignature);

    private static bool TryReadXyzTag(byte[] bytes, Dictionary<uint, TagRange> tags, uint signature, out XyzValue value) {
        value = default;
        return tags.TryGetValue(signature, out TagRange range) &&
            TryReadXyz(bytes, range.Offset, range.Length, requireTypeHeader: true, out value);
    }

    private static bool TryReadXyz(byte[] bytes, int offset, int length, bool requireTypeHeader, out XyzValue value) {
        value = default;
        int valueOffset = requireTypeHeader ? 8 : 0;
        if (offset < 0 || length < valueOffset + 12 || offset > bytes.Length - length ||
            (requireTypeHeader && ReadUInt32(bytes, offset) != XyzSignature)) {
            return false;
        }

        double x = ReadS15Fixed16(bytes, offset + valueOffset);
        double y = ReadS15Fixed16(bytes, offset + valueOffset + 4);
        double z = ReadS15Fixed16(bytes, offset + valueOffset + 8);
        if (!IsFinite(x) || !IsFinite(y) || !IsFinite(z)) return false;
        value = new XyzValue(x, y, z);
        return true;
    }

    private static bool TryReadToneCurve(byte[] bytes, Dictionary<uint, TagRange> tags, uint signature, out ToneCurve curve) {
        curve = ToneCurve.Identity;
        if (!tags.TryGetValue(signature, out TagRange range) || range.Length < 12) return false;
        uint type = ReadUInt32(bytes, range.Offset);
        if (type == CurveTypeSignature) return TryReadSampledCurve(bytes, range, out curve);
        if (type == ParametricCurveTypeSignature) return TryReadParametricCurve(bytes, range, out curve);
        return false;
    }

    private static bool TryReadSampledCurve(byte[] bytes, TagRange range, out ToneCurve curve) {
        curve = ToneCurve.Identity;
        uint declaredCount = ReadUInt32(bytes, range.Offset + 8);
        if (declaredCount > MaximumCurveEntries) return false;
        int count = (int)declaredCount;
        if (12L + count * 2L > range.Length) return false;
        if (count == 0) return true;
        if (count == 1) {
            double gamma = ReadUInt16(bytes, range.Offset + 12) / 256D;
            if (!IsFinite(gamma) || gamma <= 0D) return false;
            curve = ToneCurve.FromGamma(gamma);
            return true;
        }

        var samples = new double[count];
        for (int index = 0; index < count; index++) {
            samples[index] = ReadUInt16(bytes, range.Offset + 12 + index * 2) / 65535D;
        }
        curve = ToneCurve.FromSamples(samples);
        return true;
    }

    private static bool TryReadParametricCurve(byte[] bytes, TagRange range, out ToneCurve curve) {
        curve = ToneCurve.Identity;
        int functionType = ReadUInt16(bytes, range.Offset + 8);
        int parameterCount = functionType switch { 0 => 1, 1 => 3, 2 => 4, 3 => 5, 4 => 7, _ => 0 };
        if (parameterCount == 0 || bytes[range.Offset + 10] != 0 || bytes[range.Offset + 11] != 0 ||
            12L + parameterCount * 4L > range.Length) return false;
        var parameters = new double[parameterCount];
        for (int index = 0; index < parameters.Length; index++) {
            parameters[index] = ReadS15Fixed16(bytes, range.Offset + 12 + index * 4);
            if (!IsFinite(parameters[index])) return false;
        }
        if (parameters[0] <= 0D ||
            (functionType > 0 && parameters[1] == 0D) ||
            !IsParametricCurveDefinedOnUnitInterval(functionType, parameters)) return false;
        curve = ToneCurve.FromParameters(functionType, parameters);
        return true;
    }

    private static bool IsParametricCurveDefinedOnUnitInterval(int functionType, double[] parameters) {
        if (functionType == 0) return true;

        double gamma = parameters[0];
        double a = parameters[1];
        double b = parameters[2];
        double branchStart = functionType <= 2
            ? Math.Max(0D, -b / a)
            : Math.Max(0D, parameters[4]);
        if (branchStart > 1D) return true;

        double startBase = functionType <= 2 && a > 0D && branchStart > 0D
            ? 0D
            : a * branchStart + b;
        double endBase = a + b;
        double minimumBase = Math.Min(startBase, endBase);
        if (minimumBase < 0D && gamma != Math.Truncate(gamma)) return false;

        double offset = functionType switch {
            2 => parameters[3],
            4 => parameters[5],
            _ => 0D
        };
        double start = Math.Pow(startBase, gamma) + offset;
        double end = Math.Pow(endBase, gamma) + offset;
        return IsFinite(start) && IsFinite(end);
    }

    private static ushort ReadUInt16(byte[] bytes, int offset) =>
        unchecked((ushort)((bytes[offset] << 8) | bytes[offset + 1]));

    private static uint ReadUInt32(byte[] bytes, int offset) =>
        unchecked(((uint)bytes[offset] << 24) |
                  ((uint)bytes[offset + 1] << 16) |
                  ((uint)bytes[offset + 2] << 8) |
                  bytes[offset + 3]);

    private static double ReadS15Fixed16(byte[] bytes, int offset) => unchecked((int)ReadUInt32(bytes, offset)) / 65536D;
    private static double Clamp01(double value) => value < 0D ? 0D : value > 1D ? 1D : value;
    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private readonly struct TagRange {
        internal TagRange(int offset, int length) {
            Offset = offset;
            Length = length;
        }
        internal int Offset { get; }
        internal int Length { get; }
    }

    private readonly struct XyzValue {
        internal XyzValue(double x, double y, double z) {
            X = x;
            Y = y;
            Z = z;
        }
        internal double X { get; }
        internal double Y { get; }
        internal double Z { get; }
        internal bool IsPositive => X > 0D && Y > 0D && Z > 0D;
    }

    private sealed class ToneCurve {
        internal static readonly ToneCurve Identity = new ToneCurve(0, Array.Empty<double>());
        private readonly int _functionType;
        private readonly double[] _values;

        private ToneCurve(int functionType, double[] values) {
            _functionType = functionType;
            _values = values;
        }

        internal static ToneCurve FromGamma(double gamma) => new ToneCurve(-1, new[] { gamma });
        internal static ToneCurve FromSamples(double[] samples) => new ToneCurve(-2, samples);
        internal static ToneCurve FromParameters(int functionType, double[] parameters) => new ToneCurve(functionType, parameters);

        internal double Evaluate(double value) {
            if (_functionType == 0 && _values.Length == 0) return value;
            if (_functionType == -1) return Clamp01(Math.Pow(value, _values[0]));
            if (_functionType == -2) {
                double position = value * (_values.Length - 1);
                int lower = (int)Math.Floor(position);
                if (lower >= _values.Length - 1) return _values[_values.Length - 1];
                double fraction = position - lower;
                return Clamp01(_values[lower] + ((_values[lower + 1] - _values[lower]) * fraction));
            }

            double g = _values[0];
            double a = _values.Length > 1 ? _values[1] : 1D;
            double b = _values.Length > 2 ? _values[2] : 0D;
            double c = _values.Length > 3 ? _values[3] : 0D;
            double d = _values.Length > 4 ? _values[4] : 0D;
            double e = _values.Length > 5 ? _values[5] : 0D;
            double f = _values.Length > 6 ? _values[6] : 0D;
            double result = _functionType switch {
                0 => Math.Pow(value, g),
                1 => value >= -b / a ? Math.Pow(a * value + b, g) : 0D,
                2 => value >= -b / a ? Math.Pow(a * value + b, g) + c : c,
                3 => value >= d ? Math.Pow(a * value + b, g) : c * value,
                4 => value >= d ? Math.Pow(a * value + b, g) + e : c * value + f,
                _ => value
            };
            return Clamp01(IsFinite(result) ? result : 0D);
        }
    }
}
