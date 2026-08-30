using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Provides bounded, dependency-free conversion for supported ICC color profiles.</summary>
/// <remarks>
/// Supports RGB and Gray matrix/TRC input profiles, RGB and CMYK LUT8/LUT16 input transforms,
/// legacy LUT8/LUT16 output transforms, and bounded ICC v4 AToB/BToA transforms using the mAB/mBA
/// types. Unsupported transform types are rejected for explicit fallback.
/// </remarks>
public sealed partial class OfficeIccColorProfile {
    private const uint InputDeviceClassSignature = 0x73636E72U;
    private const uint DisplayDeviceClassSignature = 0x6D6E7472U;
    private const uint OutputDeviceClassSignature = 0x70727472U;
    private const uint GraySignature = 0x47524159U;
    private const uint RgbSignature = 0x52474220U;
    private const uint XyzSignature = 0x58595A20U;
    private const uint LabSignature = 0x4C616220U;
    private const uint CurveTypeSignature = 0x63757276U;
    private const uint ParametricCurveTypeSignature = 0x70617261U;
    private const uint Lut8TypeSignature = 0x6D667431U;
    private const uint Lut16TypeSignature = 0x6D667432U;
    private const uint LutAToBTypeSignature = 0x6D414220U;
    private const uint LutBToATypeSignature = 0x6D424120U;
    private const uint AToB0TagSignature = 0x41324230U;
    private const uint AToB1TagSignature = 0x41324231U;
    private const uint AToB2TagSignature = 0x41324232U;
    private const uint BToA0TagSignature = 0x42324130U;
    private const uint BToA1TagSignature = 0x42324131U;
    private const uint BToA2TagSignature = 0x42324132U;
    private const uint MediaWhitePointTagSignature = 0x77747074U;
    private const int HeaderLength = 128;
    private const int TagTableHeaderLength = 4;
    private const int TagEntryLength = 12;
    private const int MaximumCurveEntries = 65536;
    private const int MaximumMabClutGridPoints = 33;
    private const double D50X = 0.9642D;
    private const double D50Y = 1D;
    private const double D50Z = 0.8249D;
    private const double IlluminantTolerance = 0.001D;

    private readonly ToneCurve _redCurve;
    private readonly ToneCurve _greenCurve;
    private readonly ToneCurve _blueCurve;
    private readonly XyzValue _redColumn;
    private readonly XyzValue _greenColumn;
    private readonly XyzValue _blueColumn;
    private readonly XyzValue _whitePoint;
    private readonly XyzValue _mediaWhitePoint;
    private readonly IDeviceToPcsTransform?[]? _deviceToPcsTransforms;
    private readonly IPcsToDeviceTransform?[]? _pcsToDeviceTransforms;

    private OfficeIccColorProfile(
        int componentCount,
        OfficeIccProfileClass profileClass,
        ToneCurve redCurve,
        ToneCurve greenCurve,
        ToneCurve blueCurve,
        XyzValue redColumn,
        XyzValue greenColumn,
        XyzValue blueColumn,
        XyzValue whitePoint,
        XyzValue mediaWhitePoint,
        IPcsToDeviceTransform?[]? pcsToDeviceTransforms = null) {
        ComponentCount = componentCount;
        ProfileClass = profileClass;
        _redCurve = redCurve;
        _greenCurve = greenCurve;
        _blueCurve = blueCurve;
        _redColumn = redColumn;
        _greenColumn = greenColumn;
        _blueColumn = blueColumn;
        _whitePoint = whitePoint;
        _mediaWhitePoint = mediaWhitePoint;
        _deviceToPcsTransforms = null;
        _pcsToDeviceTransforms = pcsToDeviceTransforms;
        RetainedByteCount = checked(
            256L +
            redCurve.RetainedByteCount +
            greenCurve.RetainedByteCount +
            blueCurve.RetainedByteCount +
            RetainedTransformBytes(pcsToDeviceTransforms));
    }

    private OfficeIccColorProfile(
        int componentCount,
        OfficeIccProfileClass profileClass,
        IDeviceToPcsTransform?[] deviceToPcsTransforms,
        IPcsToDeviceTransform?[]? pcsToDeviceTransforms,
        XyzValue whitePoint,
        XyzValue mediaWhitePoint) {
        ComponentCount = componentCount;
        ProfileClass = profileClass;
        _redCurve = ToneCurve.Identity;
        _greenCurve = ToneCurve.Identity;
        _blueCurve = ToneCurve.Identity;
        _redColumn = default;
        _greenColumn = default;
        _blueColumn = default;
        _whitePoint = whitePoint;
        _mediaWhitePoint = mediaWhitePoint;
        _deviceToPcsTransforms = deviceToPcsTransforms;
        _pcsToDeviceTransforms = pcsToDeviceTransforms;
        RetainedByteCount = checked(
            256L +
            RetainedTransformBytes(deviceToPcsTransforms) +
            RetainedTransformBytes(pcsToDeviceTransforms));
    }

    /// <summary>Gets the number of device components accepted by this profile.</summary>
    public int ComponentCount { get; }

    /// <summary>Gets the device class declared by the ICC profile header.</summary>
    public OfficeIccProfileClass ProfileClass { get; }

    /// <summary>Gets a conservative byte count for the retained parsed representation.</summary>
    internal long RetainedByteCount { get; }

    internal bool HasUncertifiableShadingInputTransform => _deviceToPcsTransforms != null;

    internal bool HasUncertifiableShadingSoftProofTransform {
        get {
            if (_deviceToPcsTransforms != null || _pcsToDeviceTransforms == null) return true;
            for (int index = 0; index < _pcsToDeviceTransforms.Length; index++) {
                if (_pcsToDeviceTransforms[index] != null &&
                    _pcsToDeviceTransforms[index] is not MatrixTrcPcsToDeviceTransform) return true;
            }
            return false;
        }
    }

    /// <summary>Attempts to parse a bounded supported ICC input profile.</summary>
    public static bool TryCreate(byte[] profileBytes, out OfficeIccColorProfile? profile) {
        profile = null;
        uint profileConnectionSpace = profileBytes == null || profileBytes.Length < HeaderLength
            ? 0U
            : ReadUInt32(profileBytes, 20);
        if (profileBytes == null || !OfficeIccProfileValidator.TryValidate(profileBytes, 0, profileBytes.Length) ||
            !TryGetProfileClass(ReadUInt32(profileBytes, 12), out OfficeIccProfileClass profileClass) ||
            (profileConnectionSpace != XyzSignature && profileConnectionSpace != LabSignature) ||
            !TryReadXyz(profileBytes, 68, profileBytes.Length - 68, requireTypeHeader: false, out XyzValue whitePoint) ||
            !whitePoint.IsPositive || !IsD50Illuminant(whitePoint)) {
            return false;
        }

        if (!TryReadTagTable(profileBytes, out Dictionary<uint, TagRange> tags)) return false;
        bool hasAuthoredDeviceToPcsTransform = HasAuthoredDeviceToPcsTransform(tags);
        uint deviceColorSpace = ReadUInt32(profileBytes, 16);
        if (!hasAuthoredDeviceToPcsTransform &&
            deviceColorSpace == GraySignature && profileConnectionSpace == XyzSignature) {
            if (!TryReadToneCurve(profileBytes, tags, 0x6B545243U, out ToneCurve grayCurve) || // kTRC
                !TryReadXyzTag(profileBytes, tags, MediaWhitePointTagSignature, out XyzValue mediaWhitePoint) ||
                !mediaWhitePoint.IsNormalizedMediaWhitePoint) return false;
            profile = new OfficeIccColorProfile(
                1,
                profileClass,
                grayCurve,
                ToneCurve.Identity,
                ToneCurve.Identity,
                whitePoint,
                default,
                default,
                whitePoint,
                mediaWhitePoint);
            return true;
        }

        if (!hasAuthoredDeviceToPcsTransform &&
            deviceColorSpace == RgbSignature && profileConnectionSpace == XyzSignature &&
            TryReadToneCurve(profileBytes, tags, 0x72545243U, out ToneCurve redCurve) && // rTRC
            TryReadToneCurve(profileBytes, tags, 0x67545243U, out ToneCurve greenCurve) && // gTRC
            TryReadToneCurve(profileBytes, tags, 0x62545243U, out ToneCurve blueCurve) && // bTRC
            TryReadXyzTag(profileBytes, tags, 0x7258595AU, out XyzValue redColumn) && // rXYZ
            TryReadXyzTag(profileBytes, tags, 0x6758595AU, out XyzValue greenColumn) && // gXYZ
            TryReadXyzTag(profileBytes, tags, 0x6258595AU, out XyzValue blueColumn)) { // bXYZ
            if (!IsUsableRgbMatrix(redColumn, greenColumn, blueColumn, whitePoint) ||
                !TryReadXyzTag(profileBytes, tags, MediaWhitePointTagSignature, out XyzValue mediaWhitePoint) ||
                !mediaWhitePoint.IsNormalizedMediaWhitePoint) return false;
            IPcsToDeviceTransform?[]? outputTransforms;
            if (HasAuthoredPcsToDeviceTransform(tags)) {
                outputTransforms = TryReadPcsToDeviceTransforms(
                    profileBytes,
                    tags,
                    expectedOutputChannels: 3,
                    pcsIsLab: false,
                    out IPcsToDeviceTransform?[] parsedOutputTransforms)
                        ? parsedOutputTransforms
                        : null;
            } else {
                outputTransforms = MatrixTrcPcsToDeviceTransform.TryCreate(
                    redCurve,
                    greenCurve,
                    blueCurve,
                    redColumn,
                    greenColumn,
                    blueColumn,
                    out MatrixTrcPcsToDeviceTransform inverse)
                        ? new IPcsToDeviceTransform?[] { inverse, null, null }
                        : null;
            }
            profile = new OfficeIccColorProfile(
                3,
                profileClass,
                redCurve,
                greenCurve,
                blueCurve,
                redColumn,
                greenColumn,
                blueColumn,
                whitePoint,
                mediaWhitePoint,
                outputTransforms);
            return true;
        }

        int lutComponentCount = deviceColorSpace == RgbSignature ? 3 : deviceColorSpace == 0x434D594BU ? 4 : 0; // CMYK
        if (lutComponentCount != 0 &&
            TryReadDeviceToPcsTransforms(
                profileBytes,
                tags,
                lutComponentCount,
                profileConnectionSpace == LabSignature,
                out IDeviceToPcsTransform?[] transforms)) {
            IPcsToDeviceTransform?[]? outputTransforms = TryReadPcsToDeviceTransforms(
                profileBytes,
                tags,
                lutComponentCount,
                profileConnectionSpace == LabSignature,
                out IPcsToDeviceTransform?[] parsedOutputTransforms)
                    ? parsedOutputTransforms
                    : null;
            XyzValue mediaWhitePoint = TryReadXyzTag(profileBytes, tags, MediaWhitePointTagSignature, out XyzValue authoredMediaWhite) && authoredMediaWhite.IsPositive
                ? authoredMediaWhite
                : whitePoint;
            profile = new OfficeIccColorProfile(lutComponentCount, profileClass, transforms, outputTransforms, whitePoint, mediaWhitePoint);
            return true;
        }

        return false;
    }

    private static bool TryGetProfileClass(uint signature, out OfficeIccProfileClass profileClass) {
        switch (signature) {
            case InputDeviceClassSignature:
                profileClass = OfficeIccProfileClass.InputDevice;
                return true;
            case DisplayDeviceClassSignature:
                profileClass = OfficeIccProfileClass.DisplayDevice;
                return true;
            case OutputDeviceClassSignature:
                profileClass = OfficeIccProfileClass.OutputDevice;
                return true;
            default:
                profileClass = default;
                return false;
        }
    }

    private static long RetainedTransformBytes(IDeviceToPcsTransform?[]? transforms) {
        if (transforms == null) return 0L;
        long total = checked(24L + transforms.LongLength * 8L);
        for (int index = 0; index < transforms.Length; index++) {
            if (transforms[index] != null) total = checked(total + transforms[index]!.RetainedByteCount);
        }
        return total;
    }

    /// <summary>Attempts to convert device components through the ICC profile to sRGB.</summary>
    public bool TryConvert(IReadOnlyList<double> components, out OfficeColor color) {
        return TryConvert(components, OfficeIccRenderingIntent.Perceptual, out color);
    }

    /// <summary>Attempts to convert device components through the ICC profile to sRGB using the requested rendering intent.</summary>
    public bool TryConvert(
        IReadOnlyList<double> components,
        OfficeIccRenderingIntent renderingIntent,
        out OfficeColor color) {
        color = OfficeColor.Black;
        if (components == null || components.Count < ComponentCount ||
            renderingIntent < OfficeIccRenderingIntent.Perceptual ||
            renderingIntent > OfficeIccRenderingIntent.AbsoluteColorimetric) return false;
        for (int index = 0; index < ComponentCount; index++) {
            if (!IsFinite(components[index])) return false;
        }

        IDeviceToPcsTransform? transform = SelectDeviceToPcsTransform(renderingIntent);
        if (transform != null) {
            if (!transform.TryTransform(components, _whitePoint, out XyzValue pcsXyz)) return false;
            pcsXyz = ApplyRenderingIntentToPcs(pcsXyz, renderingIntent);
            color = OfficeColorSpaceConverter.FromXyz(
                pcsXyz.X,
                pcsXyz.Y,
                pcsXyz.Z,
                _whitePoint.X,
                _whitePoint.Y,
                _whitePoint.Z);
            return true;
        }

        if (ComponentCount == 1) {
            double level = _redCurve.Evaluate(Clamp01(components[0]));
            XyzValue pcsXyz = ApplyRenderingIntentToPcs(
                new XyzValue(
                    _redColumn.X * level,
                    _redColumn.Y * level,
                    _redColumn.Z * level),
                renderingIntent);
            color = OfficeColorSpaceConverter.FromXyz(
                pcsXyz.X,
                pcsXyz.Y,
                pcsXyz.Z,
                _whitePoint.X,
                _whitePoint.Y,
                _whitePoint.Z);
            return true;
        }

        double red = _redCurve.Evaluate(Clamp01(components[0]));
        double green = _greenCurve.Evaluate(Clamp01(components[1]));
        double blue = _blueCurve.Evaluate(Clamp01(components[2]));
        XyzValue matrixPcsXyz = ApplyRenderingIntentToPcs(
            new XyzValue(
                (_redColumn.X * red) + (_greenColumn.X * green) + (_blueColumn.X * blue),
                (_redColumn.Y * red) + (_greenColumn.Y * green) + (_blueColumn.Y * blue),
                (_redColumn.Z * red) + (_greenColumn.Z * green) + (_blueColumn.Z * blue)),
            renderingIntent);
        color = OfficeColorSpaceConverter.FromXyz(
            matrixPcsXyz.X,
            matrixPcsXyz.Y,
            matrixPcsXyz.Z,
            _whitePoint.X,
            _whitePoint.Y,
            _whitePoint.Z);
        return true;
    }

    private XyzValue ApplyRenderingIntentToPcs(
        XyzValue pcsXyz,
        OfficeIccRenderingIntent renderingIntent) =>
        renderingIntent == OfficeIccRenderingIntent.AbsoluteColorimetric
            ? new XyzValue(
                pcsXyz.X * (_mediaWhitePoint.X / _whitePoint.X),
                pcsXyz.Y * (_mediaWhitePoint.Y / _whitePoint.Y),
                pcsXyz.Z * (_mediaWhitePoint.Z / _whitePoint.Z))
            : pcsXyz;

    private IDeviceToPcsTransform? SelectDeviceToPcsTransform(OfficeIccRenderingIntent renderingIntent) {
        if (_deviceToPcsTransforms == null) return null;
        int index = renderingIntent switch {
            OfficeIccRenderingIntent.RelativeColorimetric or OfficeIccRenderingIntent.AbsoluteColorimetric => 1,
            OfficeIccRenderingIntent.Saturation => 2,
            _ => 0
        };
        return _deviceToPcsTransforms[index] ?? _deviceToPcsTransforms[0];
    }

    private static bool TryReadDeviceToPcsTransforms(
        byte[] bytes,
        Dictionary<uint, TagRange> tags,
        int expectedInputChannels,
        bool pcsIsLab,
        out IDeviceToPcsTransform?[] transforms) {
        transforms = new IDeviceToPcsTransform?[3];
        uint[] signatures = { AToB0TagSignature, AToB1TagSignature, AToB2TagSignature };
        for (int index = 0; index < signatures.Length; index++) {
            if (!tags.TryGetValue(signatures[index], out TagRange range)) {
                if (index == 0) return false;
                continue;
            }
            if (TryReadLutTransform(bytes, range, expectedInputChannels, pcsIsLab, out LutTransform lut)) {
                transforms[index] = lut;
            } else if (bytes[8] >= 4 &&
                TryReadMabTransform(bytes, range, expectedInputChannels, pcsIsLab, out MabTransform mab)) {
                transforms[index] = mab;
            } else {
                return false;
            }
        }
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

    private static bool HasAuthoredDeviceToPcsTransform(Dictionary<uint, TagRange> tags) =>
        tags.ContainsKey(AToB0TagSignature) ||
        tags.ContainsKey(AToB1TagSignature) ||
        tags.ContainsKey(AToB2TagSignature);

    private static bool HasAuthoredPcsToDeviceTransform(Dictionary<uint, TagRange> tags) =>
        tags.ContainsKey(BToA0TagSignature) ||
        tags.ContainsKey(BToA1TagSignature) ||
        tags.ContainsKey(BToA2TagSignature);

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

    private static bool TryReadLutTransform(
        byte[] bytes,
        TagRange range,
        int expectedInputChannels,
        bool pcsIsLab,
        out LutTransform transform) {
        transform = null!;
        if (range.Length < 52) return false;
        uint type = ReadUInt32(bytes, range.Offset);
        int precision = type == Lut8TypeSignature ? 1 : type == Lut16TypeSignature ? 2 : 0;
        if (precision == 0 || (precision == 1 && !pcsIsLab) ||
            bytes[range.Offset + 4] != 0 || bytes[range.Offset + 5] != 0 ||
            bytes[range.Offset + 6] != 0 || bytes[range.Offset + 7] != 0 || bytes[range.Offset + 11] != 0) return false;
        int inputChannels = bytes[range.Offset + 8];
        int outputChannels = bytes[range.Offset + 9];
        int gridPoints = bytes[range.Offset + 10];
        if (inputChannels != expectedInputChannels || outputChannels != 3 ||
            inputChannels is < 3 or > 4 || gridPoints is < 2 or > 33 ||
            !HasIdentityLutMatrix(bytes, range.Offset + 12)) return false;

        int inputEntries = precision == 1 ? 256 : ReadUInt16(bytes, range.Offset + 48);
        int outputEntries = precision == 1 ? 256 : ReadUInt16(bytes, range.Offset + 50);
        int tableOffset = precision == 1 ? 48 : 52;
        if (inputEntries is < 2 or > 4096 || outputEntries is < 2 or > 4096) return false;
        long gridSampleCount = 1;
        for (int channel = 0; channel < inputChannels; channel++) {
            gridSampleCount *= gridPoints;
            if (gridSampleCount > int.MaxValue) return false;
        }
        long inputBytes = (long)inputChannels * inputEntries * precision;
        long clutBytes = gridSampleCount * outputChannels * precision;
        long outputBytes = (long)outputChannels * outputEntries * precision;
        long requiredLength = tableOffset + inputBytes + clutBytes + outputBytes;
        if (requiredLength != range.Length || requiredLength > int.MaxValue) return false;

        var payload = new byte[(int)requiredLength];
        Buffer.BlockCopy(bytes, range.Offset, payload, 0, payload.Length);
        transform = new LutTransform(
            payload,
            inputChannels,
            outputChannels,
            gridPoints,
            inputEntries,
            outputEntries,
            precision,
            tableOffset,
            checked(tableOffset + (int)inputBytes),
            checked(tableOffset + (int)inputBytes + (int)clutBytes),
            pcsIsLab);
        return true;
    }

    private static bool HasIdentityLutMatrix(byte[] bytes, int offset) {
        for (int row = 0; row < 3; row++) {
            for (int column = 0; column < 3; column++) {
                int expected = row == column ? 65536 : 0;
                if (unchecked((int)ReadUInt32(bytes, offset + (row * 3 + column) * 4)) != expected) return false;
            }
        }
        return true;
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
            if (index > 0 && samples[index] < samples[index - 1]) return false;
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
            (functionType > 0 && parameters[1] <= 0D) ||
            !IsParametricCurveDefinedOnUnitInterval(functionType, parameters) ||
            !IsParametricCurveMonotonic(functionType, parameters)) return false;
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

    private static bool IsParametricCurveMonotonic(int functionType, double[] parameters) {
        if (functionType <= 2) return true;
        double slope = parameters[3];
        double boundary = parameters[4];
        if (boundary <= 0D) return true;
        if (slope < 0D) return false;
        if (boundary > 1D) return true;
        double high = Math.Pow(parameters[1] * boundary + parameters[2], parameters[0]);
        double low = slope * boundary;
        if (functionType == 4) {
            high += parameters[5];
            low += parameters[6];
        }
        return IsFinite(high) && IsFinite(low) && high >= low;
    }

    private static bool IsD50Illuminant(XyzValue value) =>
        Math.Abs(value.X - D50X) <= IlluminantTolerance &&
        Math.Abs(value.Y - D50Y) <= IlluminantTolerance &&
        Math.Abs(value.Z - D50Z) <= IlluminantTolerance;

    private static bool IsUsableRgbMatrix(XyzValue red, XyzValue green, XyzValue blue, XyzValue pcsWhite) {
        double scale = Math.Max(
            Math.Max(Math.Abs(red.X), Math.Max(Math.Abs(red.Y), Math.Abs(red.Z))),
            Math.Max(
                Math.Max(Math.Abs(green.X), Math.Max(Math.Abs(green.Y), Math.Abs(green.Z))),
                Math.Max(Math.Abs(blue.X), Math.Max(Math.Abs(blue.Y), Math.Abs(blue.Z)))));
        if (!IsFinite(scale) || scale == 0D) return false;
        double determinant =
            red.X * (green.Y * blue.Z - green.Z * blue.Y) -
            green.X * (red.Y * blue.Z - red.Z * blue.Y) +
            blue.X * (red.Y * green.Z - red.Z * green.Y);
        return IsFinite(determinant) &&
            Math.Abs(determinant) > scale * scale * scale * 1e-12D &&
            Math.Abs(red.X + green.X + blue.X - pcsWhite.X) <= IlluminantTolerance &&
            Math.Abs(red.Y + green.Y + blue.Y - pcsWhite.Y) <= IlluminantTolerance &&
            Math.Abs(red.Z + green.Z + blue.Z - pcsWhite.Z) <= IlluminantTolerance;
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
        internal bool IsNormalizedMediaWhitePoint => IsPositive && Y <= 1D + 1D / 65536D;
    }

    private interface IDeviceToPcsTransform {
        long RetainedByteCount { get; }
        bool TryTransform(IReadOnlyList<double> components, XyzValue whitePoint, out XyzValue pcsXyz);
        bool TryTransform(DeviceComponentValues components, XyzValue whitePoint, out XyzValue pcsXyz);
    }

    private sealed class LutTransform : IDeviceToPcsTransform {
        private const double PcsXyzScale = 65535D / 32768D;
        private readonly byte[] _payload;
        private readonly int _inputChannels;
        private readonly int _outputChannels;
        private readonly int _gridPoints;
        private readonly int _inputEntries;
        private readonly int _outputEntries;
        private readonly int _precision;
        private readonly int _inputOffset;
        private readonly int _clutOffset;
        private readonly int _outputOffset;
        private readonly bool _pcsIsLab;

        internal LutTransform(
            byte[] payload,
            int inputChannels,
            int outputChannels,
            int gridPoints,
            int inputEntries,
            int outputEntries,
            int precision,
            int inputOffset,
            int clutOffset,
            int outputOffset,
            bool pcsIsLab) {
            _payload = payload;
            _inputChannels = inputChannels;
            _outputChannels = outputChannels;
            _gridPoints = gridPoints;
            _inputEntries = inputEntries;
            _outputEntries = outputEntries;
            _precision = precision;
            _inputOffset = inputOffset;
            _clutOffset = clutOffset;
            _outputOffset = outputOffset;
            _pcsIsLab = pcsIsLab;
        }

        public long RetainedByteCount => checked(96L + _payload.LongLength);

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
            double input0 = LookupInput(component0, 0);
            double input1 = LookupInput(component1, 1);
            double input2 = LookupInput(component2, 2);
            double input3 = _inputChannels == 4 ? LookupInput(component3, 3) : 0D;
            InterpolateClut(input0, input1, input2, input3, out double output0, out double output1, out double output2);
            output0 = LookupTable(_outputOffset, _outputEntries, output0);
            output1 = LookupTable(_outputOffset + _outputEntries * _precision, _outputEntries, output1);
            output2 = LookupTable(_outputOffset + 2 * _outputEntries * _precision, _outputEntries, output2);
            if (_pcsIsLab) {
                double lightness = _precision == 1 ? output0 * 100D : output0 * (65535D / 65280D) * 100D;
                double a = _precision == 1 ? output1 * 255D - 128D : output1 * (65535D / 256D) - 128D;
                double b = _precision == 1 ? output2 * 255D - 128D : output2 * (65535D / 256D) - 128D;
                OfficeColorSpaceConverter.ConvertLabToXyz(
                    lightness,
                    a,
                    b,
                    whitePoint.X,
                    whitePoint.Y,
                    whitePoint.Z,
                    out double x,
                    out double y,
                    out double z);
                pcsXyz = new XyzValue(x, y, z);
            } else {
                pcsXyz = new XyzValue(
                    output0 * PcsXyzScale,
                    output1 * PcsXyzScale,
                    output2 * PcsXyzScale);
            }
            return true;
        }

        private double LookupInput(double component, int channel) =>
            LookupTable(
                _inputOffset + channel * _inputEntries * _precision,
                _inputEntries,
                Clamp01(component));

        private double LookupTable(int offset, int entries, double value) {
            double position = Clamp01(value) * (entries - 1);
            int lower = (int)Math.Floor(position);
            if (lower >= entries - 1) return ReadNormalized(offset + (entries - 1) * _precision);
            double fraction = position - lower;
            double left = ReadNormalized(offset + lower * _precision);
            double right = ReadNormalized(offset + (lower + 1) * _precision);
            return left + (right - left) * fraction;
        }

        private void InterpolateClut(
            double input0,
            double input1,
            double input2,
            double input3,
            out double output0,
            out double output1,
            out double output2) {
            GetGridPosition(input0, out int lower0, out double fraction0);
            GetGridPosition(input1, out int lower1, out double fraction1);
            GetGridPosition(input2, out int lower2, out double fraction2);
            GetGridPosition(input3, out int lower3, out double fraction3);
            output0 = 0D;
            output1 = 0D;
            output2 = 0D;
            int cornerCount = 1 << _inputChannels;
            for (int corner = 0; corner < cornerCount; corner++) {
                double weight = 1D;
                int gridIndex = 0;
                for (int channel = 0; channel < _inputChannels; channel++) {
                    bool upper = (corner & (1 << channel)) != 0;
                    int lower = channel == 0 ? lower0 : channel == 1 ? lower1 : channel == 2 ? lower2 : lower3;
                    double fraction = channel == 0 ? fraction0 : channel == 1 ? fraction1 : channel == 2 ? fraction2 : fraction3;
                    weight *= upper ? fraction : 1D - fraction;
                    gridIndex = gridIndex * _gridPoints + lower + (upper ? 1 : 0);
                }
                if (weight == 0D) continue;
                int valueOffset = _clutOffset + gridIndex * _outputChannels * _precision;
                output0 += ReadNormalized(valueOffset) * weight;
                output1 += ReadNormalized(valueOffset + _precision) * weight;
                output2 += ReadNormalized(valueOffset + 2 * _precision) * weight;
            }
        }

        private void GetGridPosition(double input, out int lower, out double fraction) {
            double position = Clamp01(input) * (_gridPoints - 1);
            lower = Math.Min((int)Math.Floor(position), _gridPoints - 2);
            fraction = position - lower;
        }

        private double ReadNormalized(int offset) => _precision == 1
            ? _payload[offset] / 255D
            : ReadUInt16(_payload, offset) / 65535D;
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

        internal long RetainedByteCount => checked(32L + _values.LongLength * 8L);

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

        internal bool IsInvertible {
            get {
                if (_functionType == 0 && _values.Length == 0) return true;
                if (_functionType == -1) return _values[0] > 0D;
                if (_functionType == -2) {
                    for (int index = 1; index < _values.Length; index++) {
                        if (_values[index] < _values[index - 1]) return false;
                    }
                    return _values[_values.Length - 1] - _values[0] > 1E-9D;
                }

                double a = _values.Length > 1 ? _values[1] : 1D;
                double b = _values.Length > 2 ? _values[2] : 0D;
                double c = _values.Length > 3 ? _values[3] : 0D;
                double d = _values.Length > 4 ? _values[4] : 0D;
                double e = _values.Length > 5 ? _values[5] : 0D;
                double f = _values.Length > 6 ? _values[6] : 0D;
                if (a <= 0D || (_functionType >= 3 && c < 0D)) return false;
                if (_functionType >= 3 && d >= 0D && d <= 1D) {
                    double baseValue = (a * d) + b;
                    if (baseValue < 0D) return false;
                    double lower = (c * d) + (_functionType == 4 ? f : 0D);
                    double upper = Math.Pow(baseValue, _values[0]) + (_functionType == 4 ? e : 0D);
                    if (!IsFinite(lower) || !IsFinite(upper) ||
                        Math.Abs(upper - lower) > 2D / 65536D) return false;
                }
                return Evaluate(1D) - Evaluate(0D) > 1E-9D;
            }
        }

        internal double EvaluateInverse(double value) {
            value = Clamp01(value);
            if (_functionType == 0 && _values.Length == 0) return value;
            if (_functionType == -1) return Clamp01(Math.Pow(value, 1D / _values[0]));
            if (_functionType == -2) {
                int lower = 0;
                int upper = _values.Length - 1;
                if (value <= _values[lower]) return 0D;
                if (value >= _values[upper]) return 1D;
                while (upper - lower > 1) {
                    int middle = lower + ((upper - lower) / 2);
                    if (_values[middle] <= value) lower = middle;
                    else upper = middle;
                }
                double span = _values[upper] - _values[lower];
                double fraction = span <= 1E-15D ? 0D : (value - _values[lower]) / span;
                return Clamp01((lower + fraction) / (_values.Length - 1D));
            }

            double low = 0D;
            double high = 1D;
            for (int iteration = 0; iteration < 28; iteration++) {
                double middle = (low + high) * 0.5D;
                if (Evaluate(middle) < value) low = middle;
                else high = middle;
            }
            return (low + high) * 0.5D;
        }
    }
}
