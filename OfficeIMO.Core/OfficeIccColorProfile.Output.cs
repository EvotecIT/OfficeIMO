using System;
using System.Collections;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeIccColorProfile {
    /// <summary>Gets whether the profile exposes a supported PCS-to-device output transform.</summary>
    public bool HasOutputTransform => SelectPcsToDeviceTransform(OfficeIccRenderingIntent.Perceptual) != null;

    /// <summary>Attempts to convert an sRGB color to the profile's device components.</summary>
    public bool TryConvertToDevice(
        OfficeColor color,
        out double[] deviceComponents) =>
        TryConvertToDevice(color, OfficeIccRenderingIntent.Perceptual, out deviceComponents);

    /// <summary>Attempts to convert an sRGB color to the profile's device components using the requested rendering intent.</summary>
    public bool TryConvertToDevice(
        OfficeColor color,
        OfficeIccRenderingIntent renderingIntent,
        out double[] deviceComponents) {
        deviceComponents = Array.Empty<double>();
        if (!TryConvertToDeviceComponents(color, renderingIntent, out DeviceComponentValues values)) return false;
        deviceComponents = values.ToArray();
        return true;
    }

    /// <summary>Attempts to write the profile's device components into a caller-provided destination.</summary>
    public bool TryConvertToDevice(
        OfficeColor color,
        double[] destination) =>
        TryConvertToDevice(color, OfficeIccRenderingIntent.Perceptual, destination);

    /// <summary>Attempts to write the profile's device components into a caller-provided destination using the requested rendering intent.</summary>
    public bool TryConvertToDevice(
        OfficeColor color,
        OfficeIccRenderingIntent renderingIntent,
        double[] destination) {
        if (destination == null || destination.Length < ComponentCount ||
            !TryConvertToDeviceComponents(color, renderingIntent, out DeviceComponentValues values)) return false;
        values.CopyTo(destination);
        return true;
    }

    /// <summary>Attempts to soft-proof an sRGB color through the profile's output and input transforms.</summary>
    public bool TrySoftProof(
        OfficeColor color,
        out OfficeColor proofedColor) =>
        TrySoftProof(color, OfficeIccRenderingIntent.Perceptual, out proofedColor);

    /// <summary>Attempts to soft-proof an sRGB color through the profile's output and input transforms using the requested rendering intent.</summary>
    public bool TrySoftProof(
        OfficeColor color,
        OfficeIccRenderingIntent renderingIntent,
        out OfficeColor proofedColor) {
        proofedColor = OfficeColor.Black;
        if (!TryConvertToDeviceComponents(color, renderingIntent, out DeviceComponentValues values) ||
            !TryConvertDeviceComponents(values, renderingIntent, out OfficeColor converted)) return false;
        proofedColor = OfficeColor.FromRgba(converted.R, converted.G, converted.B, color.A);
        return true;
    }

    private bool TryConvertToDeviceComponents(
        OfficeColor color,
        OfficeIccRenderingIntent renderingIntent,
        out DeviceComponentValues deviceComponents) {
        deviceComponents = default;
        if (renderingIntent < OfficeIccRenderingIntent.Perceptual ||
            renderingIntent > OfficeIccRenderingIntent.AbsoluteColorimetric) return false;
        IPcsToDeviceTransform? transform = SelectPcsToDeviceTransform(renderingIntent);
        if (transform == null) return false;
        OfficeColorSpaceConverter.ConvertRgbToXyz(
            color.R / 255D,
            color.G / 255D,
            color.B / 255D,
            _whitePoint.X,
            _whitePoint.Y,
            _whitePoint.Z,
            out double x,
            out double y,
            out double z);
        if (renderingIntent == OfficeIccRenderingIntent.AbsoluteColorimetric) {
            x *= _whitePoint.X / _mediaWhitePoint.X;
            y *= _whitePoint.Y / _mediaWhitePoint.Y;
            z *= _whitePoint.Z / _mediaWhitePoint.Z;
        }
        return transform.TryTransform(new XyzValue(x, y, z), _whitePoint, out deviceComponents);
    }

    private bool TryConvertDeviceComponents(
        DeviceComponentValues components,
        OfficeIccRenderingIntent renderingIntent,
        out OfficeColor color) {
        color = OfficeColor.Black;
        IDeviceToPcsTransform? transform = SelectDeviceToPcsTransform(renderingIntent);
        XyzValue pcsXyz;
        if (transform != null) {
            if (!transform.TryTransform(components, _whitePoint, out pcsXyz)) return false;
        } else if (ComponentCount == 1) {
            double level = _redCurve.Evaluate(Clamp01(components[0]));
            pcsXyz = new XyzValue(
                _redColumn.X * level,
                _redColumn.Y * level,
                _redColumn.Z * level);
        } else {
            double red = _redCurve.Evaluate(Clamp01(components[0]));
            double green = _greenCurve.Evaluate(Clamp01(components[1]));
            double blue = _blueCurve.Evaluate(Clamp01(components[2]));
            pcsXyz = new XyzValue(
                (_redColumn.X * red) + (_greenColumn.X * green) + (_blueColumn.X * blue),
                (_redColumn.Y * red) + (_greenColumn.Y * green) + (_blueColumn.Y * blue),
                (_redColumn.Z * red) + (_greenColumn.Z * green) + (_blueColumn.Z * blue));
        }

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

    private IPcsToDeviceTransform? SelectPcsToDeviceTransform(OfficeIccRenderingIntent renderingIntent) {
        if (_pcsToDeviceTransforms == null) return null;
        int index = renderingIntent switch {
            OfficeIccRenderingIntent.RelativeColorimetric or OfficeIccRenderingIntent.AbsoluteColorimetric => 1,
            OfficeIccRenderingIntent.Saturation => 2,
            _ => 0
        };
        return _pcsToDeviceTransforms[index] ?? _pcsToDeviceTransforms[0];
    }

    private static bool TryReadPcsToDeviceTransforms(
        byte[] bytes,
        Dictionary<uint, TagRange> tags,
        int expectedOutputChannels,
        bool pcsIsLab,
        out IPcsToDeviceTransform?[] transforms) {
        transforms = new IPcsToDeviceTransform?[3];
        uint[] signatures = { BToA0TagSignature, BToA1TagSignature, BToA2TagSignature };
        for (int index = 0; index < signatures.Length; index++) {
            if (!tags.TryGetValue(signatures[index], out TagRange range)) {
                if (index == 0) return false;
                continue;
            }
            if (TryReadLutPcsToDeviceTransform(
                    bytes,
                    range,
                    expectedOutputChannels,
                    pcsIsLab,
                    out LutPcsToDeviceTransform legacyTransform)) {
                transforms[index] = legacyTransform;
            } else if (TryReadMbaTransform(
                    bytes,
                    range,
                    expectedOutputChannels,
                    pcsIsLab,
                    out MbaTransform transform)) {
                transforms[index] = transform;
            } else {
                return false;
            }
        }
        return true;
    }

    private interface IPcsToDeviceTransform {
        long RetainedByteCount { get; }
        bool TryTransform(XyzValue pcsXyz, XyzValue whitePoint, out DeviceComponentValues components);
    }

    private static long RetainedTransformBytes(IPcsToDeviceTransform?[]? transforms) {
        if (transforms == null) return 0L;
        long total = checked(24L + transforms.LongLength * 8L);
        for (int index = 0; index < transforms.Length; index++) {
            if (transforms[index] != null) total = checked(total + transforms[index]!.RetainedByteCount);
        }
        return total;
    }

    private sealed class MatrixTrcPcsToDeviceTransform : IPcsToDeviceTransform {
        private readonly ToneCurve _redCurve;
        private readonly ToneCurve _greenCurve;
        private readonly ToneCurve _blueCurve;
        private readonly double _m00;
        private readonly double _m01;
        private readonly double _m02;
        private readonly double _m10;
        private readonly double _m11;
        private readonly double _m12;
        private readonly double _m20;
        private readonly double _m21;
        private readonly double _m22;

        private MatrixTrcPcsToDeviceTransform(
            ToneCurve redCurve,
            ToneCurve greenCurve,
            ToneCurve blueCurve,
            double m00,
            double m01,
            double m02,
            double m10,
            double m11,
            double m12,
            double m20,
            double m21,
            double m22) {
            _redCurve = redCurve;
            _greenCurve = greenCurve;
            _blueCurve = blueCurve;
            _m00 = m00;
            _m01 = m01;
            _m02 = m02;
            _m10 = m10;
            _m11 = m11;
            _m12 = m12;
            _m20 = m20;
            _m21 = m21;
            _m22 = m22;
        }

        public long RetainedByteCount => 128L;

        internal static bool TryCreate(
            ToneCurve redCurve,
            ToneCurve greenCurve,
            ToneCurve blueCurve,
            XyzValue red,
            XyzValue green,
            XyzValue blue,
            out MatrixTrcPcsToDeviceTransform transform) {
            transform = null!;
            if (!redCurve.IsInvertible || !greenCurve.IsInvertible || !blueCurve.IsInvertible) return false;
            double determinant =
                (red.X * ((green.Y * blue.Z) - (blue.Y * green.Z))) -
                (green.X * ((red.Y * blue.Z) - (blue.Y * red.Z))) +
                (blue.X * ((red.Y * green.Z) - (green.Y * red.Z)));
            if (!IsFinite(determinant) || Math.Abs(determinant) < 1E-12D) return false;
            double inverse = 1D / determinant;
            transform = new MatrixTrcPcsToDeviceTransform(
                redCurve,
                greenCurve,
                blueCurve,
                ((green.Y * blue.Z) - (blue.Y * green.Z)) * inverse,
                ((blue.X * green.Z) - (green.X * blue.Z)) * inverse,
                ((green.X * blue.Y) - (blue.X * green.Y)) * inverse,
                ((blue.Y * red.Z) - (red.Y * blue.Z)) * inverse,
                ((red.X * blue.Z) - (blue.X * red.Z)) * inverse,
                ((blue.X * red.Y) - (red.X * blue.Y)) * inverse,
                ((red.Y * green.Z) - (green.Y * red.Z)) * inverse,
                ((green.X * red.Z) - (red.X * green.Z)) * inverse,
                ((red.X * green.Y) - (green.X * red.Y)) * inverse);
            return true;
        }

        public bool TryTransform(XyzValue pcsXyz, XyzValue whitePoint, out DeviceComponentValues components) {
            double redLinear = (_m00 * pcsXyz.X) + (_m01 * pcsXyz.Y) + (_m02 * pcsXyz.Z);
            double greenLinear = (_m10 * pcsXyz.X) + (_m11 * pcsXyz.Y) + (_m12 * pcsXyz.Z);
            double blueLinear = (_m20 * pcsXyz.X) + (_m21 * pcsXyz.Y) + (_m22 * pcsXyz.Z);
            if (!IsFinite(redLinear) || !IsFinite(greenLinear) || !IsFinite(blueLinear)) {
                components = default;
                return false;
            }
            components = new DeviceComponentValues(
                3,
                _redCurve.EvaluateInverse(Clamp01(redLinear)),
                _greenCurve.EvaluateInverse(Clamp01(greenLinear)),
                _blueCurve.EvaluateInverse(Clamp01(blueLinear)),
                0D);
            return true;
        }
    }

    private readonly struct DeviceComponentValues : IReadOnlyList<double> {
        private readonly double _component0;
        private readonly double _component1;
        private readonly double _component2;
        private readonly double _component3;

        internal DeviceComponentValues(
            int count,
            double component0,
            double component1,
            double component2,
            double component3) {
            Count = count;
            _component0 = component0;
            _component1 = component1;
            _component2 = component2;
            _component3 = component3;
        }

        public int Count { get; }

        public double this[int index] => index switch {
            0 when Count > 0 => _component0,
            1 when Count > 1 => _component1,
            2 when Count > 2 => _component2,
            3 when Count > 3 => _component3,
            _ => throw new ArgumentOutOfRangeException(nameof(index))
        };

        internal double[] ToArray() => Count == 3
            ? new[] { _component0, _component1, _component2 }
            : new[] { _component0, _component1, _component2, _component3 };

        internal void CopyTo(double[] destination) {
            for (int index = 0; index < Count; index++) destination[index] = this[index];
        }

        public IEnumerator<double> GetEnumerator() {
            for (int index = 0; index < Count; index++) yield return this[index];
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
