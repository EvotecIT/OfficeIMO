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
            if (!TryReadMbaTransform(bytes, range, expectedOutputChannels, pcsIsLab, out MbaTransform transform)) {
                return false;
            }
            transforms[index] = transform;
        }
        return true;
    }

    private interface IPcsToDeviceTransform {
        bool TryTransform(XyzValue pcsXyz, XyzValue whitePoint, out DeviceComponentValues components);
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
