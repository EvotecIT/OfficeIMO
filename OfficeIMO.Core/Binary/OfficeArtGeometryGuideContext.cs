using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing.Binary;

// This context deliberately has no output-device DPI. Pixel-based SG operands
// need that device context and cannot be inferred from the stored frame size.
internal sealed class OfficeArtGeometryGuideContext {
    private readonly IReadOnlyList<OfficeArtProperty> _properties;
    private readonly double _width, _height;
    internal OfficeArtGeometryGuideContext(IReadOnlyList<OfficeArtProperty> properties, double width, double height) {
        _properties = properties; _width = width; _height = height;
    }
    internal bool TryParameter(int id, out int value, out OfficeArtCustomPathFailure failure) {
        value = 0;
        failure = OfficeArtCustomPathFailure.InvalidGuide;
        long result;
        switch (id) {
            case 0x0140: result = HalfFloor((long)Scalar(0x0140, 0) + Scalar(0x0142, 21600)); break;
            case 0x0141: result = HalfFloor((long)Scalar(0x0141, 0) + Scalar(0x0143, 21600)); break;
            case 0x0142: result = (long)Scalar(0x0142, 21600) - Scalar(0x0140, 0); break;
            case 0x0143: result = (long)Scalar(0x0143, 21600) - Scalar(0x0141, 0); break;
            case >= 0x0147 and <= 0x014E: result = Scalar(id, 0); break;
            case 0x0153:
            case 0x0154: result = Scalar(id, int.MinValue); break;
            case 0x01FC:
                result = OfficeArtProperty.GetBoolean(_properties, 0x01FF, 0x00080000, 0x00000008) != false ? 1 : 0; break;
            case 0x04FC:
            case 0x04FD:
            case 0x04FE:
            case 0x04FF:
                double emus = (id is 0x04FC or 0x04FE ? _width : _height) * 12700D;
                if (double.IsNaN(emus) || double.IsInfinity(emus) || emus < 0) return false;
                double rounded = Math.Round(emus, MidpointRounding.AwayFromZero);
                if (id is 0x04FE or 0x04FF) rounded = Math.Floor(rounded / 2D);
                if (rounded > int.MaxValue) return false;
                result = (long)rounded;
                break;
            default: failure = OfficeArtCustomPathFailure.GuideParameter; return false;
        }
        if (result < int.MinValue || result > int.MaxValue) return false;
        value = (int)result; failure = OfficeArtCustomPathFailure.None; return true;
    }
    private static long HalfFloor(long value) => value / 2 - (value < 0 && value % 2 != 0 ? 1 : 0);

    private int Scalar(int id, int fallback) {
        OfficeArtProperty? property = OfficeArtProperty.GetScalarProperty(_properties, id);
        return property == null ? fallback : unchecked((int)property.Value);
    }
}
