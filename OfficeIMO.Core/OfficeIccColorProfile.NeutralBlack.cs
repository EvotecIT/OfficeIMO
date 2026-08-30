using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeIccColorProfile {
    internal bool CanDeriveNeutralBlack(OfficeIccRenderingIntent renderingIntent) =>
        ComponentCount == 4 &&
        SelectPcsToDeviceTransform(renderingIntent) != null &&
        SelectDeviceToPcsTransform(renderingIntent) != null;

    internal bool TryDeriveNeutralBlack(
        OfficeColor color,
        OfficeIccRenderingIntent renderingIntent,
        out double black) {
        black = 0D;
        if (color.R != color.G || color.G != color.B ||
            !CanDeriveNeutralBlack(renderingIntent) ||
            !TryConvertToDeviceComponents(color, renderingIntent, out DeviceComponentValues profiledComponents) ||
            !TryGetDevicePcsLuminance(profiledComponents, renderingIntent, out double targetLuminance) ||
            !TryGetBlackAxisLuminance(0D, renderingIntent, out double zeroBlackLuminance) ||
            !TryGetBlackAxisLuminance(1D, renderingIntent, out double fullBlackLuminance)) {
            return false;
        }

        bool descending = zeroBlackLuminance >= fullBlackLuminance;
        double previousLuminance = zeroBlackLuminance;
        for (int sample = 1; sample <= 16; sample++) {
            if (!TryGetBlackAxisLuminance(sample / 16D, renderingIntent, out double sampledLuminance)) return false;
            if ((descending && sampledLuminance > previousLuminance + 0.000001D) ||
                (!descending && sampledLuminance < previousLuminance - 0.000001D)) {
                return false;
            }
            previousLuminance = sampledLuminance;
        }
        double minimum = Math.Min(zeroBlackLuminance, fullBlackLuminance);
        double maximum = Math.Max(zeroBlackLuminance, fullBlackLuminance);
        if (targetLuminance <= minimum) {
            black = descending ? 1D : 0D;
            return true;
        }
        if (targetLuminance >= maximum) {
            black = descending ? 0D : 1D;
            return true;
        }

        double lower = 0D;
        double upper = 1D;
        for (int iteration = 0; iteration < 20; iteration++) {
            double candidate = (lower + upper) * 0.5D;
            if (!TryGetBlackAxisLuminance(candidate, renderingIntent, out double luminance)) return false;
            if ((descending && luminance > targetLuminance) || (!descending && luminance < targetLuminance)) {
                lower = candidate;
            } else {
                upper = candidate;
            }
        }
        black = (lower + upper) * 0.5D;
        return true;
    }

    private bool TryGetBlackAxisLuminance(
        double black,
        OfficeIccRenderingIntent renderingIntent,
        out double luminance) =>
        TryGetDevicePcsLuminance(
            new DeviceComponentValues(4, 0D, 0D, 0D, Clamp01(black)),
            renderingIntent,
            out luminance);

    private bool TryGetDevicePcsLuminance(
        DeviceComponentValues components,
        OfficeIccRenderingIntent renderingIntent,
        out double luminance) {
        luminance = 0D;
        IDeviceToPcsTransform? transform = SelectDeviceToPcsTransform(renderingIntent);
        if (transform == null || !transform.TryTransform(components, _whitePoint, out XyzValue pcsXyz)) return false;
        pcsXyz = ApplyRenderingIntentToPcs(pcsXyz, renderingIntent);
        if (!IsFinite(pcsXyz.Y)) return false;
        luminance = pcsXyz.Y;
        return true;
    }
}
