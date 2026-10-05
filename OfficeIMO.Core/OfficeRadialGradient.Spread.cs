using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRadialGradient {
    /// <summary>Behavior beyond the stop interval. Defaults to endpoint padding.</summary>
    public OfficeGradientSpreadMode SpreadMode { get; private set; }

    /// <summary>Returns a detached gradient with the requested spread behavior.</summary>
    public OfficeRadialGradient WithSpreadMode(OfficeGradientSpreadMode spreadMode) {
        if (spreadMode < OfficeGradientSpreadMode.Pad || spreadMode > OfficeGradientSpreadMode.Reflect)
            throw new ArgumentOutOfRangeException(nameof(spreadMode));
        var copy = Clone(); copy.SpreadMode = spreadMode; return copy;
    }

    private bool _paintSpreadAverage;
    private OfficeColor _spreadAverage;

    // SVG's point-focus tangent compatibility rule differs from native XPS
    // endpoint paint and ordinary two-circle transparent-cone behavior.
    internal OfficeRadialGradient WithSvgBoundarySpreadAverage() {
        var copy = Clone(); copy._paintSpreadAverage = true; copy.RefreshSpreadAverage(); return copy;
    }

    private void RefreshSpreadAverage() {
        if (_paintSpreadAverage) _spreadAverage = OfficeGradientColors.Average(Stops, ColorInterpolation);
    }

    internal double SampleRatio(double x, double y) {
        double ratio = SampleUnboundedRatio(x, y);
        if (double.IsNaN(ratio)) return _paintOutsideCone ? OutsideStopOffset : double.NaN;
        if (_paintSpreadAverage && SpreadMode != OfficeGradientSpreadMode.Pad && double.IsInfinity(ratio)) return double.NaN;
        if (SpreadMode == OfficeGradientSpreadMode.Pad || double.IsInfinity(ratio))
            return Math.Max(0D, Math.Min(1D, ratio));
        if (SpreadMode == OfficeGradientSpreadMode.Repeat) {
            double mapped = ratio - Math.Floor(ratio);
            // Native fields reverse the circle sequence: select the original
            // start color at an exact seam, independently of outside-cone paint.
            return _paintOutsideCone && mapped == 0D ? 1D : mapped;
        }
        double reflected = ratio % 2D;
        if (reflected < 0D) reflected += 2D;
        return reflected > 1D ? 2D - reflected : reflected;
    }
}
