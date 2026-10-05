using System.Linq;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRadialGradient {
    // Used by owned native-format adapters; normal SVG uses its original field.
    private bool _paintOutsideCone;
    internal OfficeColor? OutsideColor => _paintOutsideCone ? Stops[0].Color : (OfficeColor?)null;

    internal OfficeRadialGradient WithFirstPadIntersection() {
        double sx = (StartX - EndX) / EndRadiusX;
        double sy = (StartY - EndY) / EndRadiusY;
        if (sx * sx + sy * sy < 1D) return this;
        // Reverse the circle sequence and stops. The shared/PDF maximum physical
        // root then selects the original smallest containing ellipse. Express the
        // ellipse as an affine circle so a zero-radius end remains well-defined.
        var reverse = new OfficeRadialGradient(0D, 0D, 1D, sx, sy, 0D,
            Stops.Reverse().Select(stop => new OfficeGradientStop(1D - stop.Offset, stop.Color)).ToArray())
            .TransformCoordinates(new OfficeTransform(EndRadiusX, 0D, 0D, EndRadiusY, EndX, EndY).Then(CoordinateTransform));
        reverse.ColorInterpolation = ColorInterpolation;
        reverse._paintOutsideCone = true;
        return reverse;
    }
}
