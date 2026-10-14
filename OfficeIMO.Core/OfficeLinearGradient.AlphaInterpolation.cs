namespace OfficeIMO.Drawing;

public sealed partial class OfficeLinearGradient {
    // SVG/XPS paint interpolates unassociated color channels and alpha. Keep this
    // imported intent distinct from the canvas default used by CSS backgrounds.
    internal bool InterpolateAlphaSeparately { get; private set; }

    internal OfficeLinearGradient WithSeparateAlphaInterpolation(bool enabled = true) {
        var copy = Clone();
        copy.InterpolateAlphaSeparately = enabled;
        return copy;
    }
}
