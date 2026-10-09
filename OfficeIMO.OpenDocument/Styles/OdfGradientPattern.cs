namespace OfficeIMO.OpenDocument;

/// <summary>The six native two-color OpenDocument gradient styles.</summary>
public enum OdfGradientStyle {
    /// <summary>Interpolation along an axis, clockwise from vertical.</summary>
    Linear,
    /// <summary>Mirrored interpolation, with the start color at the axis center.</summary>
    Axial,
    /// <summary>Circular interpolation, with the end color at the center.</summary>
    Radial,
    /// <summary>Elliptical interpolation.</summary>
    Ellipsoid,
    /// <summary>Concentric square interpolation.</summary>
    Square,
    /// <summary>Concentric rectangular interpolation.</summary>
    Rectangular
}

/// <summary>Immutable native gradient colors, angle, border, intensities and center.</summary>
/// <remarks>Fractions use one for 100%. The model retains all six styles; drawing projection has a narrower, reported profile.</remarks>
public sealed class OdfGradientPattern {
    /// <summary>Creates a two-color gradient. Centers may lie outside the filled area.</summary>
    public OdfGradientPattern(OdfGradientStyle style, OdfColor startColor, OdfColor endColor,
        double angleDegrees = 0, double border = 0, double startIntensity = 1, double endIntensity = 1,
        double centerX = 0.5, double centerY = 0.5) {
        if (!Enum.IsDefined(typeof(OdfGradientStyle), style)) throw new ArgumentOutOfRangeException(nameof(style));
        ValidateFinite(angleDegrees, nameof(angleDegrees)); ValidateFraction(border, nameof(border));
        ValidateFraction(startIntensity, nameof(startIntensity)); ValidateFraction(endIntensity, nameof(endIntensity));
        ValidateCenter(centerX, nameof(centerX)); ValidateCenter(centerY, nameof(centerY));
        Style = style; StartColor = startColor; EndColor = endColor; AngleDegrees = angleDegrees; Border = border;
        StartIntensity = startIntensity; EndIntensity = endIntensity; CenterX = centerX; CenterY = centerY;
    }
    /// <summary>Native interpolation style.</summary>
    public OdfGradientStyle Style { get; }
    /// <summary>Declared start color before intensity adjustment.</summary>
    public OdfColor StartColor { get; }
    /// <summary>Declared end color before intensity adjustment.</summary>
    public OdfColor EndColor { get; }
    /// <summary>Clockwise angle from vertical in degrees, ignored for radial gradients.</summary>
    public double AngleDegrees { get; }
    /// <summary>Solid border fraction from zero to one.</summary>
    public double Border { get; }
    /// <summary>Start color intensity from zero to one.</summary>
    public double StartIntensity { get; }
    /// <summary>End color intensity from zero to one.</summary>
    public double EndIntensity { get; }
    /// <summary>Horizontal center as a fraction of the geometry bounds.</summary>
    public double CenterX { get; }
    /// <summary>Vertical center as a fraction of the geometry bounds.</summary>
    public double CenterY { get; }
    private static void ValidateFraction(double value, string parameter) {
        ValidateFinite(value, parameter);
        if (value < 0 || value > 1) throw new ArgumentOutOfRangeException(parameter, "Gradient fractions must be between zero and one.");
    }
    private static void ValidateFinite(double value, string parameter) {
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(parameter, "Gradient values must be finite.");
    }
    private static void ValidateCenter(double value, string parameter) {
        ValidateFinite(value, parameter);
        if (double.IsInfinity(value * 100)) throw new ArgumentOutOfRangeException(parameter, "Gradient centers require a finite percentage representation.");
    }
}
