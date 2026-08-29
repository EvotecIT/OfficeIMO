using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Defines the line pattern used to decorate shared drawing text.
/// </summary>
public enum OfficeTextDecorationStyle {
    /// <summary>No decoration line.</summary>
    None,
    /// <summary>One solid decoration line.</summary>
    Single,
    /// <summary>Two parallel solid decoration lines.</summary>
    Double,
    /// <summary>A dotted decoration line.</summary>
    Dotted,
    /// <summary>A dashed decoration line.</summary>
    Dashed,
    /// <summary>A wavy decoration line.</summary>
    Wavy
}

/// <summary>
/// Defines vertical baseline placement for shared drawing text.
/// </summary>
public enum OfficeTextBaseline {
    /// <summary>Normal baseline and font size.</summary>
    Normal,
    /// <summary>Raised, reduced superscript text.</summary>
    Superscript,
    /// <summary>Lowered, reduced subscript text.</summary>
    Subscript
}

/// <summary>Resolved fixed-layout geometry for a cumulative superscript or subscript level.</summary>
public readonly struct OfficeTextScriptGeometry {
    internal OfficeTextScriptGeometry(double renderedFontSize, double baselineOffset) {
        RenderedFontSize = renderedFontSize;
        BaselineOffset = baselineOffset;
    }

    /// <summary>Gets the font size after every nested script reduction has been applied.</summary>
    public double RenderedFontSize { get; }

    /// <summary>Gets the cumulative top-down baseline offset; negative values raise text.</summary>
    public double BaselineOffset { get; }

    /// <summary>
    /// Resolves cumulative script geometry using the supplied destination-specific offset factors.
    /// </summary>
    public static OfficeTextScriptGeometry Resolve(
        double sourceFontSize,
        int baselineLevel,
        double superscriptOffsetFactor = 0.30D,
        double subscriptOffsetFactor = 0.15D) {
        if (double.IsNaN(sourceFontSize) || double.IsInfinity(sourceFontSize) || sourceFontSize <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(sourceFontSize));
        }
        if (baselineLevel < -32 || baselineLevel > 32) {
            throw new ArgumentOutOfRangeException(nameof(baselineLevel));
        }
        if (double.IsNaN(superscriptOffsetFactor) || double.IsInfinity(superscriptOffsetFactor) || superscriptOffsetFactor < 0D) {
            throw new ArgumentOutOfRangeException(nameof(superscriptOffsetFactor));
        }
        if (double.IsNaN(subscriptOffsetFactor) || double.IsInfinity(subscriptOffsetFactor) || subscriptOffsetFactor < 0D) {
            throw new ArgumentOutOfRangeException(nameof(subscriptOffsetFactor));
        }

        int magnitude = Math.Abs(baselineLevel);
        if (magnitude == 0) return new OfficeTextScriptGeometry(sourceFontSize, 0D);
        const double scriptScale = 0.65D;
        double renderedFontSize = sourceFontSize * Math.Pow(scriptScale, magnitude);
        double geometricSum = (1D - Math.Pow(scriptScale, magnitude)) / (1D - scriptScale);
        double offset = sourceFontSize * geometricSum * (baselineLevel > 0
            ? -superscriptOffsetFactor
            : subscriptOffsetFactor);
        return new OfficeTextScriptGeometry(renderedFontSize, offset);
    }
}
