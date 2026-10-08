using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Shared font, layout and effect settings for raster text measurement and drawing.</summary>
/// <remarks>Operations snapshot scalar settings and the font collection. Font and shaping providers, and the optional
/// diagnostic sink, are retained references. Drawing mutates pixels and cancellation does not undo pixels already painted.</remarks>
public sealed class OfficeRasterTextOptions {
    /// <summary>Positive finite text size in pixels.</summary>
    public double FontSize { get; set; } = 16D;
    /// <summary>Requested managed font family or fallback list.</summary>
    public string? FontFamily { get; set; }
    /// <summary>Existing bold, italic, underline and strikethrough style flags.</summary>
    public OfficeFontStyle FontStyle { get; set; }
    /// <summary>Explicit outline font; null uses the canvas default font.</summary>
    public OfficeTrueTypeFont? Font { get; set; }
    /// <summary>Scoped managed font faces used consistently for measurement and drawing.</summary>
    public OfficeFontFaceCollection? Fonts { get; set; }
    /// <summary>Optional shaping provider supported by the shared canvas.</summary>
    public IOfficeTextShapingProvider? TextShapingProvider { get; set; }
    /// <summary>Optional language passed to the shaping provider.</summary>
    public string? TextShapingLanguage { get; set; }
    /// <summary>Optional collection receiving canvas fidelity diagnostics.</summary>
    public ICollection<OfficeImageExportDiagnostic>? DiagnosticSink { get; set; }
    /// <summary>Optional source identifier attached to diagnostics.</summary>
    public string? DiagnosticSource { get; set; }
    /// <summary>Horizontal alignment in the drawing rectangle.</summary>
    public OfficeTextAlignment HorizontalAlignment { get; set; }
    /// <summary>Vertical alignment in the drawing rectangle.</summary>
    public OfficeTextVerticalAlignment VerticalAlignment { get; set; }
    /// <summary>Positive finite multiplier of font size used for line height.</summary>
    public double LineHeight { get; set; } = 1.2D;
    /// <summary>Wraps to the drawing width or the explicitly supplied measurement width.</summary>
    public bool Wrap { get; set; }
    /// <summary>Clips glyphs, outline and shadow to the drawing rectangle.</summary>
    public bool Clip { get; set; }
    /// <summary>Optional shadow color.</summary>
    public OfficeColor? ShadowColor { get; set; }
    /// <summary>Finite horizontal shadow offset in pixels, rounded to the nearest pixel.</summary>
    public double ShadowOffsetX { get; set; }
    /// <summary>Finite vertical shadow offset in pixels, rounded to the nearest pixel.</summary>
    public double ShadowOffsetY { get; set; }
    /// <summary>Optional outline color.</summary>
    public OfficeColor? OutlineColor { get; set; }
    /// <summary>Rounded coverage-expansion radius from zero through sixty-four pixels.</summary>
    public double OutlineWidth { get; set; }

    /// <summary>Returns an independent copy of the settings and scoped font collection for a composed operation.</summary>
    /// <remarks>Capture once before planning multiple frames, then use that copy for planning and drawing.
    /// Font and shaping providers and the diagnostic sink remain retained references, as for individual text operations.</remarks>
    public OfficeRasterTextOptions Clone() {
        var result = (OfficeRasterTextOptions)MemberwiseClone();
        result.Fonts = Fonts?.Clone();
        return result;
    }

    internal void Validate() {
        if (!Finite(FontSize) || FontSize <= 0D) throw new ArgumentOutOfRangeException(nameof(FontSize));
        if (!Finite(LineHeight) || LineHeight <= 0D || !Finite(FontSize * LineHeight)) throw new ArgumentOutOfRangeException(nameof(LineHeight));
        if (!Finite(OutlineWidth) || OutlineWidth < 0D || OutlineWidth > 64D) throw new ArgumentOutOfRangeException(nameof(OutlineWidth));
        if (!Finite(ShadowOffsetX) || !Finite(ShadowOffsetY)) throw new ArgumentOutOfRangeException(nameof(ShadowOffsetX));
        if ((FontStyle & ~(OfficeFontStyle.Bold | OfficeFontStyle.Italic | OfficeFontStyle.Underline | OfficeFontStyle.Strikethrough)) != 0) throw new ArgumentOutOfRangeException(nameof(FontStyle));
        if (HorizontalAlignment < OfficeTextAlignment.Left || HorizontalAlignment > OfficeTextAlignment.Justify) throw new ArgumentOutOfRangeException(nameof(HorizontalAlignment));
        if (VerticalAlignment < OfficeTextVerticalAlignment.Top || VerticalAlignment > OfficeTextVerticalAlignment.Bottom) throw new ArgumentOutOfRangeException(nameof(VerticalAlignment));
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
