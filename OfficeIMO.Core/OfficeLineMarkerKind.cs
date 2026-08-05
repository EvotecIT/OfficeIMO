namespace OfficeIMO.Drawing;

/// <summary>
/// Describes reusable line marker shapes for open Drawing lines.
/// </summary>
public enum OfficeLineMarkerKind {
    /// <summary>No marker should be rendered.</summary>
    None,

    /// <summary>A filled triangular arrowhead.</summary>
    Triangle,

    /// <summary>A filled arrowhead with an inset base.</summary>
    Stealth,

    /// <summary>A filled diamond marker.</summary>
    Diamond,

    /// <summary>A filled oval marker.</summary>
    Oval,

    /// <summary>A generic filled arrow marker.</summary>
    Arrow
}
