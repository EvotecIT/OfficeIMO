namespace OfficeIMO.Drawing;

/// <summary>
/// Shared stroke dash style that OfficeIMO packages can map into their own drawing formats.
/// </summary>
public enum OfficeStrokeDashStyle {
    /// <summary>Continuous stroke.</summary>
    Solid,

    /// <summary>Dashed stroke.</summary>
    Dash,

    /// <summary>Dotted stroke.</summary>
    Dot,

    /// <summary>Alternating dash and dot stroke.</summary>
    DashDot,

    /// <summary>Alternating dash and two dot stroke.</summary>
    DashDotDot
}
