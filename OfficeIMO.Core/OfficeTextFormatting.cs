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
