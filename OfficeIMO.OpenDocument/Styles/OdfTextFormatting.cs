namespace OfficeIMO.OpenDocument;

/// <summary>Native ODF text-decoration line styles.</summary>
public enum OdfTextDecorationStyle {
    /// <summary>No decoration line.</summary>
    None,
    /// <summary>Solid line.</summary>
    Solid,
    /// <summary>Dotted line.</summary>
    Dotted,
    /// <summary>Dashed line.</summary>
    Dash,
    /// <summary>Long-dashed line.</summary>
    LongDash,
    /// <summary>Alternating dot and dash line.</summary>
    DotDash,
    /// <summary>Alternating two dots and a dash line.</summary>
    DotDotDash,
    /// <summary>Wavy line.</summary>
    Wave
}

/// <summary>Native ODF text-decoration line counts.</summary>
public enum OdfTextDecorationType {
    /// <summary>No decoration.</summary>
    None,
    /// <summary>Single decoration line.</summary>
    Single,
    /// <summary>Double decoration line.</summary>
    Double
}

/// <summary>Native ODF text baseline placement.</summary>
public enum OdfTextPosition {
    /// <summary>Normal baseline.</summary>
    Normal,
    /// <summary>Superscript baseline.</summary>
    Superscript,
    /// <summary>Subscript baseline.</summary>
    Subscript
}

/// <summary>Native ODF display-time text transformations.</summary>
public enum OdfTextTransform {
    /// <summary>No display-time transformation.</summary>
    None,
    /// <summary>Uppercase transformation.</summary>
    Uppercase,
    /// <summary>Lowercase transformation.</summary>
    Lowercase,
    /// <summary>Capitalize words.</summary>
    Capitalize
}
