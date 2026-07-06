namespace OfficeIMO.Pdf;

/// <summary>
/// Border style used by generated AcroForm field dictionaries and appearance streams.
/// </summary>
public enum PdfFormFieldBorderStyle {
    /// <summary>Solid rectangular or circular field border.</summary>
    Solid = 0,

    /// <summary>Dashed field border. Uses <see cref="PdfFormFieldStyle.BorderDashPattern"/> when provided.</summary>
    Dashed = 1,

    /// <summary>Underline border drawn along the lower edge of the field rectangle.</summary>
    Underline = 2,

    /// <summary>Beveled border with highlighted top and left edges.</summary>
    Beveled = 3,

    /// <summary>Inset border with shadowed top and left edges.</summary>
    Inset = 4
}
