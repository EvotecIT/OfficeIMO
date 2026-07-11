namespace OfficeIMO.Pdf;

/// <summary>Controls whether known structural PDF defects are rejected or explicitly recovered.</summary>
public enum PdfParsingMode {
    /// <summary>Recover supported defects and expose every recovery through <see cref="PdfRepairReport"/>.</summary>
    Lenient = 0,

    /// <summary>Reject structural defects instead of applying recovery behavior.</summary>
    Strict = 1
}
