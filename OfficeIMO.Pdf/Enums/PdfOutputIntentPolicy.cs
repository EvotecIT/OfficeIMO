namespace OfficeIMO.Pdf;

/// <summary>
/// Declares the intended policy for a generated PDF output intent.
/// </summary>
public enum PdfOutputIntentPolicy {
    /// <summary>No profile-specific output-intent policy has been declared.</summary>
    Unspecified,
    /// <summary>The output intent is intended to represent sRGB IEC61966-2.1 output.</summary>
    SrgbIec6196621
}
