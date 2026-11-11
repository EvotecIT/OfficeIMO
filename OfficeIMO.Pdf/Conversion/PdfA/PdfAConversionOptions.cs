namespace OfficeIMO.Pdf;

/// <summary>Controls conversion of an existing PDF into a PDF/A archival file.</summary>
public sealed class PdfAConversionOptions {
    /// <summary>
    /// Gets or sets the target profile. <see cref="PdfComplianceProfile.PdfA2B"/>, <see cref="PdfComplianceProfile.PdfA2U"/>,
    /// <see cref="PdfComplianceProfile.PdfA3B"/> and <see cref="PdfComplianceProfile.PdfA3U"/> are supported.
    /// </summary>
    public PdfComplianceProfile Profile { get; set; } = PdfComplianceProfile.PdfA2B;

    /// <summary>
    /// Gets or sets the owner password of an encrypted source. PDF/A forbids encryption, so the converted file is written
    /// without it; removal requires owner authorization.
    /// </summary>
    public string? Password { get; set; }

    /// <summary>
    /// Gets or sets whether fonts the source does not embed may be embedded from installed system fonts. A substitute is
    /// used only when every glyph the pages draw exists in it with the advance width the source already positions it by.
    /// Defaults to <see langword="true"/>.
    /// </summary>
    public bool EmbedSubstituteFonts { get; set; } = true;

    /// <summary>
    /// Gets or sets whether embedded files may be removed when the target is PDF/A-2, which cannot carry arbitrary
    /// attachments. When <see langword="false"/>, a source with attachments is refused. Defaults to <see langword="true"/>.
    /// </summary>
    public bool RemoveEmbeddedFiles { get; set; } = true;

    /// <summary>
    /// Gets or sets an ICC profile for the output intent. When <see langword="null"/>, an existing PDF/A output intent is kept,
    /// or the built-in sRGB IEC61966-2.1 profile is added. Supply a CMYK profile to convert files painted in device CMYK.
    /// </summary>
    public byte[]? OutputIntentProfile { get; set; }

    /// <summary>Gets or sets the output condition identifier written with <see cref="OutputIntentProfile"/>.</summary>
    public string? OutputConditionIdentifier { get; set; }
}
