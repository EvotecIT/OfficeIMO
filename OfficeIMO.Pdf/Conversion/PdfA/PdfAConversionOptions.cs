namespace OfficeIMO.Pdf;

/// <summary>Controls conversion of an existing PDF into a PDF/A archival file.</summary>
/// <remarks>
/// An encrypted source is converted when it was loaded with its owner password in <see cref="PdfLoadOptions.Password"/>:
/// PDF/A forbids encryption, so the converted file is written without it.
/// </remarks>
public sealed class PdfAConversionOptions {
    private PdfAFontSubstitution _fontSubstitution = PdfAFontSubstitution.InstalledFonts;

    /// <summary>
    /// Gets or sets the target profile: <see cref="PdfComplianceProfile.PdfA2B"/> (the default), <see cref="PdfComplianceProfile.PdfA2U"/>,
    /// <see cref="PdfComplianceProfile.PdfA3B"/>, <see cref="PdfComplianceProfile.PdfA3U"/>, <see cref="PdfComplianceProfile.PdfA4"/>
    /// or <see cref="PdfComplianceProfile.PdfA4F"/>. PDF/A-4 files are written as PDF 2.0; PDF/A-4f requires at least one
    /// embedded file in the source. Other profiles are rejected with <see cref="ArgumentException"/>.
    /// </summary>
    public PdfComplianceProfile Profile { get; set; } = PdfComplianceProfile.PdfA2B;

    /// <summary>
    /// Gets or sets the output intent to add. When <see langword="null"/> (the default), an existing PDF/A output intent is
    /// kept, or the built-in sRGB IEC61966-2.1 output intent is added. Supply one with a CMYK profile to convert files painted
    /// in device CMYK. A source that already has an output intent profile is refused, because PDF/A requires every output
    /// intent to share one profile.
    /// </summary>
    /// <remarks>
    /// The intent is written as a <c>GTS_PDFA1</c> output intent with its profile, output condition identifier, output
    /// condition, registry name and info. Its <see cref="PdfOutputIntent.Subtype"/> must be
    /// <see cref="PdfOutputIntentSubtype.GtsPdfA1"/>; <see cref="PdfOutputIntent.Policy"/> declares readiness for generated
    /// documents and is not used by conversion.
    /// </remarks>
    public PdfOutputIntent? OutputIntent { get; set; }

    /// <summary>
    /// Gets or sets what happens to embedded files when the target is PDF/A-2 or PDF/A-4, which cannot carry arbitrary
    /// attachments. PDF/A-3 and PDF/A-4f keep attachments. Defaults to <see cref="PdfAAttachmentPolicy.Remove"/>.
    /// </summary>
    public PdfAAttachmentPolicy Attachments { get; set; } = PdfAAttachmentPolicy.Remove;

    /// <summary>
    /// Gets or sets where fonts the source does not embed may be found. Defaults to
    /// <see cref="PdfAFontSubstitution.InstalledFonts"/>.
    /// </summary>
    public PdfAFontSubstitution FontSubstitution {
        get => _fontSubstitution;
        set => _fontSubstitution = value ?? throw new ArgumentNullException(nameof(value));
    }
}

/// <summary>Chooses what PDF/A conversion does with embedded files that the target profile cannot carry.</summary>
public enum PdfAAttachmentPolicy {
    /// <summary>Remove the embedded files and report each removal as an omission.</summary>
    Remove,
    /// <summary>Refuse the source, with a <see cref="PdfAConversionRefusalReason.EmbeddedFileNotAllowed"/> violation per file.</summary>
    Refuse
}
