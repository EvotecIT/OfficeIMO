namespace OfficeIMO.Pdf;

/// <summary>How an append-only form-field revision treats the XFA packet of a hybrid AcroForm/XFA form.</summary>
public enum PdfFormXfaHandling {
    /// <summary>
    /// Keep the XFA packet unchanged. AcroForm viewers show the new values; XFA-aware viewers such as Adobe
    /// Acrobat may keep showing the XFA data, which the revision does not update.
    /// </summary>
    Preserve = 0,

    /// <summary>
    /// Remove /XFA from the AcroForm in the appended revision so every viewer uses the AcroForm values. Earlier
    /// revisions keep the XFA packet; Adobe Reader treats usage rights as no longer applying to the changed form.
    /// </summary>
    Remove = 1
}
