namespace OfficeIMO.Pdf;

/// <summary>Options for append-only AcroForm field updates.</summary>
public sealed class PdfIncrementalFormFieldUpdateOptions {
    /// <summary>Cancellation observed during source admission, field traversal and append-only serialization.</summary>
    public System.Threading.CancellationToken CancellationToken { get; set; }

    /// <summary>Set AcroForm /NeedAppearances so viewers may regenerate appearances.</summary>
    public bool KeepNeedAppearances { get; set; } = true;

    /// <summary>Append simple widget normal appearance streams for updated text, choice, checkbox, and radio fields.</summary>
    public bool GenerateAppearanceStreams { get; set; }

    /// <summary>How a hybrid AcroForm/XFA form's XFA packet is treated. The default keeps it unchanged.</summary>
    public PdfFormXfaHandling XfaHandling { get; set; } = PdfFormXfaHandling.Preserve;
}
