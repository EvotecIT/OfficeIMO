namespace OfficeIMO.Pdf;

/// <summary>Options for append-only AcroForm field updates.</summary>
public sealed class PdfIncrementalFormFieldUpdateOptions {
    /// <summary>Set AcroForm /NeedAppearances so viewers may regenerate appearances.</summary>
    public bool KeepNeedAppearances { get; set; } = true;

    /// <summary>Append simple widget normal appearance streams for updated text, choice, checkbox, and radio fields.</summary>
    public bool GenerateAppearanceStreams { get; set; }
}
