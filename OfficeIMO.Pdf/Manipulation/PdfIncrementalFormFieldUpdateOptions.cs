namespace OfficeIMO.Pdf;

/// <summary>Options for append-only AcroForm field updates.</summary>
public sealed class PdfIncrementalFormFieldUpdateOptions : IPdfOperationCancellation {
    // Operation cancellation; public methods take it as a parameter and bind it to a per-call copy.
    System.Threading.CancellationToken IPdfOperationCancellation.CancellationToken { get => CancellationToken; set => CancellationToken = value; }

    internal System.Threading.CancellationToken CancellationToken { get; set; }

    object IPdfOperationCancellation.CloneForOperation() => MemberwiseClone();

    /// <summary>Set AcroForm /NeedAppearances so viewers may regenerate appearances.</summary>
    public bool KeepNeedAppearances { get; set; } = true;

    /// <summary>Append simple widget normal appearance streams for updated text, choice, checkbox, and radio fields.</summary>
    public bool GenerateAppearanceStreams { get; set; }

    /// <summary>How a hybrid AcroForm/XFA form's XFA packet is treated. The default keeps it unchanged.</summary>
    public PdfFormXfaHandling XfaHandling { get; set; } = PdfFormXfaHandling.Preserve;
}
