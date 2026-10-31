namespace OfficeIMO.Pdf;

/// <summary>Controls native PDF output and destination replacement.</summary>
public sealed class PdfSaveOptions {
    /// <summary>Accepts fidelity loss introduced by native PDF serialization. The default blocks it.</summary>
    public bool AllowLoss { get; set; }
    /// <summary>Encrypts this output after native fidelity validation. Null preserves existing security.</summary>
    /// <remarks>The input must be unencrypted and must not retain PDF/A, PDF/X or electronic-invoice identification. Use Security.Reencrypt with the current owner password to change existing security. This option does not change the live document's security configuration.</remarks>
    public PdfStandardEncryptionOptions? Encryption { get; set; }
    /// <summary>Controls whether a file destination may replace an existing file.</summary>
    public OfficeConversionFileConflictPolicy ConflictPolicy { get; set; } = OfficeConversionFileConflictPolicy.Replace;
}
