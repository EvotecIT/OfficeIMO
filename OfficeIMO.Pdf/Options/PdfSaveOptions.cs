namespace OfficeIMO.Pdf;

/// <summary>Controls native PDF output and destination replacement.</summary>
public sealed class PdfSaveOptions {
    /// <summary>Accepts fidelity loss introduced by native PDF serialization. The default blocks it.</summary>
    public bool AllowLoss { get; set; }
    /// <summary>Controls whether a file destination may replace an existing file.</summary>
    public OfficeConversionFileConflictPolicy ConflictPolicy { get; set; } = OfficeConversionFileConflictPolicy.Replace;
}
