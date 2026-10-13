namespace OfficeIMO.Reader;

/// <summary>
/// Controls optional native inspection metadata. Limits span root, nested, folder and batch reads;
/// exceeding them terminates the operation rather than returning silently incomplete evidence.
/// Owning format parser limits remain in force and formulas, macros and external targets remain inert.
/// </summary>
public sealed class ReaderInspectionOptions {
    /// <summary>Include review comments and their authors, replies and resolution state when supported.</summary>
    public bool IncludeComments { get; set; } = true;
    /// <summary>Include tracked revision evidence without accepting or rejecting changes.</summary>
    public bool IncludeRevisions { get; set; } = true;
    /// <summary>Include authored formula expressions and available saved cached values without calculation.</summary>
    public bool IncludeFormulas { get; set; } = true;
    /// <summary>Include authored hidden sheet, row, column, slide and shape state when supported.</summary>
    public bool IncludeVisibility { get; set; } = true;
    /// <summary>Include embedded-payload and VBA metadata without extracting or executing payloads.</summary>
    public bool IncludePackageContent { get; set; } = true;
    /// <summary>Include bounded owner-provided concealment and Unicode evidence. This is heuristic inspection, not a safety guarantee.</summary>
    public bool IncludeContentSafety { get; set; } = true;
    /// <summary>Maximum inspection records across the operation, including coverage entries. Default: 4,096.</summary>
    public long MaxRecords { get; set; } = 4096;
    /// <summary>Maximum retained inspection metadata characters, including names, values and attributes. Default: one million.</summary>
    public long MaxCharacters { get; set; } = 1_000_000;

    internal ReaderInspectionOptions CloneValidated() {
        if (MaxRecords < 1) throw new ArgumentOutOfRangeException(nameof(MaxRecords));
        if (MaxCharacters < 1) throw new ArgumentOutOfRangeException(nameof(MaxCharacters));
        return (ReaderInspectionOptions)MemberwiseClone();
    }
}

/// <summary>Stable metadata categories for native inspection evidence in the existing Reader JSON and structured contracts.</summary>
public static class OfficeDocumentInspectionCategories {
    /// <summary>Inspection coverage, native scope and explicit completeness boundaries.</summary>
    public const string Coverage = "inspection.coverage";
    /// <summary>Review comments, including replies.</summary>
    public const string Comment = "inspection.comment";
    /// <summary>Tracked revision metadata and affected text.</summary>
    public const string Revision = "inspection.revision";
    /// <summary>Authored formulas and available saved cached values.</summary>
    public const string Formula = "inspection.formula";
    /// <summary>Authored hidden container or object state.</summary>
    public const string Visibility = "inspection.visibility";
    /// <summary>Inert embedded package and OLE payload descriptors.</summary>
    public const string EmbeddedPayload = "inspection.embedded-payload";
    /// <summary>Inert VBA project descriptors.</summary>
    public const string VbaProject = "inspection.vba-project";
    /// <summary>Structural observations from the exact loader source bytes.</summary>
    public const string PackageSecurity = "inspection.package-security";
    /// <summary>Bounded concealment and instruction-like content observations.</summary>
    public const string ContentSafety = "inspection.content-safety";
    /// <summary>Exact Unicode text-integrity evidence.</summary>
    public const string TextIntegrity = "inspection.text-integrity";
}
