namespace OfficeIMO.Pdf;

/// <summary>Typed permissions stored in a PDF Standard security `/P` mask.</summary>
[Flags]
public enum PdfStandardPermissions {
    /// <summary>No optional document operations are allowed.</summary>
    None = 0,

    /// <summary>Print at the quality allowed by the encryption revision.</summary>
    Print = 1 << 0,

    /// <summary>Modify document contents.</summary>
    ModifyContents = 1 << 1,

    /// <summary>Copy or otherwise extract text and graphics.</summary>
    CopyContents = 1 << 2,

    /// <summary>Add or modify annotations and interactive form fields.</summary>
    ModifyAnnotations = 1 << 3,

    /// <summary>Fill existing interactive form fields.</summary>
    FillForms = 1 << 4,

    /// <summary>Extract content for accessibility purposes.</summary>
    Accessibility = 1 << 5,

    /// <summary>Assemble the document by inserting, rotating, or deleting pages.</summary>
    AssembleDocument = 1 << 6,

    /// <summary>Print a faithful high-quality representation.</summary>
    HighQualityPrint = 1 << 7,

    /// <summary>Allow every typed Standard security operation.</summary>
    All = Print | ModifyContents | CopyContents | ModifyAnnotations | FillForms | Accessibility | AssembleDocument | HighQualityPrint
}
