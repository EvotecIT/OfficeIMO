namespace OfficeIMO.Studio.Features.Workspace;

internal enum PdfWorkspaceOperationKind {
    Reorder,
    Rotate,
    Delete,
    Duplicate,
    Import,
    Crop,
    InsertBlank,
    Annotation,
    TextEdit,
    ImageEdit,
    AddedContent,
    FormFill,
    FormAuthor,
    FormFlatten,
    Security,
    Signature,
    BatesNumbering,
    Redaction,
    Watermark,
    PageNumbers,
    RecoveryRestore,
    Undo,
    Redo
}

internal sealed record PdfWorkspaceOperation(
    long Revision,
    PdfWorkspaceOperationKind Kind,
    string Description,
    IReadOnlyList<int> PageNumbers,
    DateTimeOffset Timestamp);

internal sealed record PdfWorkspaceProgress(string Stage, double Fraction);
