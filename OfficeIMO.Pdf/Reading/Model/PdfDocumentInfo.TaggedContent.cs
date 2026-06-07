namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentInfo {
    /// <summary>Tagged PDF structure metadata discovered from /MarkInfo and /StructTreeRoot.</summary>
    public PdfTaggedContentInfo? TaggedContent { get; }

    /// <summary>True when readable tagged PDF structure metadata was discovered.</summary>
    public bool HasReadableTaggedContent => TaggedContent is not null;
}
