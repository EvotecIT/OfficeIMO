namespace OfficeIMO.Pdf;

/// <summary>Existing-document bookmark editing operations.</summary>
public sealed class PdfDocumentBookmarks {
    private readonly PdfDocument _document;
    internal PdfDocumentBookmarks(PdfDocument document) { _document = document; }
    /// <summary>Applies a transactional bookmark edit.</summary>
    public PdfBookmarkEditResult Edit(Action<PdfBookmarkEditSession> edit) => PdfBookmarkEditor.Edit(_document.GetBytesForOperation(), edit, _document.ReadOptions);
    /// <summary>Reports unresolved bookmark destinations.</summary>
    public IReadOnlyList<PdfBookmarkValidationIssue> Validate() => PdfBookmarkEditor.Validate(_document.GetBytesForOperation(), _document.ReadOptions);
}
