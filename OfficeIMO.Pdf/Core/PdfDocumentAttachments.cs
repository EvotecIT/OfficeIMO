namespace OfficeIMO.Pdf;

/// <summary>Existing-document embedded and associated file editing operations.</summary>
public sealed class PdfDocumentAttachments {
    private readonly PdfDocument _document;
    internal PdfDocumentAttachments(PdfDocument document) { _document = document; }
    /// <summary>Applies a transactional attachment collection edit.</summary>
    public PdfAttachmentEditResult Edit(Action<PdfAttachmentEditSession> edit) => PdfAttachmentEditor.Edit(_document.GetBytesForOperation(), edit, _document.ReadOptions);
    /// <summary>Adds one attachment.</summary>
    public PdfAttachmentEditResult Add(PdfEmbeddedFile attachment) => PdfAttachmentEditor.Add(_document.GetBytesForOperation(), attachment, _document.ReadOptions);
    /// <summary>Replaces one attachment by file name.</summary>
    public PdfAttachmentEditResult Replace(string fileName, PdfEmbeddedFile replacement) => PdfAttachmentEditor.Replace(_document.GetBytesForOperation(), fileName, replacement, _document.ReadOptions);
    /// <summary>Renames one attachment.</summary>
    public PdfAttachmentEditResult Rename(string fileName, string newFileName) => PdfAttachmentEditor.Rename(_document.GetBytesForOperation(), fileName, newFileName, _document.ReadOptions);
    /// <summary>Removes one attachment.</summary>
    public PdfAttachmentEditResult Remove(string fileName) => PdfAttachmentEditor.Remove(_document.GetBytesForOperation(), fileName, _document.ReadOptions);
}
