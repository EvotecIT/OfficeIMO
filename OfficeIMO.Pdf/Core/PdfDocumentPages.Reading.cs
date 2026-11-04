using System.Collections;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentPages : IReadOnlyList<PdfReadPage> {
    private IReadOnlyList<PdfReadPage> ReadPages => _document.GetReadSnapshot().Document.Pages;
    /// <summary>Gets the document's page count.</summary>
    public int Count => ReadPages.Count;
    /// <summary>Gets a page by its zero-based index.</summary>
    public PdfReadPage this[int index] => ReadPages[index];
    /// <summary>Enumerates pages in document order.</summary>
    public IEnumerator<PdfReadPage> GetEnumerator() => ReadPages.GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
