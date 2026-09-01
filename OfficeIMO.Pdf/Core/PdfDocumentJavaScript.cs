namespace OfficeIMO.Pdf;

/// <summary>Explicit active-content operations for named document-level JavaScript.</summary>
public sealed class PdfDocumentJavaScript {
    private readonly PdfDocument _document;

    internal PdfDocumentJavaScript(PdfDocument document) { _document = document; }

    /// <summary>Lists named document-level scripts and their exact source text.</summary>
    public IReadOnlyList<PdfJavaScript> List() => _document.Reader.JavaScripts();

    /// <summary>
    /// Applies a transactional named-script collection edit. JavaScript is active content and may execute in capable PDF viewers.
    /// The default sanitization policy removes authored scripts unless JavaScript is explicitly allowed.
    /// </summary>
    public PdfJavaScriptEditResult Edit(Action<PdfJavaScriptEditSession> edit) =>
        PdfJavaScriptEditor.Edit(_document.GetBytesForOperation(), edit, _document.ReadOptions);

    /// <summary>Adds a named script or replaces the source of an existing script with the same name.</summary>
    public PdfJavaScriptEditResult AddOrReplace(string name, string script) =>
        PdfJavaScriptEditor.AddOrReplace(_document.GetBytesForOperation(), name, script, _document.ReadOptions);

    /// <summary>Removes a named script when it exists.</summary>
    public PdfJavaScriptEditResult Remove(string name) =>
        PdfJavaScriptEditor.Remove(_document.GetBytesForOperation(), name, _document.ReadOptions);

    /// <summary>Removes every named document-level script.</summary>
    public PdfJavaScriptEditResult Clear() =>
        PdfJavaScriptEditor.Clear(_document.GetBytesForOperation(), _document.ReadOptions);
}
