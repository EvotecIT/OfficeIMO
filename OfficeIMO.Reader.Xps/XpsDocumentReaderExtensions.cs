using OfficeIMO.Xps;

namespace OfficeIMO.Reader.Xps;

/// <summary>Projects an already loaded native document into the shared Reader envelope.</summary>
public static class XpsDocumentReaderExtensions {
    /// <summary>Reads literal Unicode, native semantics and page identity without saving and reopening the source.</summary>
    /// <remarks>Package limits apply when loading; already loaded documents use their existing native limits. No package source hash is inferred for a mutable loaded document.</remarks>
    public static OfficeDocumentReadResult ToOfficeDocumentReadResult(this XpsDocument document, string? sourceName = null,
        ReaderOptions? readerOptions = null, ReaderXpsOptions? xpsOptions = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        var options = DocumentReaderEngine.NormalizeOptions(readerOptions);
        using var scope = ReaderReadScope.Enter(options);
        string name = string.IsNullOrWhiteSpace(sourceName) ? document.Format == XpsFormat.OpenXps ? "document.oxps" : "document.xps" : sourceName!;
        return ReaderReadScope.Complete(XpsReaderAdapter.Project(document,
            new OfficeDocumentSource { Path = name, SourceId = DocumentReaderEngine.BuildPortableSourceId(name) }, options,
            (xpsOptions ?? new ReaderXpsOptions()).Clone(), cancellationToken));
    }
}
