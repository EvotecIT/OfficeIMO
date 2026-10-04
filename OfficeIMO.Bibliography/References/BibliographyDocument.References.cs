namespace OfficeIMO.Bibliography;

public sealed partial class BibliographyDocument {
    /// <summary>
    /// Resolves local BibTeX/BibLaTeX crossref and whole-entry xdata references into an independent editable snapshot.
    /// Never changes this document, reads files, or fetches metadata. Ambiguous or missing keys are diagnosed rather than guessed.
    /// </summary>
    public BibliographyReferenceResult ResolveReferences(BibliographyReferenceOptions? options = null, CancellationToken cancellationToken = default) =>
        BibliographyReferenceResolver.Resolve(this, (options ?? new BibliographyReferenceOptions()).Snapshot(), cancellationToken);

    internal BibliographyDocument(BibliographyDocument source, IList<BibliographyItem> items, IList<BibliographyNativeEntry> entries,
        CancellationToken cancellationToken) {
        SourceFormat = source.SourceFormat;
        Items = items;
        NativeEntries = entries;
        Diagnostics = Array.AsReadOnly(source.Diagnostics.ToArray());
        CslJsonSingleObjectRoot = source.CslJsonSingleObjectRoot;
        EndNoteRecordsRoot = source.EndNoteRecordsRoot;
        EndNoteRootElementName = source.EndNoteRootElementName;
        EndNoteRecordsElementName = source.EndNoteRecordsElementName;
        _baselineFingerprint = BibliographyFingerprint.Create(this, cancellationToken);
    }
}
