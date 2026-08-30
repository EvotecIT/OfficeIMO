using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>Read-only Pages structure recovered from a shared IWA object graph.</summary>
public sealed class IWorkPagesProjection {
    private readonly IWorkSourceDocument _source;
    private readonly bool _supportsEditableReconstruction;

    internal IWorkPagesProjection(IWorkSourceDocument source, IReadOnlyList<string> paragraphs,
        IReadOnlyList<string> headers, IReadOnlyList<string> footers, IReadOnlyList<string> textBoxes,
        IReadOnlyList<IWorkDiagnostic> diagnostics, bool supportsEditableReconstruction) {
        _source = source;
        Paragraphs = Array.AsReadOnly(paragraphs.ToArray());
        Headers = Array.AsReadOnly(headers.ToArray());
        Footers = Array.AsReadOnly(footers.ToArray());
        TextBoxes = Array.AsReadOnly(textBoxes.ToArray());
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
        _supportsEditableReconstruction = supportsEditableReconstruction;
    }

    /// <summary>Gets body paragraphs in source order.</summary>
    public IReadOnlyList<string> Paragraphs { get; }
    /// <summary>Gets distinct section header text recovered from the source.</summary>
    public IReadOnlyList<string> Headers { get; }
    /// <summary>Gets distinct section footer text recovered from the source.</summary>
    public IReadOnlyList<string> Footers { get; }
    /// <summary>Gets floating text-box content in object order.</summary>
    public IReadOnlyList<string> TextBoxes { get; }
    /// <summary>Gets projection diagnostics.</summary>
    public IReadOnlyList<IWorkDiagnostic> Diagnostics { get; }
    /// <summary>Gets whether the supported editable document structure was recovered completely.</summary>
    public bool HasEditableContent => _supportsEditableReconstruction;

    /// <summary>Creates an import report for an OfficeIMO semantic-owner projection.</summary>
    public IWorkImportReport CreateImportReport(IWorkProjectionKind kind, IWorkPreviewAsset? preview = null) {
        ValidateReportRequest(kind, preview);
        return _source.CreateReport(kind, Diagnostics, preview,
            kind == IWorkProjectionKind.VisualFallback
                ? 0
                : Paragraphs.Count + Headers.Count + Footers.Count + TextBoxes.Count);
    }

    private void ValidateReportRequest(IWorkProjectionKind kind, IWorkPreviewAsset? preview) {
        if (kind == IWorkProjectionKind.EditableReconstruction && !HasEditableContent) {
            throw new InvalidOperationException("Editable Pages content was not recovered.");
        }
        if (kind == IWorkProjectionKind.VisualFallback && preview == null) {
            throw new ArgumentNullException(nameof(preview), "A visual fallback report requires the preview used by the owner.");
        }
    }
}

public sealed partial class IWorkSourceDocument {
    /// <summary>Reads a Pages package into a bounded semantic source projection, or returns a diagnostic-only projection in visual-only mode.</summary>
    public IWorkPagesProjection ReadPages() {
        if (Kind != IWorkDocumentKind.Pages) throw new InvalidOperationException($"The source is {Kind}, not Pages.");
        if (RequestedImportMode == IWorkImportMode.VisualOnly) {
            return new IWorkPagesProjection(this, Array.Empty<string>(), Array.Empty<string>(),
                Array.Empty<string>(), Array.Empty<string>(),
                new[] { IWorkProjectionDiagnostics.SemanticProjectionSkipped }, supportsEditableReconstruction: false);
        }
        return IWorkPagesReader.Read(this);
    }
}

internal static class IWorkPagesReader {
    private const uint DocumentArchive = 10000;
    private const uint SectionArchive = 10011;
    private const uint HeadersFootersArchive = 10143;
    private const uint TextStorageArchive = 2001;
    private const uint ShapeInfoArchive = 2011;

    internal static IWorkPagesProjection Read(IWorkSourceDocument source) {
        var diagnostics = new List<IWorkDiagnostic>();
        var paragraphs = new List<string>();
        var headers = new List<string>();
        var footers = new List<string>();
        var textBoxes = new List<string>();
        IWorkObjectIndex index = source.Index;
        IWorkArchiveRecord? document = index.FirstOfType(DocumentArchive);
        if (document == null) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_PAGES_DOCUMENT_MISSING",
                "No supported Pages document root was found; editable reconstruction is unavailable."));
            return new IWorkPagesProjection(source, paragraphs, headers, footers, textBoxes, diagnostics,
                supportsEditableReconstruction: false);
        }

        bool supportsEditableReconstruction = true;
        IWorkWireMessage documentMessage = index.Message(document);
        IWorkArchiveRecord? body = index.Dereference(documentMessage, 4);
        if (body == null || body.MessageType != TextStorageArchive) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_PAGES_BODY_MISSING",
                "The Pages document root does not reference a supported body text storage.", document.EntryPath, document.Identifier));
        } else {
            foreach (string paragraph in SplitParagraphs(StorageText(index.Message(body)))) paragraphs.Add(paragraph);
            ReadHeadersAndFooters(index, body, headers, footers, diagnostics,
                ref supportsEditableReconstruction);
        }

        var reachable = new HashSet<ulong>(index.ReachableFrom(document)
            .Select(record => record.Identifier));
        var skippedStorages = new HashSet<ulong>();
        if (body != null) skippedStorages.Add(body.Identifier);
        foreach (IWorkArchiveRecord record in index.PrimaryRecords
                     .Where(record => record.MessageType == HeadersFootersArchive)
                     .Where(record => reachable.Contains(record.Identifier))) {
            IWorkWireMessage message = index.Message(record);
            foreach (int field in new[] { 1, 2 }) {
                foreach (IWorkArchiveRecord storage in index.DereferenceAll(message, field)) skippedStorages.Add(storage.Identifier);
            }
        }
        var seenTextStorages = new HashSet<ulong>();
        foreach (IWorkArchiveRecord shape in index.PrimaryRecords
                     .Where(record => record.MessageType == ShapeInfoArchive)
                     .Where(record => reachable.Contains(record.Identifier))
                     .OrderBy(record => record.Identifier)) {
            IWorkArchiveRecord? storage = index.Dereference(index.Message(shape), 2);
            if (storage == null || storage.MessageType != TextStorageArchive
                || skippedStorages.Contains(storage.Identifier) || !seenTextStorages.Add(storage.Identifier)) continue;
            string text = StorageText(index.Message(storage)).Trim();
            if (text.Length == 0) continue;
            textBoxes.Add(text);
        }
        return new IWorkPagesProjection(source, paragraphs, headers, footers, textBoxes, diagnostics,
            supportsEditableReconstruction);
    }

    private static void ReadHeadersAndFooters(IWorkObjectIndex index, IWorkArchiveRecord body,
        List<string> headers, List<string> footers,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        IWorkWireMessage? sectionTable = IWorkObjectIndex.TryGetMessage(index.Message(body), 17);
        if (sectionTable == null) return;
        IReadOnlyList<IWorkWireMessage> entries = IWorkObjectIndex.TryGetMessages(
            sectionTable, 1, out bool malformedEntries);
        if (malformedEntries) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_PAGES_SECTION_UNSUPPORTED",
                "The Pages section table is malformed; editable reconstruction is incomplete.",
                body.EntryPath, body.Identifier));
        }
        var seenHeaders = new HashSet<string>(StringComparer.Ordinal);
        var seenFooters = new HashSet<string>(StringComparer.Ordinal);
        foreach (IWorkWireMessage entry in entries) {
            IReadOnlyList<IWorkArchiveRecord> referencedSections = index.DereferenceAll(
                entry, 2, out int unresolvedSectionCount);
            if (unresolvedSectionCount > 0 || referencedSections.Count != 1
                || referencedSections[0].MessageType != SectionArchive) {
                supportsEditableReconstruction = false;
                if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_PAGES_SECTION_UNSUPPORTED")) {
                    diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                        "IWORK_PAGES_SECTION_UNSUPPORTED",
                        "The Pages section table contains an unresolved or unsupported section; editable reconstruction is incomplete.",
                        body.EntryPath, body.Identifier));
                }
                continue;
            }
            IWorkArchiveRecord section = referencedSections[0];
            IWorkWireMessage sectionMessage = index.Message(section);
            foreach (int field in new[] { 23, 24, 25 }) {
                if (!sectionMessage.HasBytes(field)) continue;
                IWorkArchiveRecord? archive = index.Dereference(sectionMessage, field);
                if (archive == null || archive.MessageType != HeadersFootersArchive) {
                    supportsEditableReconstruction = false;
                    if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_PAGES_HEADER_FOOTER_UNSUPPORTED")) {
                        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                            "IWORK_PAGES_HEADER_FOOTER_UNSUPPORTED",
                            "A Pages section contains an unresolved or unsupported header/footer archive; editable reconstruction is incomplete.",
                            section.EntryPath, section.Identifier));
                    }
                    continue;
                }
                IWorkWireMessage archiveMessage = index.Message(archive);
                AddDistinctStorageText(index, archiveMessage, 1, archive, headers, seenHeaders,
                    diagnostics, ref supportsEditableReconstruction);
                AddDistinctStorageText(index, archiveMessage, 2, archive, footers, seenFooters,
                    diagnostics, ref supportsEditableReconstruction);
            }
        }
    }

    private static void AddDistinctStorageText(IWorkObjectIndex index, IWorkWireMessage message, int field,
        IWorkArchiveRecord archive, List<string> destination, HashSet<string> seen,
        List<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        IReadOnlyList<IWorkArchiveRecord> storages = index.DereferenceAll(
            message, field, out int unresolvedStorageCount);
        if (unresolvedStorageCount > 0) {
            supportsEditableReconstruction = false;
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_PAGES_HEADER_FOOTER_UNSUPPORTED",
                "A Pages header or footer contains an unresolved text reference; editable reconstruction is incomplete.",
                archive.EntryPath, archive.Identifier));
        }
        foreach (IWorkArchiveRecord storage in storages) {
            if (storage.MessageType != TextStorageArchive) continue;
            string text = StorageText(index.Message(storage)).Trim();
            if (text.Length == 0 || !seen.Add(text)) continue;
            destination.Add(text);
        }
    }

    internal static string StorageText(IWorkWireMessage storage) {
        var text = new System.Text.StringBuilder();
        foreach (byte[] bytes in storage.GetRepeatedBytes(3)) {
            try {
                text.Append(new System.Text.UTF8Encoding(false, true).GetString(bytes));
            } catch (System.Text.DecoderFallbackException) {
                // The field is not a valid text run on this producer version.
            }
        }
        return CleanText(text.ToString());
    }

    internal static string CleanText(string value) => value
        .Replace("\uFFFC", string.Empty)
        .Replace("\uFFFB", string.Empty)
        .Replace("\u2028", "\n")
        .Replace("\u2029", "\n")
        .Replace("\r\n", "\n")
        .Replace("\r", "\n");

    private static IEnumerable<string> SplitParagraphs(string text) => text
        .Split(new[] { '\n' }, StringSplitOptions.None)
        .Select(paragraph => paragraph.Trim())
        .Where(paragraph => paragraph.Length > 0);
}
