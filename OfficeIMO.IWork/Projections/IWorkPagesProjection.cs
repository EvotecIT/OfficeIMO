using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

/// <summary>Read-only Pages structure recovered from a shared IWA object graph.</summary>
public sealed class IWorkPagesProjection {
    private readonly IWorkSourceDocument _source;
    private readonly IReadOnlyCollection<ulong> _recognizedIdentifiers;

    internal IWorkPagesProjection(IWorkSourceDocument source, IReadOnlyList<string> paragraphs,
        IReadOnlyList<string> headers, IReadOnlyList<string> footers, IReadOnlyList<string> textBoxes,
        IReadOnlyCollection<ulong> recognizedIdentifiers, IReadOnlyList<IWorkDiagnostic> diagnostics) {
        _source = source;
        Paragraphs = paragraphs;
        Headers = headers;
        Footers = footers;
        TextBoxes = textBoxes;
        _recognizedIdentifiers = recognizedIdentifiers;
        Diagnostics = diagnostics;
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
    /// <summary>Gets whether supported editable content was recovered.</summary>
    public bool HasEditableContent => Paragraphs.Count > 0 || Headers.Count > 0 || Footers.Count > 0 || TextBoxes.Count > 0;

    /// <summary>Creates an import report for an OfficeIMO semantic-owner projection.</summary>
    public IWorkImportReport CreateImportReport(IWorkProjectionKind kind, IWorkPreviewAsset? preview = null) {
        ValidateReportRequest(kind, preview);
        return _source.CreateReport(kind, _recognizedIdentifiers, Diagnostics, preview,
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
                Array.Empty<string>(), Array.Empty<string>(), Array.Empty<ulong>(),
                new[] { IWorkProjectionDiagnostics.SemanticProjectionSkipped });
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
        var recognized = new HashSet<ulong>();
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
            return new IWorkPagesProjection(source, paragraphs, headers, footers, textBoxes, recognized, diagnostics);
        }

        recognized.Add(document.Identifier);
        IWorkWireMessage documentMessage = index.Message(document);
        IWorkArchiveRecord? body = index.Dereference(documentMessage, 4);
        if (body == null || body.MessageType != TextStorageArchive) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning, "IWORK_PAGES_BODY_MISSING",
                "The Pages document root does not reference a supported body text storage.", document.EntryPath, document.Identifier));
        } else {
            recognized.Add(body.Identifier);
            foreach (string paragraph in SplitParagraphs(StorageText(index.Message(body)))) paragraphs.Add(paragraph);
            ReadHeadersAndFooters(index, body, recognized, headers, footers);
        }

        var skippedStorages = new HashSet<ulong>();
        if (body != null) skippedStorages.Add(body.Identifier);
        foreach (IWorkArchiveRecord record in index.PrimaryRecords.Where(record => record.MessageType == HeadersFootersArchive)) {
            IWorkWireMessage message = index.Message(record);
            foreach (int field in new[] { 1, 2 }) {
                foreach (IWorkArchiveRecord storage in index.DereferenceAll(message, field)) skippedStorages.Add(storage.Identifier);
            }
        }
        var seenTextStorages = new HashSet<ulong>();
        foreach (IWorkArchiveRecord shape in index.PrimaryRecords
                     .Where(record => record.MessageType == ShapeInfoArchive)
                     .OrderBy(record => record.Identifier)) {
            IWorkArchiveRecord? storage = index.Dereference(index.Message(shape), 2);
            if (storage == null || storage.MessageType != TextStorageArchive
                || skippedStorages.Contains(storage.Identifier) || !seenTextStorages.Add(storage.Identifier)) continue;
            string text = StorageText(index.Message(storage)).Trim();
            if (text.Length == 0) continue;
            recognized.Add(shape.Identifier);
            recognized.Add(storage.Identifier);
            textBoxes.Add(text);
        }
        return new IWorkPagesProjection(source, paragraphs, headers, footers, textBoxes, recognized, diagnostics);
    }

    private static void ReadHeadersAndFooters(IWorkObjectIndex index, IWorkArchiveRecord body,
        HashSet<ulong> recognized, List<string> headers, List<string> footers) {
        IWorkWireMessage? sectionTable = IWorkObjectIndex.TryGetMessage(index.Message(body), 17);
        if (sectionTable == null) return;
        foreach (IWorkWireMessage entry in IWorkObjectIndex.TryGetMessages(sectionTable, 1)) {
            IWorkArchiveRecord? section = index.Dereference(entry, 2);
            if (section == null || section.MessageType != SectionArchive) continue;
            recognized.Add(section.Identifier);
            IWorkWireMessage sectionMessage = index.Message(section);
            foreach (int field in new[] { 23, 24, 25 }) {
                IWorkArchiveRecord? archive = index.Dereference(sectionMessage, field);
                if (archive == null || archive.MessageType != HeadersFootersArchive) continue;
                recognized.Add(archive.Identifier);
                IWorkWireMessage archiveMessage = index.Message(archive);
                AddDistinctStorageText(index, archiveMessage, 1, recognized, headers);
                AddDistinctStorageText(index, archiveMessage, 2, recognized, footers);
            }
        }
    }

    private static void AddDistinctStorageText(IWorkObjectIndex index, IWorkWireMessage message, int field,
        HashSet<ulong> recognized, List<string> destination) {
        foreach (IWorkArchiveRecord storage in index.DereferenceAll(message, field)) {
            if (storage.MessageType != TextStorageArchive) continue;
            string text = StorageText(index.Message(storage)).Trim();
            if (text.Length == 0 || destination.Contains(text)) continue;
            recognized.Add(storage.Identifier);
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
