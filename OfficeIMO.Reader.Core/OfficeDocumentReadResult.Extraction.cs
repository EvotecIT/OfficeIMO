using System.Threading;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentReadResultExtensions {
    /// <summary>Extracts tables from a materialized read without reading the source again.</summary>
    public static IReadOnlyList<ReaderTable> ExtractTables(this OfficeDocumentReadResult document, CancellationToken cancellationToken = default) =>
        Materialize(RequireDocument(document).EnumerateTables(), cancellationToken);

    /// <summary>Exports discovered tables as deterministic CSV, Markdown and JSON payloads.</summary>
    public static IReadOnlyList<ReaderTableExportBundle> ExportTables(this OfficeDocumentReadResult document, bool indentedJson = false, CancellationToken cancellationToken = default) =>
        DocumentReaderEngine.ExportTables(document.ExtractTables(cancellationToken), indentedJson, cancellationToken);

    /// <summary>Extracts visuals from a materialized read.</summary>
    public static IReadOnlyList<ReaderVisual> ExtractVisuals(this OfficeDocumentReadResult document, CancellationToken cancellationToken = default) =>
        Materialize(RequireDocument(document).EnumerateVisuals(), cancellationToken);

    /// <summary>Exports discovered visuals and JSON sidecars.</summary>
    public static IReadOnlyList<ReaderVisualExportBundle> ExportVisuals(this OfficeDocumentReadResult document, bool indentedJson = false, CancellationToken cancellationToken = default) =>
        DocumentReaderEngine.ExportVisuals(document.ExtractVisuals(cancellationToken), indentedJson, cancellationToken);

    /// <summary>Extracts assets already attached to a materialized read.</summary>
    public static IReadOnlyList<OfficeDocumentAsset> ExtractAssets(this OfficeDocumentReadResult document, Func<OfficeDocumentAsset, bool>? predicate = null, CancellationToken cancellationToken = default) =>
        DocumentReaderEngine.ExtractAssets(RequireDocument(document), predicate, cancellationToken);

    /// <summary>Extracts a bounded structured view from a materialized read.</summary>
    public static OfficeDocumentStructuredExtractionResult ExtractStructured(this OfficeDocumentReadResult document, OfficeDocumentStructuredExtractionOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeDocumentStructuredExtractor.Extract(RequireDocument(document), options, cancellationToken);

    /// <summary>ExtractTables from an existing pipeline stage without rereading the source.</summary>
    public static IReadOnlyList<ReaderTable> ExtractTables(this IEnumerable<ReaderChunk> items, CancellationToken cancellationToken = default) =>
        DocumentReaderEngine.ExtractTables(items, cancellationToken);

    /// <summary>ExtractVisuals from an existing pipeline stage without rereading the source.</summary>
    public static IReadOnlyList<ReaderVisual> ExtractVisuals(this IEnumerable<ReaderChunk> items, CancellationToken cancellationToken = default) =>
        DocumentReaderEngine.ExtractVisuals(items, cancellationToken);

    /// <summary>ExportTables from an existing pipeline stage without rereading the source.</summary>
    public static IReadOnlyList<ReaderTableExportBundle> ExportTables(this IEnumerable<ReaderTable> items, bool indentedJson = false, CancellationToken cancellationToken = default) =>
        DocumentReaderEngine.ExportTables(items, indentedJson, cancellationToken);

    /// <summary>ExportVisuals from an existing pipeline stage without rereading the source.</summary>
    public static IReadOnlyList<ReaderVisualExportBundle> ExportVisuals(this IEnumerable<ReaderVisual> items, bool indentedJson = false, CancellationToken cancellationToken = default) =>
        DocumentReaderEngine.ExportVisuals(items, indentedJson, cancellationToken);

    private static IReadOnlyList<T> Materialize<T>(IEnumerable<T> items, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var result = new List<T>();
        foreach (T item in items) {
            cancellationToken.ThrowIfCancellationRequested();
            result.Add(item);
        }
        return result.ToArray();
    }

    private static OfficeDocumentReadResult RequireDocument(OfficeDocumentReadResult document) =>
        document ?? throw new ArgumentNullException(nameof(document));
}
