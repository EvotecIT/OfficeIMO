using OfficeIMO.Drawing;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Pdf;

public static partial class PdfImageExportExtensions {
    internal static OfficeImageExportBatchOperation<PdfImageExportOptions> CreateImageExportOperation(
        this PdfDocument document, PdfImageExportOptions? options = null) {
        Guard.NotNull(document, nameof(document));
        return CreateOperation(PdfImageExportDocumentSource.FromDeferred(
            token => document.GetReadSnapshot(cancellationToken: token).Document), options);
    }

    internal static OfficeImageExportBatchOperation<PdfImageExportOptions> CreateImageExportOperation(
        this PdfReadDocument document, PdfImageExportOptions? options = null) =>
        CreateOperation(PdfImageExportDocumentSource.FromLoaded(document), options);

    // Adapters retain source diagnostics, including findings emitted while materializing the PDF.
    internal static OfficeImageExportBatchOperation<PdfImageExportOptions> CreateImageExportOperation(
        this PdfDocumentConversionResult conversion, PdfImageExportOptions? options = null) {
        Guard.NotNull(conversion, nameof(conversion));
        return CreateOperation(PdfImageExportDocumentSource.FromDeferred(
            token => PdfReadDocument.Open(conversion.ToBytes(token), options: null, cancellationToken: token),
            () => PdfImageExportEngine.MapConversionDiagnostics(conversion)), options);
    }

    /// <summary>Streams page results to in-process hosts without exposing the execution operation.</summary>
    internal static void ExportImageResultsEach(this PdfDocument document, PdfImageExportOptions options,
        OfficeImageExportConsumer consumer, CancellationToken cancellationToken = default) =>
        document.CreateImageExportOperation(options).ExportEach(consumer, cancellationToken);

    /// <summary>Returns per-page evidence for a single-image host response.</summary>
    internal static IReadOnlyList<OfficeImageExportResult> ExportImageResults(this PdfDocument document,
        PdfImageExportOptions options) => document.CreateImageExportOperation(options).Export();

    /// <summary>Writes bounded staged files and returns metadata for verified host publication.</summary>
    internal static Task<OfficeImageExportBatchSaveResult> SaveImageFilesAsync(this PdfDocument document,
        string folderPath, PdfImageExportOptions options, CancellationToken cancellationToken) =>
        document.CreateImageExportOperation(options).SaveFilesAsync(folderPath, cancellationToken);

    private static OfficeImageExportBatchOperation<PdfImageExportOptions> CreateOperation(
        PdfImageExportDocumentSource source, PdfImageExportOptions? options) {
        PdfImageExportOptions snapshot = options?.Clone() ?? new PdfImageExportOptions();
        return new OfficeImageExportBatchOperation<PdfImageExportOptions>(snapshot, snapshot.Format,
            snapshot.FileConflictPolicy,
            (format, effective) => PdfImageExportEngine.Export(source.Get, format, effective,
                document => effective.PageSelector?.ResolveSelection(document.Pages.Count),
                diagnosticsFactory: source.GetDiagnostics),
            exportEach: (format, effective, consumer, token) => PdfImageExportEngine.ExportEach(
                source.Get, format, effective,
                document => effective.PageSelector?.ResolveSelection(document.Pages.Count), consumer,
                diagnosticsFactory: source.GetDiagnostics, cancellationToken: token));
    }

    private sealed class PdfImageExportDocumentSource {
        private readonly Func<CancellationToken, PdfReadDocument>? _factory;
        private readonly Func<IReadOnlyList<OfficeImageExportDiagnostic>>? _diagnosticsFactory;
        private IReadOnlyList<OfficeImageExportDiagnostic> _diagnostics = Array.Empty<OfficeImageExportDiagnostic>();
        private readonly object _sync = new();
        private PdfReadDocument? _document;

        private PdfImageExportDocumentSource(PdfReadDocument document) {
            _document = document;
        }

        private PdfImageExportDocumentSource(Func<CancellationToken, PdfReadDocument> factory,
            Func<IReadOnlyList<OfficeImageExportDiagnostic>>? diagnosticsFactory) {
            _factory = factory;
            _diagnosticsFactory = diagnosticsFactory;
        }

        internal static PdfImageExportDocumentSource FromLoaded(PdfReadDocument document) {
            Guard.NotNull(document, nameof(document));
            return new PdfImageExportDocumentSource(document);
        }

        internal static PdfImageExportDocumentSource FromDeferred(Func<CancellationToken, PdfReadDocument> factory,
            Func<IReadOnlyList<OfficeImageExportDiagnostic>>? diagnosticsFactory = null) {
            Guard.NotNull(factory, nameof(factory));
            return new PdfImageExportDocumentSource(factory, diagnosticsFactory);
        }

        internal IReadOnlyList<OfficeImageExportDiagnostic> GetDiagnostics() {
            lock (_sync) return _diagnostics;
        }

        internal PdfReadDocument Get(CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            lock (_sync) {
                cancellationToken.ThrowIfCancellationRequested();
                if (_document != null) return _document;
                PdfReadDocument document = _factory!(cancellationToken);
                IReadOnlyList<OfficeImageExportDiagnostic> diagnostics = _diagnosticsFactory?.Invoke()
                    ?? Array.Empty<OfficeImageExportDiagnostic>();
                cancellationToken.ThrowIfCancellationRequested();
                _diagnostics = diagnostics;
                _document = document;
                return document;
            }
        }
    }
}
