using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Fluent five-format image export for one PDF page.</summary>
public sealed class PdfPageImageExportBuilder : OfficeImageExportBuilder<PdfPageImageExportBuilder, PdfImageExportOptions> {
    internal PdfPageImageExportBuilder(PdfReadPage page, PdfImageExportOptions? options = null)
        : base(
            options?.Clone() ?? new PdfImageExportOptions(),
            (format, effective, cancellationToken) =>
                PdfImageExportEngine.Export(page, format, effective, cancellationToken: cancellationToken)) {
    }

    /// <summary>Fits the output within the requested maximum pixel width or height.</summary>
    public PdfPageImageExportBuilder AsThumbnail(int maximumDimension) {
        Guard.PositiveInteger(maximumDimension, nameof(maximumDimension));
        Options.ThumbnailMaxDimension = maximumDimension;
        return this;
    }
}

/// <summary>Fluent five-format batch image export for PDF pages.</summary>
public sealed class PdfDocumentImageExportBuilder : OfficeImageExportBatchBuilder<PdfDocumentImageExportBuilder, PdfImageExportOptions> {
    private readonly PageSelectionState _selection;

    internal PdfDocumentImageExportBuilder(
        PdfReadDocument document,
        PdfImageExportOptions? options = null,
        IReadOnlyList<OfficeImageExportDiagnostic>? initialDiagnostics = null)
        : this(
            PdfImageExportDocumentSource.FromLoaded(document),
            options?.Clone() ?? new PdfImageExportOptions(),
            new PageSelectionState(),
            initialDiagnostics) {
    }

    internal PdfDocumentImageExportBuilder(
        PdfDocument document,
        PdfImageExportOptions? options = null,
        IReadOnlyList<OfficeImageExportDiagnostic>? initialDiagnostics = null)
        : this(
            PdfImageExportDocumentSource.FromDeferred(() => document.GetReadSnapshot().Document),
            options?.Clone() ?? new PdfImageExportOptions(),
            new PageSelectionState(),
            initialDiagnostics) {
    }

    internal PdfDocumentImageExportBuilder(
        PdfDocumentConversionResult conversion,
        PdfImageExportOptions? options = null,
        IReadOnlyList<OfficeImageExportDiagnostic>? initialDiagnostics = null)
        : this(
            PdfImageExportDocumentSource.FromDeferred(() => PdfReadDocument.Open(conversion.ToBytes())),
            options?.Clone() ?? new PdfImageExportOptions(),
            new PageSelectionState(),
            initialDiagnostics) {
    }

    private PdfDocumentImageExportBuilder(
        PdfImageExportDocumentSource source,
        PdfImageExportOptions options,
        PageSelectionState selection,
        IReadOnlyList<OfficeImageExportDiagnostic>? initialDiagnostics)
        : base(
            options,
            (format, effective) => Export(source, selection, format, effective, initialDiagnostics),
            (format, effective, consumer, cancellationToken) =>
                ExportEach(source, selection, format, effective, consumer, initialDiagnostics, cancellationToken)) {
        _selection = selection;
    }

    /// <summary>Selects caller-ordered one-based PDF pages.</summary>
    public PdfDocumentImageExportBuilder Pages(PdfPageSelection selection) {
        _selection.Selection = selection ?? throw new ArgumentNullException(nameof(selection));
        _selection.Selector = null;
        return this;
    }

    /// <summary>Selects document-relative pages such as <c>1-3,last</c>.</summary>
    public PdfDocumentImageExportBuilder Pages(string selector) {
        if (string.IsNullOrWhiteSpace(selector)) throw new ArgumentException("A page selector is required.", nameof(selector));
        _selection.Selector = PdfPageSelector.Parse(selector);
        _selection.Selection = null;
        return this;
    }

    /// <summary>Exports all document pages in source order.</summary>
    public PdfDocumentImageExportBuilder AllPages() {
        _selection.Selection = null;
        _selection.Selector = null;
        return this;
    }

    /// <summary>Fits every output page within the requested maximum pixel width or height.</summary>
    public PdfDocumentImageExportBuilder AsThumbnails(int maximumDimension) {
        Guard.PositiveInteger(maximumDimension, nameof(maximumDimension));
        Options.ThumbnailMaxDimension = maximumDimension;
        return this;
    }

    /// <summary>Limits the number of selected pages accepted by one export.</summary>
    public PdfDocumentImageExportBuilder WithMaximumPages(int maximumPages) {
        Guard.PositiveInteger(maximumPages, nameof(maximumPages));
        Options.MaximumOutputCount = maximumPages;
        return this;
    }

    private sealed class PageSelectionState {
        internal PdfPageSelection? Selection { get; set; }
        internal PdfPageSelector? Selector { get; set; }

        internal PdfPageSelection? Resolve(PdfReadDocument document) =>
            Selector?.ResolveSelection(document.Pages.Count) ?? Selection;
    }

    private static IReadOnlyList<OfficeImageExportResult> Export(
        PdfImageExportDocumentSource source,
        PageSelectionState selection,
        OfficeImageExportFormat format,
        PdfImageExportOptions options,
        IReadOnlyList<OfficeImageExportDiagnostic>? initialDiagnostics) {
        PdfReadDocument document = source.Get(CancellationToken.None);
        return PdfImageExportEngine.Export(
            document,
            format,
            options,
            selection.Resolve(document),
            initialDiagnostics);
    }

    private static void ExportEach(
        PdfImageExportDocumentSource source,
        PageSelectionState selection,
        OfficeImageExportFormat format,
        PdfImageExportOptions options,
        OfficeImageExportConsumer consumer,
        IReadOnlyList<OfficeImageExportDiagnostic>? initialDiagnostics,
        CancellationToken cancellationToken) {
        PdfReadDocument document = source.Get(cancellationToken);
        PdfImageExportEngine.ExportEach(
            document,
            format,
            options,
            selection.Resolve(document),
            consumer,
            initialDiagnostics,
            cancellationToken);
    }

    private sealed class PdfImageExportDocumentSource {
        private readonly Func<PdfReadDocument>? _factory;
        private readonly object _sync = new();
        private PdfReadDocument? _document;

        private PdfImageExportDocumentSource(PdfReadDocument document) {
            _document = document;
        }

        private PdfImageExportDocumentSource(Func<PdfReadDocument> factory) {
            _factory = factory;
        }

        internal static PdfImageExportDocumentSource FromLoaded(PdfReadDocument document) {
            Guard.NotNull(document, nameof(document));
            return new PdfImageExportDocumentSource(document);
        }

        internal static PdfImageExportDocumentSource FromDeferred(Func<PdfReadDocument> factory) {
            Guard.NotNull(factory, nameof(factory));
            return new PdfImageExportDocumentSource(factory);
        }

        internal PdfReadDocument Get(CancellationToken cancellationToken) {
            cancellationToken.ThrowIfCancellationRequested();
            if (_document != null) return _document;

            lock (_sync) {
                cancellationToken.ThrowIfCancellationRequested();
                return _document ??= _factory!();
            }
        }
    }
}

/// <summary>Canonical PDF page image-export entry points.</summary>
public static class PdfImageExportExtensions {
    /// <summary>Exports pages from an authored or opened PDF document using the shared five-format result contract.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this PdfDocument document,
        OfficeImageExportFormat format,
        PdfImageExportOptions? options = null,
        PdfPageSelection? selection = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(document, nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        var snapshot = document.GetReadSnapshot();
        return PdfImageExportEngine.Export(
            snapshot.Document,
            format,
            options?.Clone() ?? new PdfImageExportOptions(),
            selection,
            initialDiagnostics: null,
            cancellationToken);
    }

    /// <summary>Exports one loaded PDF page using the shared five-format result contract.</summary>
    public static OfficeImageExportResult ExportImage(
        this PdfReadPage page,
        OfficeImageExportFormat format,
        PdfImageExportOptions? options = null) =>
        PdfImageExportEngine.Export(page, format, options?.Clone() ?? new PdfImageExportOptions());

    /// <summary>Exports selected loaded PDF pages using the shared five-format result contract.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this PdfReadDocument document,
        OfficeImageExportFormat format,
        PdfImageExportOptions? options = null,
        PdfPageSelection? selection = null,
        CancellationToken cancellationToken = default) =>
        PdfImageExportEngine.Export(
            document,
            format,
            options?.Clone() ?? new PdfImageExportOptions(),
            selection,
            initialDiagnostics: null,
            cancellationToken);

    /// <summary>
    /// Exports pages from a source-to-PDF conversion while preserving its conversion diagnostics.
    /// This is the shared paged-image bridge for Markdown, AsciiDoc, LaTeX, RTF, OneNote, and other PDF adapters.
    /// </summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this PdfDocumentConversionResult conversion,
        OfficeImageExportFormat format,
        PdfImageExportOptions? options = null,
        PdfPageSelection? selection = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(conversion, nameof(conversion));
        cancellationToken.ThrowIfCancellationRequested();
        return PdfImageExportEngine.Export(
            PdfReadDocument.Open(conversion.ToBytes()),
            format,
            options?.Clone() ?? new PdfImageExportOptions(),
            selection,
            PdfImageExportEngine.MapConversionDiagnostics(conversion),
            cancellationToken);
    }

    /// <summary>Starts fluent image export for one loaded PDF page.</summary>
    public static PdfPageImageExportBuilder ToImage(this PdfReadPage page) =>
        new PdfPageImageExportBuilder(page);

    /// <summary>Starts fluent image export for one loaded PDF page using a cloned options snapshot.</summary>
    public static PdfPageImageExportBuilder ToImage(this PdfReadPage page, PdfImageExportOptions options) =>
        new PdfPageImageExportBuilder(page, options ?? throw new ArgumentNullException(nameof(options)));

    /// <summary>Starts fluent image export for all pages in a loaded PDF document.</summary>
    public static PdfDocumentImageExportBuilder ToImages(this PdfReadDocument document) =>
        CreateDocumentBuilder(document, options: null);

    /// <summary>Starts fluent image export using a cloned options snapshot.</summary>
    public static PdfDocumentImageExportBuilder ToImages(this PdfReadDocument document, PdfImageExportOptions options) =>
        CreateDocumentBuilder(document, options ?? throw new ArgumentNullException(nameof(options)));

    /// <summary>
    /// Starts fluent image export for all pages in an authored or opened PDF document.
    /// The read snapshot is captured on first export so a pre-canceled operation does not materialize the document.
    /// </summary>
    public static PdfDocumentImageExportBuilder ToImages(this PdfDocument document) =>
        CreateDocumentBuilder(document, options: null);

    /// <summary>
    /// Starts fluent image export for an authored or opened PDF document using a cloned options snapshot.
    /// The read snapshot is captured on first export.
    /// </summary>
    public static PdfDocumentImageExportBuilder ToImages(this PdfDocument document, PdfImageExportOptions options) =>
        CreateDocumentBuilder(document, options ?? throw new ArgumentNullException(nameof(options)));

    /// <summary>
    /// Starts fluent paged-image export from a source-to-PDF conversion and carries its diagnostics into every page result.
    /// </summary>
    public static PdfDocumentImageExportBuilder ToImages(this PdfDocumentConversionResult conversion) =>
        CreateDocumentBuilder(conversion, options: null);

    /// <summary>Starts fluent paged-image export with a cloned options snapshot and retained conversion diagnostics.</summary>
    public static PdfDocumentImageExportBuilder ToImages(
        this PdfDocumentConversionResult conversion,
        PdfImageExportOptions options) =>
        CreateDocumentBuilder(conversion, options ?? throw new ArgumentNullException(nameof(options)));

    private static PdfDocumentImageExportBuilder CreateDocumentBuilder(
        PdfReadDocument document,
        PdfImageExportOptions? options) {
        Guard.NotNull(document, nameof(document));
        return new PdfDocumentImageExportBuilder(document, options);
    }

    private static PdfDocumentImageExportBuilder CreateDocumentBuilder(
        PdfDocument document,
        PdfImageExportOptions? options) {
        Guard.NotNull(document, nameof(document));
        return new PdfDocumentImageExportBuilder(document, options);
    }

    private static PdfDocumentImageExportBuilder CreateDocumentBuilder(
        PdfDocumentConversionResult conversion,
        PdfImageExportOptions? options) {
        Guard.NotNull(conversion, nameof(conversion));
        return new PdfDocumentImageExportBuilder(
            conversion,
            options,
            PdfImageExportEngine.MapConversionDiagnostics(conversion));
    }
}
