using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

/// <summary>Fluent image export for a Visio page.</summary>
public sealed class VisioPageImageExportBuilder : OfficeImageExportBuilder<VisioPageImageExportBuilder, VisioImageExportOptions> {
    internal VisioPageImageExportBuilder(VisioPage page, VisioImageExportOptions? options = null)
        : base(options?.Clone() ?? new VisioImageExportOptions(), CreateExporter(page)) {
    }

    /// <summary>Includes or excludes shape and connector text.</summary>
    public VisioPageImageExportBuilder IncludeText(bool include = true) { Options.RenderText = include; return this; }

    /// <summary>Includes or excludes built-in stencil artwork.</summary>
    public VisioPageImageExportBuilder IncludeStencilArtwork(bool include = true) { Options.RenderStencilArtwork = include; return this; }

    /// <summary>Includes or excludes connector labels.</summary>
    public VisioPageImageExportBuilder IncludeConnectorLabels(bool include = true) { Options.RenderConnectorLabels = include; return this; }

    private static Func<OfficeImageExportFormat, VisioImageExportOptions, System.Threading.CancellationToken, OfficeImageExportResult> CreateExporter(VisioPage page) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        return (format, options, cancellationToken) => page.ExportImage(format, options, cancellationToken);
    }
}

/// <summary>Fluent image export for one selected Visio document page.</summary>
public sealed class VisioDocumentImageExportBuilder : OfficeImageExportBuilder<VisioDocumentImageExportBuilder, VisioImageExportOptions> {
    internal VisioDocumentImageExportBuilder(VisioDocument document, VisioImageExportOptions? options = null)
        : base(options?.Clone() ?? new VisioImageExportOptions(), CreateExporter(document)) {
    }

    /// <summary>Selects the zero-based document page to export.</summary>
    public VisioDocumentImageExportBuilder Page(int pageIndex) {
        if (pageIndex < 0) throw new ArgumentOutOfRangeException(nameof(pageIndex));
        Options.PageIndex = pageIndex;
        return this;
    }

    /// <summary>Includes or excludes shape and connector text.</summary>
    public VisioDocumentImageExportBuilder IncludeText(bool include = true) { Options.RenderText = include; return this; }

    private static Func<OfficeImageExportFormat, VisioImageExportOptions, System.Threading.CancellationToken, OfficeImageExportResult> CreateExporter(VisioDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return (format, options, cancellationToken) => document.ExportImage(format, options, cancellationToken);
    }
}

/// <summary>Fluent batch image export for Visio document pages.</summary>
public sealed class VisioDocumentImageBatchExportBuilder : OfficeImageExportBatchBuilder<VisioDocumentImageBatchExportBuilder, VisioImageExportOptions> {
    internal VisioDocumentImageBatchExportBuilder(VisioDocument document, VisioImageExportOptions? options = null)
        : base(
            options?.Clone() ?? new VisioImageExportOptions(),
            CreateExporter(document),
            CreateStreamingExporter(document)) {
    }

    /// <summary>Exports from the specified zero-based page index.</summary>
    public VisioDocumentImageBatchExportBuilder FromPage(int pageIndex) {
        if (pageIndex < 0) throw new ArgumentOutOfRangeException(nameof(pageIndex));
        Options.PageIndex = pageIndex;
        return this;
    }

    /// <summary>Limits batch output to the requested number of pages.</summary>
    public VisioDocumentImageBatchExportBuilder TakePages(int pageCount) {
        if (pageCount < 1) throw new ArgumentOutOfRangeException(nameof(pageCount));
        Options.PageCount = pageCount;
        return this;
    }

    /// <summary>Exports all pages from the beginning of the document.</summary>
    public VisioDocumentImageBatchExportBuilder AllPages() {
        Options.PageIndex = 0;
        Options.PageCount = null;
        return this;
    }

    /// <summary>Includes or excludes shape and connector text.</summary>
    public VisioDocumentImageBatchExportBuilder IncludeText(bool include = true) { Options.RenderText = include; return this; }

    private static Func<OfficeImageExportFormat, VisioImageExportOptions, IReadOnlyList<OfficeImageExportResult>> CreateExporter(VisioDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return (format, options) => document.ExportImages(format, options);
    }

    private static Action<OfficeImageExportFormat, VisioImageExportOptions, OfficeImageExportConsumer, System.Threading.CancellationToken> CreateStreamingExporter(VisioDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return (format, options, consumer, cancellationToken) =>
            document.ExportImages(format, consumer, options, cancellationToken);
    }
}

/// <summary>Fluent image-export entry points for Visio documents and pages.</summary>
public static class VisioImageExportFluentExtensions {
    /// <summary>Starts fluent image export for this Visio page.</summary>
    public static VisioPageImageExportBuilder ToImage(this VisioPage page) => new VisioPageImageExportBuilder(page);

    /// <summary>Starts fluent image export using a cloned options snapshot.</summary>
    public static VisioPageImageExportBuilder ToImage(this VisioPage page, VisioImageExportOptions options) =>
        new VisioPageImageExportBuilder(page, options ?? throw new ArgumentNullException(nameof(options)));

    /// <summary>Starts fluent image export for one selected page in this Visio document.</summary>
    public static VisioDocumentImageExportBuilder ToImage(this VisioDocument document) => new VisioDocumentImageExportBuilder(document);

    /// <summary>Starts fluent selected-page export using a cloned options snapshot.</summary>
    public static VisioDocumentImageExportBuilder ToImage(this VisioDocument document, VisioImageExportOptions options) =>
        new VisioDocumentImageExportBuilder(document, options ?? throw new ArgumentNullException(nameof(options)));

    /// <summary>Starts fluent batch image export for pages in this Visio document.</summary>
    public static VisioDocumentImageBatchExportBuilder ToImages(this VisioDocument document) => new VisioDocumentImageBatchExportBuilder(document);

    /// <summary>Starts fluent page-batch export using a cloned options snapshot.</summary>
    public static VisioDocumentImageBatchExportBuilder ToImages(this VisioDocument document, VisioImageExportOptions options) =>
        new VisioDocumentImageBatchExportBuilder(document, options ?? throw new ArgumentNullException(nameof(options)));
}
