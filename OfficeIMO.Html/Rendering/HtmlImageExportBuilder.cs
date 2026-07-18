using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Fluent image export for one rendered HTML surface.</summary>
public sealed class HtmlPageImageExportBuilder : OfficeImageExportBuilder<HtmlPageImageExportBuilder, HtmlRenderOptions> {
    private readonly PageSelection _selection;

    internal HtmlPageImageExportBuilder(HtmlConversionDocument document, HtmlRenderOptions? options = null)
        : this(document, options?.Clone() ?? new HtmlRenderOptions(), new PageSelection()) {
    }

    private HtmlPageImageExportBuilder(
        HtmlConversionDocument document,
        HtmlRenderOptions options,
        PageSelection selection)
        : base(options, CreateExporter(document, selection), CreateAsyncExporter(document, selection)) {
        _selection = selection;
    }

    /// <summary>Selects the zero-based rendered page or continuous surface to export.</summary>
    public HtmlPageImageExportBuilder Page(int pageIndex) {
        if (pageIndex < 0) throw new ArgumentOutOfRangeException(nameof(pageIndex));
        _selection.PageIndex = pageIndex;
        return this;
    }

    /// <summary>Uses continuous layout for the rendered surface.</summary>
    public HtmlPageImageExportBuilder Continuous() {
        Options.Mode = HtmlRenderMode.Continuous;
        _selection.PageIndex = 0;
        return this;
    }

    /// <summary>Uses paged layout and selects the requested zero-based page.</summary>
    public HtmlPageImageExportBuilder Paged(int pageIndex = 0) {
        Options.Mode = HtmlRenderMode.Paged;
        return Page(pageIndex);
    }

    private static Func<OfficeImageExportFormat, HtmlRenderOptions, OfficeImageExportResult> CreateExporter(
        HtmlConversionDocument document,
        PageSelection selection) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return (format, options) => document.ExportImage(format, options, selection.PageIndex);
    }

    private static Func<OfficeImageExportFormat, HtmlRenderOptions, CancellationToken, Task<OfficeImageExportResult>> CreateAsyncExporter(
        HtmlConversionDocument document,
        PageSelection selection) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return (format, options, cancellationToken) =>
            document.ExportImageAsync(format, options, selection.PageIndex, cancellationToken);
    }

    private sealed class PageSelection {
        internal int PageIndex { get; set; }
    }
}

/// <summary>Fluent batch image export for all rendered HTML pages.</summary>
public sealed class HtmlPageImageBatchExportBuilder : OfficeImageExportBatchBuilder<HtmlPageImageBatchExportBuilder, HtmlRenderOptions> {
    internal HtmlPageImageBatchExportBuilder(HtmlConversionDocument document, HtmlRenderOptions? options = null)
        : base(options?.Clone() ?? new HtmlRenderOptions(), CreateExporter(document), CreateAsyncExporter(document)) {
    }

    /// <summary>Uses continuous layout, which produces one rendered surface.</summary>
    public HtmlPageImageBatchExportBuilder Continuous() {
        Options.Mode = HtmlRenderMode.Continuous;
        return this;
    }

    /// <summary>Uses paged layout and exports every rendered page.</summary>
    public HtmlPageImageBatchExportBuilder Paged() {
        Options.Mode = HtmlRenderMode.Paged;
        return this;
    }

    private static Func<OfficeImageExportFormat, HtmlRenderOptions, IReadOnlyList<OfficeImageExportResult>> CreateExporter(
        HtmlConversionDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return document.ExportImages;
    }

    private static Func<OfficeImageExportFormat, HtmlRenderOptions, CancellationToken, Task<IReadOnlyList<OfficeImageExportResult>>> CreateAsyncExporter(
        HtmlConversionDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return document.ExportImagesAsync;
    }
}

public static partial class HtmlImageExportExtensions {
    /// <summary>Starts fluent image export for one rendered HTML surface.</summary>
    public static HtmlPageImageExportBuilder ToImage(this HtmlConversionDocument document) =>
        new HtmlPageImageExportBuilder(document);

    /// <summary>Starts fluent image export for one rendered HTML surface using a cloned options snapshot.</summary>
    public static HtmlPageImageExportBuilder ToImage(this HtmlConversionDocument document, HtmlRenderOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        return new HtmlPageImageExportBuilder(document, options);
    }

    /// <summary>Starts fluent image export for all rendered HTML pages.</summary>
    public static HtmlPageImageBatchExportBuilder ToImages(this HtmlConversionDocument document) =>
        new HtmlPageImageBatchExportBuilder(document);

    /// <summary>Starts fluent image export for all rendered HTML pages using a cloned options snapshot.</summary>
    public static HtmlPageImageBatchExportBuilder ToImages(this HtmlConversionDocument document, HtmlRenderOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        return new HtmlPageImageBatchExportBuilder(document, options);
    }
}
