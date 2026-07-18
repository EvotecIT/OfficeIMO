using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.OneNote;

/// <summary>Selection and rendering options for section- and notebook-level page image export.</summary>
public sealed class OneNotePageBatchRenderingOptions : OneNotePageRenderingOptions {
    /// <summary>Zero-based first page index in the flattened page sequence.</summary>
    public int PageIndex { get; set; }

    /// <summary>Maximum number of pages to export, or all remaining pages when absent.</summary>
    public int? PageCount { get; set; }

    internal OneNotePageBatchRenderingOptions CloneBatch() {
        OneNotePageBatchRenderingOptions clone = CopyTo(new OneNotePageBatchRenderingOptions());
        clone.PageIndex = PageIndex;
        clone.PageCount = PageCount;
        return clone;
    }

    internal void ValidateBatch() {
        Validate();
        if (PageIndex < 0) throw new ArgumentOutOfRangeException(nameof(PageIndex));
        if (PageCount.HasValue && PageCount.Value < 1) throw new ArgumentOutOfRangeException(nameof(PageCount));
    }
}

/// <summary>Fluent image export for one OneNote page.</summary>
public sealed class OneNotePageImageExportBuilder : OfficeImageExportBuilder<OneNotePageImageExportBuilder, OneNotePageRenderingOptions> {
    internal OneNotePageImageExportBuilder(OneNotePage page, OneNotePageRenderingOptions? options = null)
        : base(
            options?.Clone() ?? new OneNotePageRenderingOptions(),
            (format, effective, cancellationToken) =>
                OneNotePageImageRenderer.Render(page, format, effective, cancellationToken: cancellationToken)) { }

    /// <summary>Includes or excludes the page title.</summary>
    public OneNotePageImageExportBuilder IncludeTitle(bool include = true) { Options.IncludeTitle = include; return this; }

    /// <summary>Includes or excludes embedded images and printout backgrounds.</summary>
    public OneNotePageImageExportBuilder IncludeImages(bool include = true) { Options.IncludeImages = include; return this; }

    /// <summary>Includes or excludes native ink.</summary>
    public OneNotePageImageExportBuilder IncludeInk(bool include = true) { Options.IncludeInk = include; return this; }

    /// <summary>Includes or excludes structured math typesetting.</summary>
    public OneNotePageImageExportBuilder IncludeMath(bool include = true) { Options.IncludeMath = include; return this; }

}

/// <summary>Fluent batch image export for a OneNote section or notebook.</summary>
public sealed class OneNotePageImageBatchExportBuilder : OfficeImageExportBatchBuilder<OneNotePageImageBatchExportBuilder, OneNotePageBatchRenderingOptions> {
    internal OneNotePageImageBatchExportBuilder(OneNoteSection section, OneNotePageBatchRenderingOptions? options = null)
        : base(
            options?.CloneBatch() ?? new OneNotePageBatchRenderingOptions(),
            (format, effective) => OneNoteImageExportEngine.ExportSection(section, format, effective),
            (format, effective, consumer, cancellationToken) =>
                OneNoteImageExportEngine.ExportSection(section, format, effective, consumer, cancellationToken)) { }

    internal OneNotePageImageBatchExportBuilder(OneNoteNotebook notebook, OneNotePageBatchRenderingOptions? options = null)
        : base(
            options?.CloneBatch() ?? new OneNotePageBatchRenderingOptions(),
            (format, effective) => OneNoteImageExportEngine.ExportNotebook(notebook, format, effective),
            (format, effective, consumer, cancellationToken) =>
                OneNoteImageExportEngine.ExportNotebook(notebook, format, effective, consumer, cancellationToken)) { }

    /// <summary>Starts with the specified zero-based flattened page index.</summary>
    public OneNotePageImageBatchExportBuilder FromPage(int pageIndex) {
        if (pageIndex < 0) throw new ArgumentOutOfRangeException(nameof(pageIndex));
        Options.PageIndex = pageIndex;
        return this;
    }

    /// <summary>Limits output to the requested number of pages.</summary>
    public OneNotePageImageBatchExportBuilder TakePages(int pageCount) {
        if (pageCount < 1) throw new ArgumentOutOfRangeException(nameof(pageCount));
        Options.PageCount = pageCount;
        return this;
    }

    /// <summary>Exports every page from the beginning.</summary>
    public OneNotePageImageBatchExportBuilder AllPages() { Options.PageIndex = 0; Options.PageCount = null; return this; }

    /// <summary>Includes or excludes native ink.</summary>
    public OneNotePageImageBatchExportBuilder IncludeInk(bool include = true) { Options.IncludeInk = include; return this; }

    /// <summary>Includes or excludes structured math typesetting.</summary>
    public OneNotePageImageBatchExportBuilder IncludeMath(bool include = true) { Options.IncludeMath = include; return this; }

}

/// <summary>OneNote page rendering and image-export entry points.</summary>
public static class OneNoteImageExportExtensions {
    /// <summary>Exports a page using the shared five-format result contract.</summary>
    public static OfficeImageExportResult ExportImage(
        this OneNotePage page,
        OfficeImageExportFormat format,
        OneNotePageRenderingOptions? options = null,
        CancellationToken cancellationToken = default) =>
        OneNotePageImageRenderer.Render(
            page,
            format,
            options?.Clone() ?? new OneNotePageRenderingOptions(),
            cancellationToken: cancellationToken);

    /// <summary>Exports selected section pages using the shared five-format result contract.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this OneNoteSection section,
        OfficeImageExportFormat format,
        OneNotePageBatchRenderingOptions? options = null) =>
        OneNoteImageExportEngine.ExportSection(
            section,
            format,
            options?.CloneBatch() ?? new OneNotePageBatchRenderingOptions());

    /// <summary>Exports selected notebook pages using the shared five-format result contract.</summary>
    public static IReadOnlyList<OfficeImageExportResult> ExportImages(
        this OneNoteNotebook notebook,
        OfficeImageExportFormat format,
        OneNotePageBatchRenderingOptions? options = null) =>
        OneNoteImageExportEngine.ExportNotebook(
            notebook,
            format,
            options?.CloneBatch() ?? new OneNotePageBatchRenderingOptions());

    /// <summary>Starts fluent PNG/JPEG/TIFF/SVG/WebP export for a page.</summary>
    public static OneNotePageImageExportBuilder ToImage(this OneNotePage page) => new OneNotePageImageExportBuilder(page);

    /// <summary>Starts fluent export using a cloned options snapshot.</summary>
    public static OneNotePageImageExportBuilder ToImage(this OneNotePage page, OneNotePageRenderingOptions options) =>
        new OneNotePageImageExportBuilder(page, options ?? throw new ArgumentNullException(nameof(options)));

    /// <summary>Starts batch PNG/JPEG/TIFF/SVG/WebP export for all section pages.</summary>
    public static OneNotePageImageBatchExportBuilder ToImages(this OneNoteSection section) => new OneNotePageImageBatchExportBuilder(section);

    /// <summary>Starts batch export using a cloned options snapshot.</summary>
    public static OneNotePageImageBatchExportBuilder ToImages(this OneNoteSection section, OneNotePageBatchRenderingOptions options) =>
        new OneNotePageImageBatchExportBuilder(section, options ?? throw new ArgumentNullException(nameof(options)));

    /// <summary>Starts batch PNG/JPEG/TIFF/SVG/WebP export for all notebook pages.</summary>
    public static OneNotePageImageBatchExportBuilder ToImages(this OneNoteNotebook notebook) => new OneNotePageImageBatchExportBuilder(notebook);

    /// <summary>Starts batch export using a cloned options snapshot.</summary>
    public static OneNotePageImageBatchExportBuilder ToImages(this OneNoteNotebook notebook, OneNotePageBatchRenderingOptions options) =>
        new OneNotePageImageBatchExportBuilder(notebook, options ?? throw new ArgumentNullException(nameof(options)));

    /// <summary>Creates the reusable Drawing scene for a page.</summary>
    public static OfficeDrawing ToDrawing(this OneNotePage page, OneNotePageRenderingOptions? options = null) => OneNotePageRenderer.Render(page, options);
}

internal static class OneNoteImageExportEngine {
    internal static IReadOnlyList<OfficeImageExportResult> ExportSection(
        OneNoteSection section,
        OfficeImageExportFormat format,
        OneNotePageBatchRenderingOptions options) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        var results = new List<OfficeImageExportResult>();
        ExportSection(section, format, options, results.Add);
        return results.AsReadOnly();
    }

    internal static void ExportSection(
        OneNoteSection section,
        OfficeImageExportFormat format,
        OneNotePageBatchRenderingOptions options,
        OfficeImageExportConsumer consumer,
        CancellationToken cancellationToken = default) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        OneNotePageReference[] pages = Select(OneNotePageTraversal.Flatten(section), options).ToArray();
        OfficeImageExportBatchProcessor.ForEachOrdered(
            pages,
            options.MaximumDegreeOfParallelism,
            (item, _, token) => OneNotePageImageRenderer.Render(
                item.Page,
                format,
                options,
                item.Page.Title,
                item.SectionPath + "/page[" + item.Index + "]",
                token),
            consumer,
            cancellationToken,
            options);
    }

    internal static IReadOnlyList<OfficeImageExportResult> ExportNotebook(
        OneNoteNotebook notebook,
        OfficeImageExportFormat format,
        OneNotePageBatchRenderingOptions options) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        var results = new List<OfficeImageExportResult>();
        ExportNotebook(notebook, format, options, results.Add);
        return results.AsReadOnly();
    }

    internal static void ExportNotebook(
        OneNoteNotebook notebook,
        OfficeImageExportFormat format,
        OneNotePageBatchRenderingOptions options,
        OfficeImageExportConsumer consumer,
        CancellationToken cancellationToken = default) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        OneNotePageReference[] pages = Select(OneNotePageTraversal.Flatten(notebook), options).ToArray();
        OfficeImageExportBatchProcessor.ForEachOrdered(
            pages,
            options.MaximumDegreeOfParallelism,
            (item, _, token) => OneNotePageImageRenderer.Render(
                item.Page,
                format,
                options,
                item.Page.Title,
                notebook.Name + "/" + item.SectionPath + "/page[" + item.Index + "]",
                token),
            consumer,
            cancellationToken,
            options);
    }

    private static IEnumerable<OneNotePageReference> Select(
        IReadOnlyList<OneNotePageReference> pages,
        OneNotePageBatchRenderingOptions options) {
        options.ValidateBatch();
        int end = options.PageCount.HasValue
            ? (int)Math.Min((long)pages.Count, (long)options.PageIndex + options.PageCount.Value)
            : pages.Count;
        for (int index = Math.Min(options.PageIndex, pages.Count); index < end; index++) {
            yield return pages[index];
        }
    }
}
