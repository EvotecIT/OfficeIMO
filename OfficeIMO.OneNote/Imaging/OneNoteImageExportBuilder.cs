using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote;

/// <summary>Selection and rendering options for section- and notebook-level page image export.</summary>
public sealed class OneNotePageBatchRenderingOptions : OneNotePageRenderingOptions {
    /// <summary>Zero-based first page index in the flattened page sequence.</summary>
    public int PageIndex { get; set; }

    /// <summary>Maximum number of pages to export, or all remaining pages when absent.</summary>
    public int? PageCount { get; set; }

    internal OneNotePageBatchRenderingOptions CloneBatch() => new OneNotePageBatchRenderingOptions {
        Scale = Scale,
        BackgroundColor = BackgroundColor,
        RasterEncoding = RasterEncoding?.Clone() ?? new OfficeRasterEncodingOptions(),
        IncludeTitle = IncludeTitle,
        IncludeImages = IncludeImages,
        IncludeInk = IncludeInk,
        IncludeMath = IncludeMath,
        IncludeAttachmentPlaceholders = IncludeAttachmentPlaceholders,
        MaxImageBytes = MaxImageBytes,
        MaximumRasterPixels = MaximumRasterPixels,
        ImageCodec = ImageCodec,
        AutomaticPageWidthPoints = AutomaticPageWidthPoints,
        AutomaticPageHeightPoints = AutomaticPageHeightPoints,
        AutomaticPagePaddingPoints = AutomaticPagePaddingPoints,
        DefaultFont = DefaultFont,
        Ink = Ink?.Clone() ?? new OfficeInkRenderOptions(),
        Math = Math?.Clone() ?? new OfficeMathRenderOptions(),
        PageIndex = PageIndex,
        PageCount = PageCount
    };
}

/// <summary>Fluent image export for one OneNote page.</summary>
public sealed class OneNotePageImageExportBuilder : OfficeImageExportBuilder<OneNotePageImageExportBuilder, OneNotePageRenderingOptions> {
    internal OneNotePageImageExportBuilder(OneNotePage page, OneNotePageRenderingOptions? options = null)
        : base(options?.Clone() ?? new OneNotePageRenderingOptions(), (format, effective) => OneNotePageImageRenderer.Render(page, format, effective)) { }

    /// <summary>Uses a raster resolution expressed in dots per inch; 72 DPI equals scale 1.</summary>
    public OneNotePageImageExportBuilder WithDpi(double dpi) => WithScale(DpiScale(dpi));

    /// <summary>Includes or excludes the page title.</summary>
    public OneNotePageImageExportBuilder IncludeTitle(bool include = true) { Options.IncludeTitle = include; return this; }

    /// <summary>Includes or excludes embedded images and printout backgrounds.</summary>
    public OneNotePageImageExportBuilder IncludeImages(bool include = true) { Options.IncludeImages = include; return this; }

    /// <summary>Includes or excludes native ink.</summary>
    public OneNotePageImageExportBuilder IncludeInk(bool include = true) { Options.IncludeInk = include; return this; }

    /// <summary>Includes or excludes structured math typesetting.</summary>
    public OneNotePageImageExportBuilder IncludeMath(bool include = true) { Options.IncludeMath = include; return this; }

    /// <summary>Uses an application-supplied decoder for additional embedded source image formats.</summary>
    public OneNotePageImageExportBuilder WithImageCodec(IOfficeRasterImageCodec imageCodec) { Options.ImageCodec = imageCodec ?? throw new ArgumentNullException(nameof(imageCodec)); return this; }

    private static double DpiScale(double dpi) {
        if (double.IsNaN(dpi) || double.IsInfinity(dpi) || dpi <= 0D) throw new ArgumentOutOfRangeException(nameof(dpi));
        return dpi / 72D;
    }
}

/// <summary>Fluent batch image export for a OneNote section or notebook.</summary>
public sealed class OneNotePageImageBatchExportBuilder : OfficeImageExportBatchBuilder<OneNotePageImageBatchExportBuilder, OneNotePageBatchRenderingOptions> {
    internal OneNotePageImageBatchExportBuilder(OneNoteSection section, OneNotePageBatchRenderingOptions? options = null)
        : base(options?.CloneBatch() ?? new OneNotePageBatchRenderingOptions(), (format, effective) => ExportSection(section, format, effective)) { }

    internal OneNotePageImageBatchExportBuilder(OneNoteNotebook notebook, OneNotePageBatchRenderingOptions? options = null)
        : base(options?.CloneBatch() ?? new OneNotePageBatchRenderingOptions(), (format, effective) => ExportNotebook(notebook, format, effective)) { }

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

    /// <summary>Uses a raster resolution expressed in dots per inch; 72 DPI equals scale 1.</summary>
    public OneNotePageImageBatchExportBuilder WithDpi(double dpi) {
        if (double.IsNaN(dpi) || double.IsInfinity(dpi) || dpi <= 0D) throw new ArgumentOutOfRangeException(nameof(dpi));
        return WithScale(dpi / 72D);
    }

    /// <summary>Includes or excludes native ink.</summary>
    public OneNotePageImageBatchExportBuilder IncludeInk(bool include = true) { Options.IncludeInk = include; return this; }

    /// <summary>Includes or excludes structured math typesetting.</summary>
    public OneNotePageImageBatchExportBuilder IncludeMath(bool include = true) { Options.IncludeMath = include; return this; }

    /// <summary>Uses an application-supplied decoder for additional embedded source image formats.</summary>
    public OneNotePageImageBatchExportBuilder WithImageCodec(IOfficeRasterImageCodec imageCodec) { Options.ImageCodec = imageCodec ?? throw new ArgumentNullException(nameof(imageCodec)); return this; }

    private static IReadOnlyList<OfficeImageExportResult> ExportSection(OneNoteSection section, OfficeImageExportFormat format, OneNotePageBatchRenderingOptions options) {
        if (section == null) throw new ArgumentNullException(nameof(section));
        return Select(OneNotePageTraversal.Flatten(section), options).Select(item => OneNotePageImageRenderer.Render(
            item.Page, format, options, item.Page.Title, item.SectionPath + "/page[" + item.Index + "]")).ToArray();
    }

    private static IReadOnlyList<OfficeImageExportResult> ExportNotebook(OneNoteNotebook notebook, OfficeImageExportFormat format, OneNotePageBatchRenderingOptions options) {
        if (notebook == null) throw new ArgumentNullException(nameof(notebook));
        return Select(OneNotePageTraversal.Flatten(notebook), options).Select(item => OneNotePageImageRenderer.Render(
            item.Page, format, options, item.Page.Title, notebook.Name + "/" + item.SectionPath + "/page[" + item.Index + "]")).ToArray();
    }

    private static IEnumerable<OneNotePageReference> Select(IReadOnlyList<OneNotePageReference> pages, OneNotePageBatchRenderingOptions options) {
        int end = options.PageCount.HasValue
            ? (int)Math.Min((long)pages.Count, (long)options.PageIndex + options.PageCount.Value)
            : pages.Count;
        for (int index = Math.Min(options.PageIndex, pages.Count); index < end; index++) yield return pages[index];
    }
}

/// <summary>OneNote page rendering and image-export entry points.</summary>
public static class OneNoteImageExportExtensions {
    /// <summary>Starts fluent PNG/JPEG/TIFF/SVG/WebP export for a page.</summary>
    public static OneNotePageImageExportBuilder ToImage(this OneNotePage page) => new OneNotePageImageExportBuilder(page);

    /// <summary>Starts batch PNG/JPEG/TIFF/SVG/WebP export for all section pages.</summary>
    public static OneNotePageImageBatchExportBuilder ToImages(this OneNoteSection section) => new OneNotePageImageBatchExportBuilder(section);

    /// <summary>Starts batch PNG/JPEG/TIFF/SVG/WebP export for all notebook pages.</summary>
    public static OneNotePageImageBatchExportBuilder ToImages(this OneNoteNotebook notebook) => new OneNotePageImageBatchExportBuilder(notebook);

    /// <summary>Creates the reusable Drawing scene for a page.</summary>
    public static OfficeDrawing ToDrawing(this OneNotePage page, OneNotePageRenderingOptions? options = null) => OneNotePageRenderer.Render(page, options);
}
