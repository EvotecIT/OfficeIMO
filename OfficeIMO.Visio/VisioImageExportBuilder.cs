using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

/// <summary>Fluent image export for a Visio page.</summary>
public sealed class VisioPageImageExportBuilder : OfficeImageExportBuilder<VisioPageImageExportBuilder, VisioImageExportOptions> {
    internal VisioPageImageExportBuilder(VisioPage page, VisioImageExportOptions? options = null)
        : base(options?.Clone() ?? new VisioImageExportOptions(), CreateExporter(page)) {
    }

    /// <summary>Uses a raster or SVG resolution expressed in dots per Visio inch.</summary>
    public VisioPageImageExportBuilder WithDpi(double pixelsPerInch) => WithScale(DpiScale(pixelsPerInch));

    /// <summary>Includes or excludes shape and connector text.</summary>
    public VisioPageImageExportBuilder IncludeText(bool include = true) { Options.RenderText = include; return this; }

    /// <summary>Includes or excludes built-in stencil artwork.</summary>
    public VisioPageImageExportBuilder IncludeStencilArtwork(bool include = true) { Options.RenderStencilArtwork = include; return this; }

    /// <summary>Includes or excludes connector labels.</summary>
    public VisioPageImageExportBuilder IncludeConnectorLabels(bool include = true) { Options.RenderConnectorLabels = include; return this; }

    private static Func<OfficeImageExportFormat, VisioImageExportOptions, OfficeImageExportResult> CreateExporter(VisioPage page) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        return (format, options) => page.ExportImage(format, options);
    }

    private static double DpiScale(double pixelsPerInch) {
        if (pixelsPerInch <= 0D || double.IsNaN(pixelsPerInch) || double.IsInfinity(pixelsPerInch)) {
            throw new ArgumentOutOfRangeException(nameof(pixelsPerInch));
        }
        return pixelsPerInch / 96D;
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

    /// <summary>Uses a raster or SVG resolution expressed in dots per Visio inch.</summary>
    public VisioDocumentImageExportBuilder WithDpi(double pixelsPerInch) {
        if (pixelsPerInch <= 0D || double.IsNaN(pixelsPerInch) || double.IsInfinity(pixelsPerInch)) {
            throw new ArgumentOutOfRangeException(nameof(pixelsPerInch));
        }
        return WithScale(pixelsPerInch / 96D);
    }

    /// <summary>Includes or excludes shape and connector text.</summary>
    public VisioDocumentImageExportBuilder IncludeText(bool include = true) { Options.RenderText = include; return this; }

    private static Func<OfficeImageExportFormat, VisioImageExportOptions, OfficeImageExportResult> CreateExporter(VisioDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return (format, options) => document.ExportImage(format, options);
    }
}

/// <summary>Fluent batch image export for Visio document pages.</summary>
public sealed class VisioDocumentImageBatchExportBuilder : OfficeImageExportBatchBuilder<VisioDocumentImageBatchExportBuilder, VisioImageExportOptions> {
    internal VisioDocumentImageBatchExportBuilder(VisioDocument document, VisioImageExportOptions? options = null)
        : base(options?.Clone() ?? new VisioImageExportOptions(), CreateExporter(document)) {
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

    /// <summary>Uses a raster or SVG resolution expressed in dots per Visio inch.</summary>
    public VisioDocumentImageBatchExportBuilder WithDpi(double pixelsPerInch) {
        if (pixelsPerInch <= 0D || double.IsNaN(pixelsPerInch) || double.IsInfinity(pixelsPerInch)) {
            throw new ArgumentOutOfRangeException(nameof(pixelsPerInch));
        }
        return WithScale(pixelsPerInch / 96D);
    }

    /// <summary>Includes or excludes shape and connector text.</summary>
    public VisioDocumentImageBatchExportBuilder IncludeText(bool include = true) { Options.RenderText = include; return this; }

    private static Func<OfficeImageExportFormat, VisioImageExportOptions, IReadOnlyList<OfficeImageExportResult>> CreateExporter(VisioDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return (format, options) => document.ExportImages(format, options);
    }
}

/// <summary>Fluent image-export entry points for Visio documents and pages.</summary>
public static class VisioImageExportFluentExtensions {
    /// <summary>Starts fluent image export for this Visio page.</summary>
    public static VisioPageImageExportBuilder ToImage(this VisioPage page) => new VisioPageImageExportBuilder(page);

    /// <summary>Starts fluent image export for one selected page in this Visio document.</summary>
    public static VisioDocumentImageExportBuilder ToImage(this VisioDocument document) => new VisioDocumentImageExportBuilder(document);

    /// <summary>Starts fluent batch image export for pages in this Visio document.</summary>
    public static VisioDocumentImageBatchExportBuilder ToImages(this VisioDocument document) => new VisioDocumentImageBatchExportBuilder(document);
}
