using OfficeIMO.Drawing;
using OfficeIMO.Email;

namespace OfficeIMO.Html;

/// <summary>Fluent image export for one rendered email surface.</summary>
public sealed class EmailImageExportBuilder :
    OfficeImageExportBuilder<EmailImageExportBuilder, EmailImageExportOptions> {
    private readonly PageSelection _selection;

    internal EmailImageExportBuilder(
        EmailDocument source,
        EmailImageExportOptions? options)
        : this(
            source,
            options?.CloneEmail() ?? new EmailImageExportOptions(),
            new PageSelection()) {
    }

    private EmailImageExportBuilder(
        EmailDocument source,
        EmailImageExportOptions options,
        PageSelection selection)
        : base(
            options,
            (format, effective) => source.ExportImage(
                format,
                effective,
                selection.PageIndex),
            (format, effective, cancellationToken) => source.ExportImageAsync(
                format,
                effective,
                selection.PageIndex,
                cancellationToken)) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        _selection = selection;
    }

    /// <summary>Uses one continuous image surface for the message.</summary>
    public EmailImageExportBuilder Continuous() {
        Options.Mode = HtmlRenderMode.Continuous;
        _selection.PageIndex = 0;
        return this;
    }

    /// <summary>Uses paged layout and selects a zero-based rendered page.</summary>
    public EmailImageExportBuilder Paged(int pageIndex = 0) {
        if (pageIndex < 0) throw new ArgumentOutOfRangeException(nameof(pageIndex));
        Options.Mode = HtmlRenderMode.Paged;
        _selection.PageIndex = pageIndex;
        return this;
    }

    private sealed class PageSelection {
        internal int PageIndex { get; set; }
    }
}

/// <summary>Fluent batch image export for every rendered email page.</summary>
public sealed class EmailImageExportBatchBuilder :
    OfficeImageExportBatchBuilder<EmailImageExportBatchBuilder, EmailImageExportOptions> {
    internal EmailImageExportBatchBuilder(
        EmailDocument source,
        EmailImageExportOptions? options)
        : base(
            options?.CloneEmail() ?? new EmailImageExportOptions(),
            (format, effective) => source.ExportImages(format, effective),
            (format, effective, consumer, cancellationToken) =>
                source.ExportImages(
                    format,
                    consumer,
                    effective,
                    cancellationToken),
            (format, effective, consumer, cancellationToken) =>
                source.ExportImagesAsync(
                    format,
                    consumer,
                    effective,
                    cancellationToken)) {
        if (source == null) throw new ArgumentNullException(nameof(source));
    }

    /// <summary>Uses one continuous image surface for the message.</summary>
    public EmailImageExportBatchBuilder Continuous() {
        Options.Mode = HtmlRenderMode.Continuous;
        return this;
    }

    /// <summary>Uses paged layout and exports every rendered page.</summary>
    public EmailImageExportBatchBuilder Paged() {
        Options.Mode = HtmlRenderMode.Paged;
        return this;
    }
}

/// <summary>Fluent email image-export entry points.</summary>
public static class EmailImageExportBuilderExtensions {
    /// <summary>Starts fluent image export for one rendered email surface.</summary>
    public static EmailImageExportBuilder ToImage(
        this EmailDocument source,
        EmailImageExportOptions? options = null) =>
        new EmailImageExportBuilder(source, options);

    /// <summary>Starts fluent image export for every rendered email page.</summary>
    public static EmailImageExportBatchBuilder ToImages(
        this EmailDocument source,
        EmailImageExportOptions? options = null) =>
        new EmailImageExportBatchBuilder(source, options);
}
