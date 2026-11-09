using OfficeIMO.Core.Internal;

namespace OfficeIMO.AsciiDoc.Markdown;

/// <summary>Loss-aware conversions between AsciiDoc and Markdown models.</summary>
public static class AsciiDocMarkdownConverterExtensions {
    /// <summary>Converts into a detached Markdown model.</summary>
    public static MarkdownDocument ToMarkdown(this AsciiDocDocument document, AsciiDocMarkdownExportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToMarkdown(options, out _, cancellationToken);

    /// <summary>Converts and retains native diagnostic locations and canonical source evidence.</summary>
    public static MarkdownDocument ToMarkdown(this AsciiDocDocument document, AsciiDocMarkdownExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var result = AsciiDocToMarkdownConverter.Convert(document, options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return OfficeConversionOutput.Complete(result.Value, result.Report, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts asynchronously and returns the model and completed report.</summary>
    public static Task<OfficeConversionResult<MarkdownDocument>> ToMarkdownAsync(this AsciiDocDocument document,
        AsciiDocMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeConversionOutput.RunAsync(() => (document.ToMarkdown(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);

    /// <summary>Writes converted Markdown and returns the report, including native evidence.</summary>
    public static OfficeConversionReport SaveAsMarkdown(this AsciiDocDocument document,
        string path, AsciiDocMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        MarkdownDocument value = document.ToMarkdown(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return OfficeConversionOutput.Save(report, path, OfficeConversionOutput.Text(value.ToMarkdown()));
    }

    /// <summary>Writes converted Markdown and returns the report, including native evidence.</summary>
    public static OfficeConversionReport SaveAsMarkdown(this AsciiDocDocument document,
        Stream stream, AsciiDocMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        MarkdownDocument value = document.ToMarkdown(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeConversionOutput.Save(report, stream, OfficeConversionOutput.Text(value.ToMarkdown()));
        return report;
    }

    /// <summary>Asynchronously writes converted Markdown and returns the report, including native evidence.</summary>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this AsciiDocDocument document,
        string path, AsciiDocMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        MarkdownDocument value = document.ToMarkdown(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return await OfficeConversionOutput.SaveAsync(report, path, OfficeConversionOutput.Text(value.ToMarkdown()), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes converted Markdown and returns the report, including native evidence.</summary>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this AsciiDocDocument document,
        Stream stream, AsciiDocMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        MarkdownDocument value = document.ToMarkdown(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeConversionOutput.SaveAsync(report, stream, OfficeConversionOutput.Text(value.ToMarkdown()), cancellationToken).ConfigureAwait(false);
        return report;
    }

    /// <summary>Converts into a detached AsciiDoc model.</summary>
    public static AsciiDocDocument ToAsciiDoc(this MarkdownDocument document, MarkdownAsciiDocExportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToAsciiDoc(options, out _, cancellationToken);

    /// <summary>Converts and retains native diagnostic locations and canonical source evidence.</summary>
    public static AsciiDocDocument ToAsciiDoc(this MarkdownDocument document, MarkdownAsciiDocExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var result = MarkdownToAsciiDocConverter.Convert(document, options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return OfficeConversionOutput.Complete(result.Value, result.Report, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts asynchronously and returns the model and completed report.</summary>
    public static Task<OfficeConversionResult<AsciiDocDocument>> ToAsciiDocAsync(this MarkdownDocument document,
        MarkdownAsciiDocExportOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeConversionOutput.RunAsync(() => (document.ToAsciiDoc(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);

    /// <summary>Writes converted AsciiDoc and returns the report, including native evidence.</summary>
    public static OfficeConversionReport SaveAsAsciiDoc(this MarkdownDocument document,
        string path, MarkdownAsciiDocExportOptions? options = null, CancellationToken cancellationToken = default) {
        AsciiDocDocument value = document.ToAsciiDoc(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return OfficeConversionOutput.Save(report, path, OfficeConversionOutput.Text(((MarkdownToAsciiDocReport)report).GeneratedSource));
    }

    /// <summary>Writes converted AsciiDoc and returns the report, including native evidence.</summary>
    public static OfficeConversionReport SaveAsAsciiDoc(this MarkdownDocument document,
        Stream stream, MarkdownAsciiDocExportOptions? options = null, CancellationToken cancellationToken = default) {
        AsciiDocDocument value = document.ToAsciiDoc(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeConversionOutput.Save(report, stream, OfficeConversionOutput.Text(((MarkdownToAsciiDocReport)report).GeneratedSource));
        return report;
    }

    /// <summary>Asynchronously writes converted AsciiDoc and returns the report, including native evidence.</summary>
    public static async Task<OfficeConversionReport> SaveAsAsciiDocAsync(this MarkdownDocument document,
        string path, MarkdownAsciiDocExportOptions? options = null, CancellationToken cancellationToken = default) {
        AsciiDocDocument value = document.ToAsciiDoc(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return await OfficeConversionOutput.SaveAsync(report, path, OfficeConversionOutput.Text(((MarkdownToAsciiDocReport)report).GeneratedSource), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes converted AsciiDoc and returns the report, including native evidence.</summary>
    public static async Task<OfficeConversionReport> SaveAsAsciiDocAsync(this MarkdownDocument document,
        Stream stream, MarkdownAsciiDocExportOptions? options = null, CancellationToken cancellationToken = default) {
        AsciiDocDocument value = document.ToAsciiDoc(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeConversionOutput.SaveAsync(report, stream, OfficeConversionOutput.Text(((MarkdownToAsciiDocReport)report).GeneratedSource), cancellationToken).ConfigureAwait(false);
        return report;
    }

    /// <summary>Converts a single block using the attributes and reference catalog in effect at that block.</summary>
    public static MarkdownDocument ToMarkdown(this AsciiDocBlock block, AsciiDocDocumentAttributes attributes,
        AsciiDocMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) =>
        block.ToMarkdown(attributes, options, out _, cancellationToken);

    /// <summary>Converts a contextual block while retaining its diagnostics and used footnotes.</summary>
    public static MarkdownDocument ToMarkdown(this AsciiDocBlock block, AsciiDocDocumentAttributes attributes,
        AsciiDocMarkdownExportOptions? options, out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var result = AsciiDocToMarkdownConverter.ConvertBlock(block, attributes, options);
        cancellationToken.ThrowIfCancellationRequested();
        return OfficeConversionOutput.Complete(result.Value, result.Report, options?.FailOnLoss == true, out report);
    }
}
