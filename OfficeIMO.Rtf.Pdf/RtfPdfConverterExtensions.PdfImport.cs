using PdfCore = OfficeIMO.Pdf;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Rtf.Pdf;

/// <content>Converts the first-party logical PDF model to editable RTF with loss evidence.</content>
public static partial class RtfPdfConverterExtensions {
    /// <summary>Reconstructs editable RTF from the supported PDF semantics.</summary>
    public static RtfDocument ToRtf(this PdfCore.PdfDocument document, PdfRtfExportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToRtf(options, out _, cancellationToken);

    /// <summary>Reconstructs RTF and retains completed semantic reconstruction diagnostics.</summary>
    public static RtfDocument ToRtf(this PdfCore.PdfDocument document, PdfRtfExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(document);
        return document.Read(options?.ReadOptions, cancellationToken).ToRtf(options, out report, cancellationToken);
    }

    /// <summary>Reconstructs RTF asynchronously and returns the document with its report.</summary>
    public static Task<OfficeConversionResult<RtfDocument>> ToRtfAsync(this PdfCore.PdfDocument document,
        PdfRtfExportOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeConversionOutput.RunAsync(() => (document.ToRtf(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);

    /// <summary>Writes reconstructed RTF and returns its completed conversion report.</summary>
    public static OfficeConversionReport SaveAsRtf(this PdfCore.PdfDocument document,
        string path, PdfRtfExportOptions? options = null, CancellationToken cancellationToken = default) {
        using RtfDocument value = document.ToRtf(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeConversionFailures.Run(report, () => value.Save(path));
        return OfficeConversionOutput.WithPath(report, Path.GetFullPath(path));
    }

    /// <summary>Writes reconstructed RTF and returns its completed conversion report.</summary>
    public static OfficeConversionReport SaveAsRtf(this PdfCore.PdfDocument document,
        Stream stream, PdfRtfExportOptions? options = null, CancellationToken cancellationToken = default) {
        using RtfDocument value = document.ToRtf(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeConversionFailures.Run(report, () => value.Save(stream));
        return report;
    }

    /// <summary>Asynchronously writes reconstructed RTF and returns its completed conversion report.</summary>
    public static async Task<OfficeConversionReport> SaveAsRtfAsync(this PdfCore.PdfDocument document,
        string path, PdfRtfExportOptions? options = null, CancellationToken cancellationToken = default) {
        using RtfDocument value = document.ToRtf(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(path, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return OfficeConversionOutput.WithPath(report, Path.GetFullPath(path));
    }

    /// <summary>Asynchronously writes reconstructed RTF and returns its completed conversion report.</summary>
    public static async Task<OfficeConversionReport> SaveAsRtfAsync(this PdfCore.PdfDocument document,
        Stream stream, PdfRtfExportOptions? options = null, CancellationToken cancellationToken = default) {
        using RtfDocument value = document.ToRtf(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(stream, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return report;
    }

    /// <summary>Reconstructs editable RTF from the supported PDF semantics.</summary>
    public static RtfDocument ToRtf(this PdfCore.PdfDocumentReadResult document, PdfRtfExportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToRtf(options, out _, cancellationToken);

    /// <summary>Reconstructs RTF and retains completed semantic reconstruction diagnostics.</summary>
    public static RtfDocument ToRtf(this PdfCore.PdfDocumentReadResult document, PdfRtfExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(document);
        PdfRtfExportOptions operation = (options ?? new PdfRtfExportOptions()).CloneForConversion();
        operation.CancellationToken = cancellationToken;
        RtfDocument value = PdfRtfConverter.Convert(document, operation);
        cancellationToken.ThrowIfCancellationRequested();
        return OfficeConversionOutput.Complete(value, new PdfRtfConversionReport(operation.Report), operation.FailOnLoss, out report);
    }

    /// <summary>Reconstructs RTF asynchronously and returns the document with its report.</summary>
    public static Task<OfficeConversionResult<RtfDocument>> ToRtfAsync(this PdfCore.PdfDocumentReadResult document,
        PdfRtfExportOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeConversionOutput.RunAsync(() => (document.ToRtf(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);

    /// <summary>Writes reconstructed RTF and returns its completed conversion report.</summary>
    public static OfficeConversionReport SaveAsRtf(this PdfCore.PdfDocumentReadResult document,
        string path, PdfRtfExportOptions? options = null, CancellationToken cancellationToken = default) {
        using RtfDocument value = document.ToRtf(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeConversionFailures.Run(report, () => value.Save(path));
        return OfficeConversionOutput.WithPath(report, Path.GetFullPath(path));
    }

    /// <summary>Writes reconstructed RTF and returns its completed conversion report.</summary>
    public static OfficeConversionReport SaveAsRtf(this PdfCore.PdfDocumentReadResult document,
        Stream stream, PdfRtfExportOptions? options = null, CancellationToken cancellationToken = default) {
        using RtfDocument value = document.ToRtf(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeConversionFailures.Run(report, () => value.Save(stream));
        return report;
    }

    /// <summary>Asynchronously writes reconstructed RTF and returns its completed conversion report.</summary>
    public static async Task<OfficeConversionReport> SaveAsRtfAsync(this PdfCore.PdfDocumentReadResult document,
        string path, PdfRtfExportOptions? options = null, CancellationToken cancellationToken = default) {
        using RtfDocument value = document.ToRtf(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(path, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return OfficeConversionOutput.WithPath(report, Path.GetFullPath(path));
    }

    /// <summary>Asynchronously writes reconstructed RTF and returns its completed conversion report.</summary>
    public static async Task<OfficeConversionReport> SaveAsRtfAsync(this PdfCore.PdfDocumentReadResult document,
        Stream stream, PdfRtfExportOptions? options = null, CancellationToken cancellationToken = default) {
        using RtfDocument value = document.ToRtf(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(stream, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return report;
    }

}
