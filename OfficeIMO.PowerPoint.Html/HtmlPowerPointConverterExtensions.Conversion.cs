using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.Html;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class HtmlPowerPointConverterExtensions {
    /// <summary>Converts prepared HTML into a native presentation.</summary>
    public static PptCore.PowerPointDocument ToPowerPoint(this HtmlConversionDocument document, HtmlPowerPointExportOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToPowerPoint(options, out _, cancellationToken);

    /// <summary>Converts prepared HTML into a native presentation and returns the completed conversion report.</summary>
    public static PptCore.PowerPointDocument ToPowerPoint(this HtmlConversionDocument document, HtmlPowerPointExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlToPowerPointResult result = document.ConvertToPowerPointWithEvidence(options);
        PptCore.PowerPointDocument value = GetPresentationOrThrow(result);
        OfficeConversionReport source = HtmlConversionEvidence.CreateReport("Html", "PowerPoint", result.Report.NativeDiagnostics);
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts prepared HTML into a native presentation, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<PptCore.PowerPointDocument>> ToPowerPointAsync(this HtmlConversionDocument document, HtmlPowerPointExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (document.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsPowerPoint(this HtmlConversionDocument document, string path, HtmlPowerPointExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using PptCore.PowerPointDocument value = document.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, value.ToBytes());
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsPowerPoint(this HtmlConversionDocument document, Stream stream, HtmlPowerPointExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using PptCore.PowerPointDocument value = document.ToPowerPoint(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, value.ToBytes());
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsPowerPointAsync(this HtmlConversionDocument document, string path, HtmlPowerPointExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using PptCore.PowerPointDocument value = document.ToPowerPoint(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsPowerPointAsync(this HtmlConversionDocument document, Stream stream, HtmlPowerPointExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using PptCore.PowerPointDocument value = document.ToPowerPoint(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }
}
