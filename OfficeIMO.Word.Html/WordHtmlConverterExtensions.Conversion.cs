using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Html;

namespace OfficeIMO.Word.Html;

public static partial class WordHtmlConverterExtensions {
    /// <summary>Converts Word to HTML text.</summary>
    public static string ToHtml(this WordDocument document, WordHtmlExportOptions? options = null, CancellationToken cancellationToken = default) => document.ToHtml(options, out _, cancellationToken);
    /// <summary>Converts Word to HTML text and returns completed-stage findings.</summary>
    public static string ToHtml(this WordDocument document, WordHtmlExportOptions? options, out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        var result = document.ConvertToHtmlWithEvidence(options, cancellationToken);
        report = HtmlConversionEvidence.CreateReport("Word", "Html", result.Report.NativeDiagnostics, timer.Elapsed);
        if (options?.FailOnLoss == true) report.ThrowIfLoss();
        return result.RequireValue();
    }
    /// <summary>Converts Word to HTML text with its completed-stage report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static async Task<OfficeConversionResult<string>> ToHtmlAsync(this WordDocument document, WordHtmlExportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordHtmlExportOptions()).CloneForConversion();
        return await Task.Run(() => { string value = document.ToHtml(options, out var report, cancellationToken);
            return new OfficeConversionResult<string>(value, report); }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Converts prepared HTML into a Word document.</summary>
    public static WordDocument ToWord(this HtmlConversionDocument document, HtmlWordExportOptions? options = null, CancellationToken cancellationToken = default) => document.ToWord(options, out _, cancellationToken);
    /// <summary>Converts prepared HTML into Word and returns completed-stage findings.</summary>
    public static WordDocument ToWord(this HtmlConversionDocument document, HtmlWordExportOptions? options, out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        var result = document.ConvertToWordWithEvidence(options, cancellationToken);
        report = HtmlConversionEvidence.CreateReport("Html", "Word", result.Report.NativeDiagnostics, timer.Elapsed);
        try { if (options?.FailOnLoss == true) report.ThrowIfLoss(); return result.RequireValue(); }
        catch { result.Value.Dispose(); throw; }
    }
    /// <summary>Converts prepared HTML into Word with its completed-stage report.</summary>
    public static async Task<OfficeConversionResult<WordDocument>> ToWordAsync(this HtmlConversionDocument document, HtmlWordExportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new HtmlWordExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        var result = await document.ConvertToWordWithEvidenceAsync(options, cancellationToken).ConfigureAwait(false);
        var report = HtmlConversionEvidence.CreateReport("Html", "Word", result.Report.NativeDiagnostics, timer.Elapsed);
        try { if (options?.FailOnLoss == true) report.ThrowIfLoss(); return new OfficeConversionResult<WordDocument>(result.RequireValue(), report); }
        catch { result.Value.Dispose(); throw; }
    }

    /// <summary>Saves converted Html content with its completed-stage report.</summary>
    public static OfficeConversionReport SaveAsHtml(this WordDocument document, string path, WordHtmlExportOptions? options = null, CancellationToken cancellationToken = default) {
        var value = document.ToHtml(options, out var report, cancellationToken);
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => HtmlTextIO.Write(path, value));
        return new OfficeConversionReport(report.Source, report.Target, path, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }

    /// <summary>Saves converted Html content with its completed-stage report.</summary>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this WordDocument document, string path, WordHtmlExportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordHtmlExportOptions()).CloneForConversion();
        var converted = await document.ToHtmlAsync(options, cancellationToken).ConfigureAwait(false);
        var value = converted.Value;
        var report = converted.Report;
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => HtmlTextIO.WriteAsync(path, value, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport(report.Source, report.Target, path, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }

    /// <summary>Saves converted Html content with its completed-stage report.</summary>
    public static OfficeConversionReport SaveAsHtml(this WordDocument document, Stream stream, WordHtmlExportOptions? options = null, CancellationToken cancellationToken = default) {
        var value = document.ToHtml(options, out var report, cancellationToken);
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => HtmlTextIO.Write(stream, value));
        return new OfficeConversionReport(report.Source, report.Target, null, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }

    /// <summary>Saves converted Html content with its completed-stage report.</summary>
    public static async Task<OfficeConversionReport> SaveAsHtmlAsync(this WordDocument document, Stream stream, WordHtmlExportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordHtmlExportOptions()).CloneForConversion();
        var converted = await document.ToHtmlAsync(options, cancellationToken).ConfigureAwait(false);
        var value = converted.Value;
        var report = converted.Report;
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => HtmlTextIO.WriteAsync(stream, value, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport(report.Source, report.Target, null, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }

    /// <summary>Saves converted Word content with its completed-stage report.</summary>
    public static OfficeConversionReport SaveAsWord(this HtmlConversionDocument document, string path, HtmlWordExportOptions? options = null, CancellationToken cancellationToken = default) {
        using var value = document.ToWord(options, out var report, cancellationToken);
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => OfficeIMO.Core.Internal.OfficeFileCommit.WriteAllBytes(path, value.ToBytes()));
        return new OfficeConversionReport(report.Source, report.Target, path, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }

    /// <summary>Saves converted Word content with its completed-stage report.</summary>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this HtmlConversionDocument document, string path, HtmlWordExportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new HtmlWordExportOptions()).Clone();
        var converted = await document.ToWordAsync(options, cancellationToken).ConfigureAwait(false);
        using var value = converted.Value;
        var report = converted.Report;
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => OfficeIMO.Core.Internal.OfficeFileCommit.WriteAllBytesAsync(path, value.ToBytes(), cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport(report.Source, report.Target, path, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }

    /// <summary>Saves converted Word content with its completed-stage report.</summary>
    public static OfficeConversionReport SaveAsWord(this HtmlConversionDocument document, Stream stream, HtmlWordExportOptions? options = null, CancellationToken cancellationToken = default) {
        using var value = document.ToWord(options, out var report, cancellationToken);
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => OfficeIMO.Core.Internal.OfficeStreamWriter.WriteAllBytes(stream, value.ToBytes()));
        return new OfficeConversionReport(report.Source, report.Target, null, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }

    /// <summary>Saves converted Word content with its completed-stage report.</summary>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this HtmlConversionDocument document, Stream stream, HtmlWordExportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new HtmlWordExportOptions()).Clone();
        var converted = await document.ToWordAsync(options, cancellationToken).ConfigureAwait(false);
        using var value = converted.Value;
        var report = converted.Report;
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(report, () => OfficeIMO.Core.Internal.OfficeStreamWriter.WriteAllBytesAsync(stream, value.ToBytes(), cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport(report.Source, report.Target, null, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }
}
