using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown;

public static partial class WordMarkdownConverterExtensions {
    private static OfficeConversionReport CreateReport(WordMarkdownConversionReport evidence, string source, string target, TimeSpan elapsed, string? path = null) =>
        new OfficeConversionReport(source, target, path, elapsed: elapsed,
            diagnostics: evidence.Diagnostics.Select(d => new OfficeDiagnostic(d.Code, d.Severity, d.Message, d.Location,
                format: d.Format ?? source, lossKind: d.LossKind)));

    /// <summary>Converts Word to a Markdown document.</summary>
    public static MarkdownDocument ToMarkdown(this WordDocument document, WordMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToMarkdown(options, out _, cancellationToken);

    /// <summary>Converts Word to Markdown and returns completed-stage findings.</summary>
    public static MarkdownDocument ToMarkdown(this WordDocument document, WordMarkdownExportOptions? options, out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        using var assets = new OfficeConversionAssets(cancellationToken);
        var timer = Stopwatch.StartNew();
        var result = ConvertToMarkdownWithEvidence(document, options, cancellationToken);
        try {
            cancellationToken.ThrowIfCancellationRequested();
            report = CreateReport(result.Report, "Word", "Markdown", timer.Elapsed);
            if (options?.FailOnLoss == true) report.ThrowIfLoss();
            OfficeConversionFailures.Run(report, assets.Complete);
            return result.RequireValue();
        } catch { result.Value.Dispose(); throw; }

    }

    /// <summary>Converts Word to Markdown with its completed-stage report.</summary>
    public static async Task<OfficeConversionResult<MarkdownDocument>> ToMarkdownAsync(this WordDocument document, WordMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordMarkdownExportOptions()).CloneForConversion();
        return await Task.Run(() => { var value = document.ToMarkdown(options, out var report, cancellationToken); try { cancellationToken.ThrowIfCancellationRequested(); }
            catch { value.Dispose(); throw; }
            return new OfficeConversionResult<MarkdownDocument>(value, report); }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves converted Markdown content with a completed-stage report.</summary>
    public static OfficeConversionReport SaveAsMarkdown(this WordDocument document, string path, WordMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        using var assets = new OfficeConversionAssets(cancellationToken);
        options = CreateFileSaveOptions(path, options);
        using var value = document.ToMarkdown(options, out var report, cancellationToken);
        byte[] bytes = value.ToBytes();
        OfficeConversionFailures.Run(report, () => {
            OfficeFileCommit.Write(path, output => OfficeStreamWriter.CopyBytes(output, bytes, cancellationToken), cancellationToken);
            assets.Complete();
        });
        return new OfficeConversionReport(report.Source, report.Target, path, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }

    /// <summary>Saves converted Markdown content with a completed-stage report.</summary>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this WordDocument document, string path, WordMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        using var assets = new OfficeConversionAssets(cancellationToken);
        options = CreateFileSaveOptions(path, options);
        using var value = document.ToMarkdown(options, out var report, cancellationToken);
        byte[] bytes = value.ToBytes();
        await OfficeConversionFailures.RunAsync(report, async () => {
            await OfficeFileCommit.WriteAllBytesAsync(path, bytes, cancellationToken: cancellationToken).ConfigureAwait(false);
            assets.Complete();
        }).ConfigureAwait(false);
        return new OfficeConversionReport(report.Source, report.Target, path, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }

    /// <summary>Saves converted Markdown content with a completed-stage report.</summary>
    public static OfficeConversionReport SaveAsMarkdown(this WordDocument document, Stream stream, WordMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        using var assets = new OfficeConversionAssets(cancellationToken);
        options = (options ?? new WordMarkdownExportOptions()).CloneForConversion();

        using var value = document.ToMarkdown(options, out var report, cancellationToken);
        byte[] bytes = value.ToBytes();
        OfficeConversionFailures.Run(report, () => {
            OfficeStreamWriter.Write(stream, output => OfficeStreamWriter.CopyBytes(output, bytes, cancellationToken));
            assets.Complete();
        });
        return new OfficeConversionReport(report.Source, report.Target, null, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }

    /// <summary>Saves converted Markdown content with a completed-stage report.</summary>
    public static async Task<OfficeConversionReport> SaveAsMarkdownAsync(this WordDocument document, Stream stream, WordMarkdownExportOptions? options = null, CancellationToken cancellationToken = default) {
        using var assets = new OfficeConversionAssets(cancellationToken);
        options = (options ?? new WordMarkdownExportOptions()).CloneForConversion();

        using var value = document.ToMarkdown(options, out var report, cancellationToken);
        byte[] bytes = value.ToBytes();
        await OfficeConversionFailures.RunAsync(report, async () => {
            await OfficeStreamWriter.WriteAllBytesAsync(stream, bytes, cancellationToken).ConfigureAwait(false);
            assets.Complete();
        }).ConfigureAwait(false);
        return new OfficeConversionReport(report.Source, report.Target, null, report.ItemsWritten, report.Elapsed, report.Diagnostics);
    }

    /// <summary>Converts Markdown to a Word document.</summary>
    public static WordDocument ToWord(this MarkdownDocument document, WordMarkdownImportOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToWord(options, out _, cancellationToken);

    /// <summary>Converts Markdown to Word and returns completed-stage findings.</summary>
    public static WordDocument ToWord(this MarkdownDocument document, WordMarkdownImportOptions? options, out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        var result = ConvertToWordWithEvidence(document, options, cancellationToken);
        report = CreateReport(result.Report, "Markdown", "Word", timer.Elapsed);
        if (options?.FailOnLoss == true && report.HasLoss) {
            if (options.Template == null) result.Value.Dispose();
            report.ThrowIfLoss();
        }
        return result.RequireValue();
    }

    /// <summary>Converts Markdown to Word with its completed-stage report.</summary>
    public static async Task<OfficeConversionResult<WordDocument>> ToWordAsync(this MarkdownDocument document, WordMarkdownImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordMarkdownImportOptions()).CloneForConversion();
        return await Task.Run(() => { var value = document.ToWord(options, out var report, cancellationToken); try { cancellationToken.ThrowIfCancellationRequested(); }
            catch { if (options.Template == null) value.Dispose(); throw; }
            return new OfficeConversionResult<WordDocument>(value, report); }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Saves converted Word content with a completed-stage report.</summary>
    public static OfficeConversionReport SaveAsWord(this MarkdownDocument document, string path, WordMarkdownImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordMarkdownImportOptions()).CloneForConversion();

        var value = document.ToWord(options, out var report, cancellationToken);
        try {
            byte[] bytes = value.ToBytes();
            OfficeConversionFailures.Run(report, () => OfficeFileCommit.Write(path, output => OfficeStreamWriter.CopyBytes(output, bytes, cancellationToken), cancellationToken));
            return new OfficeConversionReport(report.Source, report.Target, path, report.ItemsWritten, report.Elapsed, report.Diagnostics);
        } finally { if (options?.Template == null) value.Dispose(); }
    }

    /// <summary>Saves converted Word content with a completed-stage report.</summary>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this MarkdownDocument document, string path, WordMarkdownImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordMarkdownImportOptions()).CloneForConversion();

        var value = document.ToWord(options, out var report, cancellationToken);
        try {
            byte[] bytes = value.ToBytes();
            await OfficeConversionFailures.RunAsync(report, () => OfficeFileCommit.WriteAllBytesAsync(path, bytes, cancellationToken: cancellationToken)).ConfigureAwait(false);
            return new OfficeConversionReport(report.Source, report.Target, path, report.ItemsWritten, report.Elapsed, report.Diagnostics);
        } finally { if (options?.Template == null) value.Dispose(); }
    }

    /// <summary>Saves converted Word content with a completed-stage report.</summary>
    public static OfficeConversionReport SaveAsWord(this MarkdownDocument document, Stream stream, WordMarkdownImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordMarkdownImportOptions()).CloneForConversion();

        var value = document.ToWord(options, out var report, cancellationToken);
        try {
            byte[] bytes = value.ToBytes();
            OfficeConversionFailures.Run(report, () => OfficeStreamWriter.Write(stream, output => OfficeStreamWriter.CopyBytes(output, bytes, cancellationToken)));
            return new OfficeConversionReport(report.Source, report.Target, null, report.ItemsWritten, report.Elapsed, report.Diagnostics);
        } finally { if (options?.Template == null) value.Dispose(); }
    }

    /// <summary>Saves converted Word content with a completed-stage report.</summary>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this MarkdownDocument document, Stream stream, WordMarkdownImportOptions? options = null, CancellationToken cancellationToken = default) {
        options = (options ?? new WordMarkdownImportOptions()).CloneForConversion();

        var value = document.ToWord(options, out var report, cancellationToken);
        try {
            byte[] bytes = value.ToBytes();
            await OfficeConversionFailures.RunAsync(report, () => OfficeStreamWriter.WriteAllBytesAsync(stream, bytes, cancellationToken)).ConfigureAwait(false);
            return new OfficeConversionReport(report.Source, report.Target, null, report.ItemsWritten, report.Elapsed, report.Diagnostics);
        } finally { if (options?.Template == null) value.Dispose(); }
    }
}
