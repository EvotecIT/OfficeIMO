using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Rtf;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    /// <summary>Converts a Word document to the semantic RTF model.</summary>
    public static RtfDocument ToRtf(this WordDocument document, WordRtfExportOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToRtf(options, out _, cancellationToken);

    /// <summary>Converts a Word document to the semantic RTF model and returns the completed conversion report.</summary>
    public static RtfDocument ToRtf(this WordDocument document, WordRtfExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        RtfConversionResult<RtfDocument> result = document.ConvertToRtfWithEvidence();
        RtfDocument value = result.Value;
        OfficeConversionReport source = result.Report;
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts a Word document to the semantic RTF model, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<RtfDocument>> ToRtfAsync(this WordDocument document, WordRtfExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (document.ToRtf(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsRtf(this WordDocument document, string path, WordRtfExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        RtfDocument value = document.ToRtf(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, value.ToBytes(options?.SaveOptions, options?.Encoding));
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsRtf(this WordDocument document, Stream stream, WordRtfExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        RtfDocument value = document.ToRtf(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, value.ToBytes(options?.SaveOptions, options?.Encoding));
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsRtfAsync(this WordDocument document, string path, WordRtfExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            RtfDocument value = document.ToRtf(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(options?.SaveOptions, options?.Encoding), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsRtfAsync(this WordDocument document, Stream stream, WordRtfExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            RtfDocument value = document.ToRtf(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(options?.SaveOptions, options?.Encoding), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Converts a semantic RTF document to a Word document.</summary>
    public static WordDocument ToWord(this RtfDocument document, RtfWordExportOptions? options = null, CancellationToken cancellationToken = default) =>
        document.ToWord(options, out _, cancellationToken);

    /// <summary>Converts a semantic RTF document to a Word document and returns the completed conversion report.</summary>
    public static WordDocument ToWord(this RtfDocument document, RtfWordExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        RtfConversionResult<WordDocument> result = document.ConvertToWordWithEvidence();
        WordDocument value = result.Value;
        OfficeConversionReport source = result.Report;
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts a semantic RTF document to a Word document, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<WordDocument>> ToWordAsync(this RtfDocument document, RtfWordExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (document.ToWord(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsWord(this RtfDocument document, string path, RtfWordExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using WordDocument value = document.ToWord(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, value.ToBytes());
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsWord(this RtfDocument document, Stream stream, RtfWordExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using WordDocument value = document.ToWord(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, value.ToBytes());
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this RtfDocument document, string path, RtfWordExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using WordDocument value = document.ToWord(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this RtfDocument document, Stream stream, RtfWordExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using WordDocument value = document.ToWord(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Converts a parsed RTF read result, including its read diagnostics, to a Word document.</summary>
    public static WordDocument ToWord(this RtfReadResult readResult, RtfWordExportOptions? options = null, CancellationToken cancellationToken = default) =>
        readResult.ToWord(options, out _, cancellationToken);

    /// <summary>Converts a parsed RTF read result, including its read diagnostics, to a Word document and returns the completed conversion report.</summary>
    public static WordDocument ToWord(this RtfReadResult readResult, RtfWordExportOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (readResult == null) throw new ArgumentNullException(nameof(readResult));
        cancellationToken.ThrowIfCancellationRequested();
        RtfConversionResult<WordDocument> result = readResult.ConvertToWordWithEvidence(options?.SourcePath);
        WordDocument value = result.Value;
        OfficeConversionReport source = result.Report;
        return OfficeConversionOutput.Complete(value, source, options?.FailOnLoss == true, out report);
    }

    /// <summary>Converts a parsed RTF read result, including its read diagnostics, to a Word document, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<WordDocument>> ToWordAsync(this RtfReadResult readResult, RtfWordExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (readResult.ToWord(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsWord(this RtfReadResult readResult, string path, RtfWordExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using WordDocument value = readResult.ToWord(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, value.ToBytes());
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsWord(this RtfReadResult readResult, Stream stream, RtfWordExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        using WordDocument value = readResult.ToWord(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, value.ToBytes());
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this RtfReadResult readResult, string path, RtfWordExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using WordDocument value = readResult.ToWord(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this RtfReadResult readResult, Stream stream, RtfWordExportOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using WordDocument value = readResult.ToWord(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }
}
