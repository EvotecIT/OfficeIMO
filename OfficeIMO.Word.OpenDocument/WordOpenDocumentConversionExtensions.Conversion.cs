using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.OpenDocument;
using OfficeIMO.Word;

namespace OfficeIMO.Word.OpenDocument;

public static partial class WordOpenDocumentConversionExtensions {
    /// <summary>Converts a Word document to an in-memory ODT text document; the LossPolicy option rejects loss.</summary>
    public static OdtDocument ToOpenDocument(this WordDocument source, WordOpenDocumentConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        source.ToOpenDocument(options, out _, cancellationToken);

    /// <summary>Converts a Word document to an in-memory ODT text document; the LossPolicy option rejects loss and returns the completed conversion report.</summary>
    public static OdtDocument ToOpenDocument(this WordDocument source, WordOpenDocumentConversionOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        OdfConversionResult<OdtDocument> result = source.ConvertToOpenDocumentWithEvidence(options);
        OdtDocument value = result.Value;
        OfficeConversionReport evidence = result.Report;
        return OfficeConversionOutput.Complete(value, evidence, false, out report);
    }

    /// <summary>Converts a Word document to an in-memory ODT text document; the LossPolicy option rejects loss, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<OdtDocument>> ToOpenDocumentAsync(this WordDocument source, WordOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (source.ToOpenDocument(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsOpenDocument(this WordDocument source, string path, WordOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        using OdtDocument value = source.ToOpenDocument(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, value.ToBytes());
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsOpenDocument(this WordDocument source, Stream stream, WordOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        using OdtDocument value = source.ToOpenDocument(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, value.ToBytes());
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsOpenDocumentAsync(this WordDocument source, string path, WordOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using OdtDocument value = source.ToOpenDocument(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsOpenDocumentAsync(this WordDocument source, Stream stream, WordOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using OdtDocument value = source.ToOpenDocument(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Converts an ODT text document to an in-memory Word document; the LossPolicy option rejects loss.</summary>
    public static WordDocument ToWord(this OdtDocument source, WordOpenDocumentConversionOptions? options = null, CancellationToken cancellationToken = default) =>
        source.ToWord(options, out _, cancellationToken);

    /// <summary>Converts an ODT text document to an in-memory Word document; the LossPolicy option rejects loss and returns the completed conversion report.</summary>
    public static WordDocument ToWord(this OdtDocument source, WordOpenDocumentConversionOptions? options, out OfficeConversionReport report,
        CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        OdfConversionResult<WordDocument> result = source.ConvertToWordWithEvidence(options);
        WordDocument value = result.Value;
        OfficeConversionReport evidence = result.Report;
        return OfficeConversionOutput.Complete(value, evidence, false, out report);
    }

    /// <summary>Converts an ODT text document to an in-memory Word document; the LossPolicy option rejects loss, returning the value and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static Task<OfficeConversionResult<WordDocument>> ToWordAsync(this OdtDocument source, WordOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        return OfficeConversionOutput.RunAsync(() => (source.ToWord(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);
    }

    /// <summary>Writes the converted output to a path and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsWord(this OdtDocument source, string path, WordOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        using WordDocument value = source.ToWord(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, path, value.ToBytes());
    }

    /// <summary>Writes the converted output to a caller-owned stream, which stays open and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsWord(this OdtDocument source, Stream stream, WordOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        using WordDocument value = source.ToWord(options, out OfficeConversionReport report, cancellationToken);
        return OfficeConversionOutput.Save(report, stream, value.ToBytes());
    }

    /// <summary>Asynchronously writes the converted output to a path and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this OdtDocument source, string path, WordOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using WordDocument value = source.ToWord(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, path, bytes, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the converted output to a caller-owned stream and returns the conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool; the output is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsWordAsync(this OdtDocument source, Stream stream, WordOpenDocumentConversionOptions? options = null,
        CancellationToken cancellationToken = default) {
        options = options?.Clone();
        var (bytes, report) = await Task.Run(() => {
            using WordDocument value = source.ToWord(options, out OfficeConversionReport converted, cancellationToken);
            return (value.ToBytes(), converted);
        }, cancellationToken).ConfigureAwait(false);
        return await OfficeConversionOutput.SaveAsync(report, stream, bytes, cancellationToken).ConfigureAwait(false);
    }
}
