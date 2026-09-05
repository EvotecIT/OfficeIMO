using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

/// <content>Converts the first-party logical PDF model to RTF.</content>
public static partial class RtfPdfConverterExtensions {
    /// <summary>Converts an opened PDF into an editable RTF document.</summary>
    public static RtfDocument ToRtfDocument(
        this PdfCore.PdfDocument document,
        PdfToRtfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        document.ToRtfDocumentResult(options, cancellationToken).Value;

    /// <summary>Converts an opened PDF into an editable RTF document with conversion diagnostics.</summary>
    public static PdfRtfConversionResult ToRtfDocumentResult(
        this PdfCore.PdfDocument document,
        PdfToRtfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (document == null) throw new ArgumentNullException(nameof(document));
        return ReadForRtf(document, options, cancellationToken).ToRtfDocumentResult(options, cancellationToken);
    }

    /// <summary>Converts an opened PDF and saves the editable RTF document to a file.</summary>
    public static OfficeOutputResult<PdfRtfConversionReport> SaveAsRtf(
        this PdfCore.PdfDocument document,
        string path,
        PdfToRtfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (document == null) throw new ArgumentNullException(nameof(document));
        return ReadForRtf(document, options, cancellationToken).SaveAsRtf(path, options, cancellationToken);
    }

    /// <summary>Converts an opened PDF and saves the editable RTF document to a caller-owned stream.</summary>
    public static OfficeOutputResult<PdfRtfConversionReport> SaveAsRtf(
        this PdfCore.PdfDocument document,
        Stream stream,
        PdfToRtfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (document == null) throw new ArgumentNullException(nameof(document));
        return ReadForRtf(document, options, cancellationToken).SaveAsRtf(stream, options, cancellationToken);
    }

    /// <summary>Converts an opened PDF and asynchronously saves the editable RTF document to a file.</summary>
    public static Task<OfficeOutputResult<PdfRtfConversionReport>> SaveAsRtfAsync(
        this PdfCore.PdfDocument document,
        string path,
        PdfToRtfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (document == null) throw new ArgumentNullException(nameof(document));
        return ReadForRtf(document, options, cancellationToken).SaveAsRtfAsync(path, options, cancellationToken);
    }

    /// <summary>Converts an opened PDF and asynchronously saves the editable RTF document to a caller-owned stream.</summary>
    public static Task<OfficeOutputResult<PdfRtfConversionReport>> SaveAsRtfAsync(
        this PdfCore.PdfDocument document,
        Stream stream,
        PdfToRtfOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (document == null) throw new ArgumentNullException(nameof(document));
        return ReadForRtf(document, options, cancellationToken).SaveAsRtfAsync(stream, options, cancellationToken);
    }

    private static PdfCore.PdfDocumentReadResult ReadForRtf(
        PdfCore.PdfDocument document,
        PdfToRtfOptions? options,
        CancellationToken cancellationToken = default) =>
        document.Read(options?.ReadOptions, cancellationToken);

    /// <summary>Converts a logical PDF model into an editable RTF document.</summary>
    public static RtfDocument ToRtfDocument(
        this PdfCore.PdfDocumentReadResult document,
        PdfToRtfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) =>
        document.ToRtfDocumentResult(options, cancellationToken).Value;

    /// <summary>Converts a logical PDF model into an editable RTF document with conversion diagnostics.</summary>
    public static PdfRtfConversionResult ToRtfDocumentResult(
        this PdfCore.PdfDocumentReadResult document,
        PdfToRtfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (document == null) throw new ArgumentNullException(nameof(document));
        PdfToRtfOptions operation = (options ?? new PdfToRtfOptions()).CloneForConversion();
        operation.CancellationToken = cancellationToken;
        RtfDocument value = PdfRtfConverter.Convert(document, operation);
        return new PdfRtfConversionResult(value, operation.Report);
    }

    /// <summary>Converts a logical PDF model and saves the editable RTF document to a file.</summary>
    public static OfficeOutputResult<PdfRtfConversionReport> SaveAsRtf(
        this PdfCore.PdfDocumentReadResult document,
        string path,
        PdfToRtfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Document path cannot be empty.", nameof(path));
        PdfRtfConversionResult result = document.ToRtfDocumentResult(options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        result.Value.Save(path);
        return OfficeOutputResult<PdfRtfConversionReport>.FromSuccess(path, result.Report);
    }

    /// <summary>Converts a logical PDF model and saves the editable RTF document to a caller-owned stream.</summary>
    public static OfficeOutputResult<PdfRtfConversionReport> SaveAsRtf(
        this PdfCore.PdfDocumentReadResult document,
        Stream stream,
        PdfToRtfOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(stream));
        PdfRtfConversionResult result = document.ToRtfDocumentResult(options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        result.Value.Save(stream);
        return OfficeOutputResult<PdfRtfConversionReport>.FromSuccess(null, result.Report);
    }

    /// <summary>Converts a logical PDF model and asynchronously saves the editable RTF document to a file.</summary>
    public static async Task<OfficeOutputResult<PdfRtfConversionReport>> SaveAsRtfAsync(
        this PdfCore.PdfDocumentReadResult document,
        string path,
        PdfToRtfOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Document path cannot be empty.", nameof(path));
        cancellationToken.ThrowIfCancellationRequested();
        PdfRtfConversionResult result = document.ToRtfDocumentResult(options, cancellationToken);
        await result.Value.SaveAsync(path, cancellationToken: cancellationToken).ConfigureAwait(false);
        return OfficeOutputResult<PdfRtfConversionReport>.FromSuccess(path, result.Report);
    }

    /// <summary>Converts a logical PDF model and asynchronously saves the editable RTF document to a caller-owned stream.</summary>
    public static async Task<OfficeOutputResult<PdfRtfConversionReport>> SaveAsRtfAsync(
        this PdfCore.PdfDocumentReadResult document,
        Stream stream,
        PdfToRtfOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(stream));
        cancellationToken.ThrowIfCancellationRequested();
        PdfRtfConversionResult result = document.ToRtfDocumentResult(options, cancellationToken);
        await result.Value.SaveAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
        return OfficeOutputResult<PdfRtfConversionReport>.FromSuccess(null, result.Report);
    }
}
