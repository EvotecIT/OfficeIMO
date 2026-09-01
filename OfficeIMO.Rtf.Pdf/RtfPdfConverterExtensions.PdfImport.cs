using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

/// <content>Converts the first-party logical PDF model to RTF.</content>
public static partial class RtfPdfConverterExtensions {
    /// <summary>Converts an opened PDF into an editable RTF document.</summary>
    public static RtfDocument ToRtfDocument(
        this PdfCore.PdfDocument document,
        PdfRtfImportOptions? options = null) =>
        document.ToRtfDocumentResult(options).Value;

    /// <summary>Converts an opened PDF into an editable RTF document with conversion diagnostics.</summary>
    public static PdfRtfConversionResult ToRtfDocumentResult(
        this PdfCore.PdfDocument document,
        PdfRtfImportOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return ReadForRtf(document, options).ToRtfDocumentResult(options);
    }

    /// <summary>Converts an opened PDF and saves the editable RTF document to a file.</summary>
    public static PdfRtfConversionReport SaveAsRtf(
        this PdfCore.PdfDocument document,
        string path,
        PdfRtfImportOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return ReadForRtf(document, options).SaveAsRtf(path, options);
    }

    /// <summary>Converts an opened PDF and saves the editable RTF document to a caller-owned stream.</summary>
    public static PdfRtfConversionReport SaveAsRtf(
        this PdfCore.PdfDocument document,
        Stream stream,
        PdfRtfImportOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return ReadForRtf(document, options).SaveAsRtf(stream, options);
    }

    /// <summary>Converts an opened PDF and asynchronously saves the editable RTF document to a file.</summary>
    public static Task<PdfRtfConversionReport> SaveAsRtfAsync(
        this PdfCore.PdfDocument document,
        string path,
        PdfRtfImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return ReadForRtf(document, options, cancellationToken).SaveAsRtfAsync(path, options, cancellationToken);
    }

    /// <summary>Converts an opened PDF and asynchronously saves the editable RTF document to a caller-owned stream.</summary>
    public static Task<PdfRtfConversionReport> SaveAsRtfAsync(
        this PdfCore.PdfDocument document,
        Stream stream,
        PdfRtfImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return ReadForRtf(document, options, cancellationToken).SaveAsRtfAsync(stream, options, cancellationToken);
    }

    private static PdfCore.PdfDocumentReadResult ReadForRtf(
        PdfCore.PdfDocument document,
        PdfRtfImportOptions? options,
        CancellationToken cancellationToken = default) =>
        document.Read(options?.ReadOptions, cancellationToken);

    /// <summary>Converts a logical PDF model into an editable RTF document.</summary>
    public static RtfDocument ToRtfDocument(
        this PdfCore.PdfDocumentReadResult document,
        PdfRtfImportOptions? options = null) =>
        document.ToRtfDocumentResult(options).Value;

    /// <summary>Converts a logical PDF model into an editable RTF document with conversion diagnostics.</summary>
    public static PdfRtfConversionResult ToRtfDocumentResult(
        this PdfCore.PdfDocumentReadResult document,
        PdfRtfImportOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        PdfRtfImportOptions operation = (options ?? new PdfRtfImportOptions()).CloneForConversion();
        RtfDocument value = PdfRtfConverter.Convert(document, operation);
        return new PdfRtfConversionResult(value, operation.Report);
    }

    /// <summary>Converts a logical PDF model and saves the editable RTF document to a file.</summary>
    public static PdfRtfConversionReport SaveAsRtf(
        this PdfCore.PdfDocumentReadResult document,
        string path,
        PdfRtfImportOptions? options = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Document path cannot be empty.", nameof(path));
        PdfRtfConversionResult result = document.ToRtfDocumentResult(options);
        result.Value.Save(path);
        return result.Report;
    }

    /// <summary>Converts a logical PDF model and saves the editable RTF document to a caller-owned stream.</summary>
    public static PdfRtfConversionReport SaveAsRtf(
        this PdfCore.PdfDocumentReadResult document,
        Stream stream,
        PdfRtfImportOptions? options = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(stream));
        PdfRtfConversionResult result = document.ToRtfDocumentResult(options);
        result.Value.Save(stream);
        return result.Report;
    }

    /// <summary>Converts a logical PDF model and asynchronously saves the editable RTF document to a file.</summary>
    public static async Task<PdfRtfConversionReport> SaveAsRtfAsync(
        this PdfCore.PdfDocumentReadResult document,
        string path,
        PdfRtfImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Document path cannot be empty.", nameof(path));
        cancellationToken.ThrowIfCancellationRequested();
        PdfRtfConversionResult result = document.ToRtfDocumentResult(options);
        await result.Value.SaveAsync(path, cancellationToken: cancellationToken).ConfigureAwait(false);
        return result.Report;
    }

    /// <summary>Converts a logical PDF model and asynchronously saves the editable RTF document to a caller-owned stream.</summary>
    public static async Task<PdfRtfConversionReport> SaveAsRtfAsync(
        this PdfCore.PdfDocumentReadResult document,
        Stream stream,
        PdfRtfImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(stream));
        cancellationToken.ThrowIfCancellationRequested();
        PdfRtfConversionResult result = document.ToRtfDocumentResult(options);
        await result.Value.SaveAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
        return result.Report;
    }
}
