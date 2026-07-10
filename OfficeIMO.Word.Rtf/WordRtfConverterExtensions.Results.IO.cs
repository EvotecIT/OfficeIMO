namespace OfficeIMO.Word.Rtf;

/// <content>Provides result-bearing RTF ingestion APIs for Word.</content>
public static partial class WordRtfConverterExtensions {
    /// <summary>Reads RTF text with the bounded profile and returns Word plus a combined report.</summary>
    public static RtfConversionResult<WordDocument> LoadFromRtfResult(this string rtf, RtfReadOptions? readOptions = null) {
        RtfReadResult read = RtfDocument.Read(rtf, readOptions ?? RtfReadOptions.CreateUntrustedProfile());
        return CreateWordReadResult(read, "string");
    }

    /// <summary>Reads RTF bytes with the bounded profile and returns Word plus a combined report.</summary>
    public static RtfConversionResult<WordDocument> LoadFromRtfResult(this byte[] rtfBytes, RtfReadOptions? readOptions = null) {
        RtfReadResult read = RtfDocument.Load(rtfBytes, readOptions ?? RtfReadOptions.CreateUntrustedProfile());
        return CreateWordReadResult(read, "bytes");
    }

    /// <summary>Reads an RTF stream with the bounded profile and returns Word plus a combined report.</summary>
    public static RtfConversionResult<WordDocument> LoadFromRtfResult(this Stream rtfStream, RtfReadOptions? readOptions = null, Encoding? encoding = null) {
        RtfReadResult read = RtfDocument.Load(rtfStream, readOptions ?? RtfReadOptions.CreateUntrustedProfile(), encoding);
        return CreateWordReadResult(read, "stream");
    }

    /// <summary>Reads an RTF file with the bounded profile and returns Word plus a combined report.</summary>
    public static RtfConversionResult<WordDocument> LoadFromRtfFileResult(string path, RtfReadOptions? readOptions = null, Encoding? encoding = null) {
        RtfReadResult read = RtfDocument.Load(path, readOptions ?? RtfReadOptions.CreateUntrustedProfile(), encoding);
        return CreateWordReadResult(read, path);
    }

    /// <summary>Asynchronously reads RTF text and returns Word plus a combined report.</summary>
    public static async Task<RtfConversionResult<WordDocument>> LoadFromRtfResultAsync(this string rtf, RtfReadOptions? readOptions = null, CancellationToken cancellationToken = default) {
        RtfReadResult read = await RtfDocument.ReadAsync(rtf, readOptions ?? RtfReadOptions.CreateUntrustedProfile(), cancellationToken).ConfigureAwait(false);
        return CreateWordReadResult(read, "string");
    }

    /// <summary>Asynchronously reads RTF bytes and returns Word plus a combined report.</summary>
    public static async Task<RtfConversionResult<WordDocument>> LoadFromRtfResultAsync(this byte[] rtfBytes, RtfReadOptions? readOptions = null, CancellationToken cancellationToken = default) {
        RtfReadResult read = await RtfDocument.LoadAsync(rtfBytes, readOptions ?? RtfReadOptions.CreateUntrustedProfile(), cancellationToken).ConfigureAwait(false);
        return CreateWordReadResult(read, "bytes");
    }

    /// <summary>Asynchronously reads an RTF stream and returns Word plus a combined report.</summary>
    public static async Task<RtfConversionResult<WordDocument>> LoadFromRtfResultAsync(this Stream rtfStream, RtfReadOptions? readOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        RtfReadResult read = await RtfDocument.LoadAsync(rtfStream, readOptions ?? RtfReadOptions.CreateUntrustedProfile(), encoding, cancellationToken).ConfigureAwait(false);
        return CreateWordReadResult(read, "stream");
    }

    /// <summary>Asynchronously reads an RTF file and returns Word plus a combined report.</summary>
    public static async Task<RtfConversionResult<WordDocument>> LoadFromRtfFileResultAsync(string path, RtfReadOptions? readOptions = null, Encoding? encoding = null, CancellationToken cancellationToken = default) {
        RtfReadResult read = await RtfDocument.LoadAsync(path, readOptions ?? RtfReadOptions.CreateUntrustedProfile(), encoding, cancellationToken).ConfigureAwait(false);
        return CreateWordReadResult(read, path);
    }

    private static RtfConversionResult<WordDocument> CreateWordReadResult(RtfReadResult read, string sourcePath) {
        RtfConversionResult<WordDocument> converted = read.Document.ToWordDocumentResult();
        var report = new RtfConversionReport();
        report.AddReadDiagnostics(read.Diagnostics, sourcePath);
        report.Merge(converted.Report);
        return new RtfConversionResult<WordDocument>(converted.Value, report);
    }
}
