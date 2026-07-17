using OfficeIMO.Email;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private static IEnumerable<ReaderChunk> ReadCalendar(string path, ReaderOptions options,
        CancellationToken cancellationToken) {
        IcsDocument document = IcsDocument.Load(path, CreateContentLineOptions(options), cancellationToken);
        return ChunkContentLineDocument(document.Serialize(CreateContentLineWriterOptions(options)),
            path, Path.GetFileName(path), options,
            ReaderInputKind.Calendar, cancellationToken);
    }

    private static IEnumerable<ReaderChunk> ReadCalendar(Stream stream, string? sourceName,
        ReaderOptions options, CancellationToken cancellationToken) {
        IcsDocument document = IcsDocument.Load(stream, CreateContentLineOptions(options), cancellationToken);
        string fileName = string.IsNullOrWhiteSpace(sourceName) ? "calendar.ics" : Path.GetFileName(sourceName!.Trim());
        return ChunkContentLineDocument(document.Serialize(CreateContentLineWriterOptions(options)),
            sourceName, fileName, options,
            ReaderInputKind.Calendar, cancellationToken);
    }

    private static IEnumerable<ReaderChunk> ReadVCard(string path, ReaderOptions options,
        CancellationToken cancellationToken) {
        VCardDocument document = VCardDocument.Load(path, CreateContentLineOptions(options), cancellationToken);
        return ChunkContentLineDocument(document.Serialize(CreateContentLineWriterOptions(options)),
            path, Path.GetFileName(path), options,
            ReaderInputKind.VCard, cancellationToken);
    }

    private static IEnumerable<ReaderChunk> ReadVCard(Stream stream, string? sourceName,
        ReaderOptions options, CancellationToken cancellationToken) {
        VCardDocument document = VCardDocument.Load(stream, CreateContentLineOptions(options), cancellationToken);
        string fileName = string.IsNullOrWhiteSpace(sourceName) ? "contact.vcf" : Path.GetFileName(sourceName!.Trim());
        return ChunkContentLineDocument(document.Serialize(CreateContentLineWriterOptions(options)),
            sourceName, fileName, options,
            ReaderInputKind.VCard, cancellationToken);
    }

    private static IEnumerable<ReaderChunk> ChunkContentLineDocument(string text, string? sourceName,
        string fileName, ReaderOptions options, ReaderInputKind kind, CancellationToken cancellationToken) =>
        ChunkPlainTextFromText(text, sourceName, fileName, options, kind, cancellationToken,
            treatAsMarkdown: false);

    private static ContentLineReaderOptions CreateContentLineOptions(ReaderOptions options) {
        ContentLineReaderOptions defaults = ContentLineReaderOptions.Default;
        long maxInputBytes = options.MaxInputBytes.GetValueOrDefault(defaults.MaxInputBytes);
        if (maxInputBytes <= 0) maxInputBytes = defaults.MaxInputBytes;
        return new ContentLineReaderOptions(maxInputBytes, defaults.MaxUnfoldedLineBytes,
            defaults.MaxComponents, defaults.MaxProperties, defaults.MaxNestingDepth, defaults.Encoding);
    }

    private static ContentLineWriterOptions CreateContentLineWriterOptions(ReaderOptions options) {
        long maximumInput = options.MaxInputBytes.GetValueOrDefault(
            ContentLineReaderOptions.Default.MaxInputBytes);
        if (maximumInput <= 0) maximumInput = ContentLineReaderOptions.Default.MaxInputBytes;
        long maximumOutput = maximumInput <= long.MaxValue / 2 ? maximumInput * 2 : long.MaxValue;
        ContentLineWriterOptions defaults = ContentLineWriterOptions.Default;
        return new ContentLineWriterOptions(defaults.FoldAtOctets, maximumOutput, defaults.Encoding);
    }
}
