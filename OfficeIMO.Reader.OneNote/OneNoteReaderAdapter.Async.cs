using OfficeIMO.OneNote;
using System.Threading.Tasks;

namespace OfficeIMO.Reader.OneNote;

internal static partial class OneNoteReaderAdapter {
    /// <summary>Schedules offline OneNote path ingestion without blocking the caller thread.</summary>
    public static Task<IReadOnlyList<ReaderChunk>> ReadAsync(
        string oneNotePath,
        ReaderOptions? readerOptions = null,
        ReaderOneNoteOptions? oneNoteOptions = null,
        CancellationToken cancellationToken = default) {
        return Task.Run<IReadOnlyList<ReaderChunk>>(
            () => Read(oneNotePath, readerOptions, oneNoteOptions, cancellationToken).ToArray(),
            cancellationToken);
    }

    /// <summary>Schedules offline OneNote stream ingestion without closing the caller-owned stream.</summary>
    public static Task<IReadOnlyList<ReaderChunk>> ReadAsync(
        Stream oneNoteStream,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        ReaderOneNoteOptions? oneNoteOptions = null,
        CancellationToken cancellationToken = default) {
        return Task.Run<IReadOnlyList<ReaderChunk>>(
            () => Read(oneNoteStream, sourceName, readerOptions, oneNoteOptions, cancellationToken).ToArray(),
            cancellationToken);
    }

    /// <summary>Schedules rich offline OneNote path ingestion without blocking the caller thread.</summary>
    public static Task<OfficeDocumentReadResult> ReadDocumentAsync(
        string oneNotePath,
        ReaderOptions? readerOptions = null,
        ReaderOneNoteOptions? oneNoteOptions = null,
        CancellationToken cancellationToken = default) {
        return Task.Run(
            () => ReadDocument(oneNotePath, readerOptions, oneNoteOptions, cancellationToken),
            cancellationToken);
    }

    /// <summary>Schedules rich offline OneNote stream ingestion without closing the caller-owned stream.</summary>
    public static Task<OfficeDocumentReadResult> ReadDocumentAsync(
        Stream oneNoteStream,
        string? sourceName = null,
        ReaderOptions? readerOptions = null,
        ReaderOneNoteOptions? oneNoteOptions = null,
        CancellationToken cancellationToken = default) {
        return Task.Run(
            () => ReadDocument(oneNoteStream, sourceName, readerOptions, oneNoteOptions, cancellationToken),
            cancellationToken);
    }

    /// <summary>Schedules projection of an already loaded OneNote section.</summary>
    public static Task<OfficeDocumentReadResult> ReadDocumentAsync(
        OneNoteSection section,
        string sourceName = "section.one",
        ReaderOptions? readerOptions = null,
        ReaderOneNoteOptions? oneNoteOptions = null,
        CancellationToken cancellationToken = default) {
        return Task.Run(
            () => ReadDocument(section, sourceName, readerOptions, oneNoteOptions, cancellationToken),
            cancellationToken);
    }

    /// <summary>Schedules projection of an already loaded OneNote notebook.</summary>
    public static Task<OfficeDocumentReadResult> ReadDocumentAsync(
        OneNoteNotebook notebook,
        string sourceName = "notebook.onepkg",
        ReaderOptions? readerOptions = null,
        ReaderOneNoteOptions? oneNoteOptions = null,
        CancellationToken cancellationToken = default) {
        return Task.Run(
            () => ReadDocument(notebook, sourceName, readerOptions, oneNoteOptions, cancellationToken),
            cancellationToken);
    }
}
