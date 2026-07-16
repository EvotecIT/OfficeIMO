using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

public sealed partial class OfficeDocumentReader {
    /// <summary>
    /// Converts a supported file to its portable Markdown representation.
    /// </summary>
    /// <remarks>
    /// This is a thin projection over <see cref="ReadDocument(string, ReaderOptions?, CancellationToken)"/>.
    /// Use <c>ReadDocument(...)</c> when source metadata, blocks, tables, assets, or diagnostics are required.
    /// </remarks>
    /// <returns>The emitted Markdown, or an empty string when the reader produced no Markdown.</returns>
    public string ConvertToMarkdown(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return ReadDocument(path, options, cancellationToken).Markdown ?? string.Empty;
    }

    /// <summary>
    /// Converts a supported stream to its portable Markdown representation. The caller retains ownership of the stream.
    /// </summary>
    /// <remarks>
    /// This is a thin projection over <see cref="ReadDocument(Stream, string?, ReaderOptions?, CancellationToken)"/>.
    /// Use <c>ReadDocument(...)</c> when source metadata, blocks, tables, assets, or diagnostics are required.
    /// </remarks>
    /// <returns>The emitted Markdown, or an empty string when the reader produced no Markdown.</returns>
    public string ConvertToMarkdown(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return ReadDocument(stream, sourceName, options, cancellationToken).Markdown ?? string.Empty;
    }

    /// <summary>
    /// Converts supported bytes to their portable Markdown representation.
    /// </summary>
    /// <remarks>
    /// This is a thin projection over <see cref="ReadDocument(byte[], string?, ReaderOptions?, CancellationToken)"/>.
    /// Use <c>ReadDocument(...)</c> when source metadata, blocks, tables, assets, or diagnostics are required.
    /// </remarks>
    /// <returns>The emitted Markdown, or an empty string when the reader produced no Markdown.</returns>
    public string ConvertToMarkdown(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        return ReadDocument(bytes, sourceName, options, cancellationToken).Markdown ?? string.Empty;
    }

    /// <summary>
    /// Asynchronously converts a supported file to its portable Markdown representation.
    /// </summary>
    /// <remarks>
    /// This is a thin projection over <see cref="ReadDocumentAsync(string, ReaderOptions?, CancellationToken)"/>.
    /// Use <c>ReadDocumentAsync(...)</c> when source metadata, blocks, tables, assets, or diagnostics are required.
    /// </remarks>
    /// <returns>The emitted Markdown, or an empty string when the reader produced no Markdown.</returns>
    public async Task<string> ConvertToMarkdownAsync(
        string path,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        OfficeDocumentReadResult document = await ReadDocumentAsync(path, options, cancellationToken).ConfigureAwait(false);
        return document.Markdown ?? string.Empty;
    }

    /// <summary>
    /// Asynchronously converts a supported stream to its portable Markdown representation. The caller retains ownership of the stream.
    /// </summary>
    /// <remarks>
    /// This is a thin projection over <see cref="ReadDocumentAsync(Stream, string?, ReaderOptions?, CancellationToken)"/>.
    /// Use <c>ReadDocumentAsync(...)</c> when source metadata, blocks, tables, assets, or diagnostics are required.
    /// </remarks>
    /// <returns>The emitted Markdown, or an empty string when the reader produced no Markdown.</returns>
    public async Task<string> ConvertToMarkdownAsync(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        OfficeDocumentReadResult document = await ReadDocumentAsync(
            stream,
            sourceName,
            options,
            cancellationToken).ConfigureAwait(false);
        return document.Markdown ?? string.Empty;
    }

    /// <summary>
    /// Asynchronously converts supported bytes to their portable Markdown representation.
    /// </summary>
    /// <remarks>
    /// This is a thin projection over <see cref="ReadDocumentAsync(byte[], string?, ReaderOptions?, CancellationToken)"/>.
    /// Use <c>ReadDocumentAsync(...)</c> when source metadata, blocks, tables, assets, or diagnostics are required.
    /// </remarks>
    /// <returns>The emitted Markdown, or an empty string when the reader produced no Markdown.</returns>
    public async Task<string> ConvertToMarkdownAsync(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        OfficeDocumentReadResult document = await ReadDocumentAsync(
            bytes,
            sourceName,
            options,
            cancellationToken).ConfigureAwait(false);
        return document.Markdown ?? string.Empty;
    }
}
