using OfficeIMO.Core.Internal;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.AsciiDoc;

public sealed partial class AsciiDocDocument {
    private static readonly Encoding Utf8WithoutBom = new UTF8Encoding(false);

    /// <summary>Loads and parses an AsciiDoc document from a caller-owned stream.</summary>
    public static AsciiDocDocument Load(Stream stream, AsciiDocParseOptions? options = null, Encoding? encoding = null) =>
        LoadResult(stream, options, encoding).Document;

    /// <summary>Loads an AsciiDoc document from a caller-owned stream with syntax and recovery diagnostics.</summary>
    public static AsciiDocParseResult LoadResult(Stream stream, AsciiDocParseOptions? options = null, Encoding? encoding = null) {
        return LoadResult(stream, options, encoding, CancellationToken.None);
    }

    /// <summary>Loads bounded text with cooperative cancellation while preserving the caller's stream.</summary>
    public static AsciiDocDocument Load(Stream stream, AsciiDocParseOptions? options, Encoding? encoding, CancellationToken cancellationToken) =>
        LoadResult(stream, options, encoding, cancellationToken).Document;

    /// <summary>Loads bounded text and returns recovery diagnostics with cooperative cancellation.</summary>
    public static AsciiDocParseResult LoadResult(Stream stream, AsciiDocParseOptions? options, Encoding? encoding, CancellationToken cancellationToken) {
        options ??= new AsciiDocParseOptions();
        AsciiDocParser.ValidateOptions(string.Empty, options);
        return ParseResult(OfficeTextReader.ReadAllText(stream, options.MaximumInputLength, encoding, cancellationToken), options, cancellationToken);
    }

    /// <summary>Asynchronously loads and parses an AsciiDoc file.</summary>
    public static async Task<AsciiDocDocument> LoadAsync(
        string path,
        AsciiDocParseOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) =>
        (await LoadResultAsync(path, options, encoding, cancellationToken).ConfigureAwait(false)).Document;

    /// <summary>Asynchronously loads an AsciiDoc file with syntax and recovery diagnostics.</summary>
    public static async Task<AsciiDocParseResult> LoadResultAsync(
        string path,
        AsciiDocParseOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        cancellationToken.ThrowIfCancellationRequested();
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, true);
        return await LoadResultAsync(stream, options, encoding, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously loads and parses an AsciiDoc document from a caller-owned stream.</summary>
    public static async Task<AsciiDocDocument> LoadAsync(
        Stream stream,
        AsciiDocParseOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) =>
        (await LoadResultAsync(stream, options, encoding, cancellationToken).ConfigureAwait(false)).Document;

    /// <summary>Asynchronously loads an AsciiDoc stream with syntax and recovery diagnostics.</summary>
    public static async Task<AsciiDocParseResult> LoadResultAsync(
        Stream stream,
        AsciiDocParseOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) {
        options ??= new AsciiDocParseOptions();
        AsciiDocParser.ValidateOptions(string.Empty, options);
        string text = await OfficeTextReader.ReadAllTextAsync(stream, options.MaximumInputLength, encoding, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return ParseResult(text, options, cancellationToken);
    }

    /// <summary>Encodes the current document text.</summary>
    public byte[] ToBytes(AsciiDocWriterOptions? options = null, Encoding? encoding = null) =>
        (encoding ?? Utf8WithoutBom).GetBytes(ToAsciiDoc(options));

    /// <summary>Encodes the current document in a new writable memory stream positioned at the beginning.</summary>
    public MemoryStream ToStream(AsciiDocWriterOptions? options = null, Encoding? encoding = null) =>
        new MemoryStream(ToBytes(options, encoding));

    /// <summary>Saves the current document text to a file.</summary>
    public void Save(string path, AsciiDocWriterOptions? options = null, Encoding? encoding = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        byte[] bytes = ToBytes(options, encoding);
        OfficeFileCommit.WriteAtomically(path, stream => stream.Write(bytes, 0, bytes.Length));
    }

    /// <summary>Writes the current document text to a caller-owned stream.</summary>
    public void Save(Stream stream, AsciiDocWriterOptions? options = null, Encoding? encoding = null) {
        OfficeStreamWriter.WriteAllBytes(stream, ToBytes(options, encoding));
    }

    /// <summary>Asynchronously saves the current document text to a file.</summary>
    public async Task SaveAsync(
        string path,
        AsciiDocWriterOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        cancellationToken.ThrowIfCancellationRequested();
        byte[] bytes = ToBytes(options, encoding);
        await OfficeFileCommit.WriteAtomicallyAsync(path,
            (stream, token) => stream.WriteAsync(bytes, 0, bytes.Length, token),
            cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the current document text to a caller-owned stream.</summary>
    public Task SaveAsync(
        Stream stream,
        AsciiDocWriterOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) =>
        OfficeStreamWriter.WriteAllBytesAsync(stream, ToBytes(options, encoding), cancellationToken);
}
