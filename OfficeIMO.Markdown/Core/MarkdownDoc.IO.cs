using OfficeIMO.Drawing.Internal;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Markdown;

public partial class MarkdownDoc {
    private static readonly Encoding Utf8WithoutBom = new UTF8Encoding(false);

    /// <summary>Loads and parses a Markdown file.</summary>
    public static MarkdownDoc Load(string path, MarkdownReaderOptions? options = null, Encoding? encoding = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Load(stream, options, encoding);
    }

    /// <summary>Loads and parses Markdown from a caller-owned stream.</summary>
    public static MarkdownDoc Load(Stream stream, MarkdownReaderOptions? options = null, Encoding? encoding = null) =>
        MarkdownReader.Parse(DecodeText(OfficeStreamReader.ReadAllBytes(stream), encoding ?? Utf8WithoutBom), options);

    /// <summary>Asynchronously loads and parses a Markdown file.</summary>
    public static async Task<MarkdownDoc> LoadAsync(
        string path,
        MarkdownReaderOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, true);
        return await LoadAsync(stream, options, encoding, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously loads and parses Markdown from a caller-owned stream.</summary>
    public static async Task<MarkdownDoc> LoadAsync(
        Stream stream,
        MarkdownReaderOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) {
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return MarkdownReader.Parse(DecodeText(bytes, encoding ?? Utf8WithoutBom), options);
    }

    /// <summary>Encodes this document as Markdown.</summary>
    public byte[] ToBytes(MarkdownWriteOptions? options = null, Encoding? encoding = null) =>
        (encoding ?? Utf8WithoutBom).GetBytes(ToMarkdown(options));

    /// <summary>Encodes this document in a new writable stream positioned at the beginning.</summary>
    public MemoryStream ToStream(MarkdownWriteOptions? options = null, Encoding? encoding = null) =>
        new MemoryStream(ToBytes(options, encoding));

    /// <summary>Saves this document as Markdown.</summary>
    public void Save(string path, MarkdownWriteOptions? options = null, Encoding? encoding = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        OfficeFileCommit.WriteAllBytes(path, ToBytes(options, encoding));
    }

    /// <summary>Writes this document as Markdown to a caller-owned stream.</summary>
    public void Save(Stream stream, MarkdownWriteOptions? options = null, Encoding? encoding = null) {
        OfficeStreamWriter.WriteAllBytes(stream, ToBytes(options, encoding));
    }

    /// <summary>Asynchronously saves this document as Markdown.</summary>
    public async Task SaveAsync(
        string path,
        MarkdownWriteOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        await OfficeFileCommit.WriteAllBytesAsync(
            path,
            ToBytes(options, encoding),
            cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes this document as Markdown to a caller-owned stream.</summary>
    public async Task SaveAsync(
        Stream stream,
        MarkdownWriteOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) {
        await OfficeStreamWriter.WriteAllBytesAsync(stream, ToBytes(options, encoding), cancellationToken).ConfigureAwait(false);
    }

    private static string DecodeText(byte[] bytes, Encoding encoding) {
        using var stream = new MemoryStream(bytes, writable: false);
        using var reader = new StreamReader(stream, encoding, true, 1024, false);
        return reader.ReadToEnd();
    }
}
