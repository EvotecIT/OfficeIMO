using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Html;

public sealed partial class HtmlConversionDocument {
    private static readonly Encoding Utf8WithoutBom = new UTF8Encoding(false);

    /// <summary>Loads and parses an HTML file.</summary>
    public static HtmlConversionDocument Load(string path, HtmlConversionDocumentOptions? options = null, Encoding? encoding = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Load(stream, WithPathBaseUri(path, options), encoding);
    }

    /// <summary>Loads and parses HTML from a caller-owned stream.</summary>
    public static HtmlConversionDocument Load(Stream stream, HtmlConversionDocumentOptions? options = null, Encoding? encoding = null) =>
        Parse(ReadText(stream, encoding ?? Utf8WithoutBom), options);

    /// <summary>Asynchronously loads and parses an HTML file.</summary>
    public static async Task<HtmlConversionDocument> LoadAsync(
        string path,
        HtmlConversionDocumentOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, true);
        return await LoadAsync(stream, WithPathBaseUri(path, options), encoding, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously loads and parses HTML from a caller-owned stream.</summary>
    public static async Task<HtmlConversionDocument> LoadAsync(
        Stream stream,
        HtmlConversionDocumentOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) {
        string html = await ReadTextAsync(stream, encoding ?? Utf8WithoutBom, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return Parse(html, options);
    }

    private static string ReadText(Stream stream, Encoding encoding) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        long position = stream.CanSeek ? stream.Position : 0;
        try {
            if (stream.CanSeek) stream.Position = 0;
            using var reader = new StreamReader(stream, encoding, true, 1024, true);
            return reader.ReadToEnd();
        } finally {
            if (stream.CanSeek) stream.Position = position;
        }
    }

    private static async Task<string> ReadTextAsync(Stream stream, Encoding encoding, CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        long position = stream.CanSeek ? stream.Position : 0;
        try {
            if (stream.CanSeek) stream.Position = 0;
            using var reader = new StreamReader(stream, encoding, true, 1024, true);
#if NET8_0_OR_GREATER
            return await reader.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
#else
            string text = await reader.ReadToEndAsync().ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            return text;
#endif
        } finally {
            if (stream.CanSeek) stream.Position = position;
        }
    }

    private static HtmlConversionDocumentOptions WithPathBaseUri(string path, HtmlConversionDocumentOptions? options) {
        HtmlConversionDocumentOptions resolved = options?.Clone() ?? new HtmlConversionDocumentOptions();
        resolved.BaseUri ??= new Uri(Path.GetFullPath(path));
        return resolved;
    }
}
