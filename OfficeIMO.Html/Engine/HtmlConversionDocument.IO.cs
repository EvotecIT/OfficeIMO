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
    public static HtmlConversionDocument Load(Stream stream, HtmlConversionDocumentOptions? options = null, Encoding? encoding = null) {
        HtmlConversionDocumentOptions resolved = (options ?? new HtmlConversionDocumentOptions()).Clone();
        resolved.Validate();
        return Parse(ReadText(stream, encoding ?? Utf8WithoutBom, resolved.Limits), resolved);
    }

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
        HtmlConversionDocumentOptions resolved = (options ?? new HtmlConversionDocumentOptions()).Clone();
        resolved.Validate();
        string html = await ReadTextAsync(stream, encoding ?? Utf8WithoutBom, resolved.Limits, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return Parse(html, resolved);
    }

    private static string ReadText(Stream stream, Encoding encoding, HtmlConversionLimits limits) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        long position = stream.CanSeek ? stream.Position : 0;
        try {
            if (stream.CanSeek) stream.Position = 0;
            using var reader = new StreamReader(stream, encoding, true, 1024, true);
            var builder = new StringBuilder();
            var buffer = new char[4096];
            int read;
            while ((read = reader.Read(buffer, 0, buffer.Length)) > 0) {
                ReserveSourceCharacters(builder.Length, read, limits);
                builder.Append(buffer, 0, read);
            }

            return builder.ToString();
        } finally {
            if (stream.CanSeek) stream.Position = position;
        }
    }

    private static async Task<string> ReadTextAsync(Stream stream, Encoding encoding, HtmlConversionLimits limits, CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        long position = stream.CanSeek ? stream.Position : 0;
        try {
            if (stream.CanSeek) stream.Position = 0;
            using var reader = new StreamReader(stream, encoding, true, 1024, true);
            var builder = new StringBuilder();
            var buffer = new char[4096];
            while (true) {
#if NET8_0_OR_GREATER
                int read = await reader.ReadAsync(buffer.AsMemory(0, buffer.Length), cancellationToken).ConfigureAwait(false);
#else
                int read = await reader.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false);
#endif
                cancellationToken.ThrowIfCancellationRequested();
                if (read == 0) return builder.ToString();
                ReserveSourceCharacters(builder.Length, read, limits);
                builder.Append(buffer, 0, read);
            }
        } finally {
            if (stream.CanSeek) stream.Position = position;
        }
    }

    private static void ReserveSourceCharacters(int current, int additional, HtmlConversionLimits limits) {
        if (!limits.MaxInputCharacters.HasValue || additional <= limits.MaxInputCharacters.Value - current) return;
        throw new HtmlDomLimitException(
            HtmlRenderDiagnosticCodes.InputCharacterLimitExceeded,
            "HTML source length exceeded the configured conversion limit while reading the input stream.",
            nameof(HtmlConversionLimits.MaxInputCharacters),
            (long)current + additional,
            limits.MaxInputCharacters.Value);
    }

    private static HtmlConversionDocumentOptions WithPathBaseUri(string path, HtmlConversionDocumentOptions? options) {
        HtmlConversionDocumentOptions resolved = options?.Clone() ?? new HtmlConversionDocumentOptions();
        resolved.BaseUri ??= new Uri(Path.GetFullPath(path));
        return resolved;
    }
}
