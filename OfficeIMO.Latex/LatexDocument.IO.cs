using OfficeIMO.Core.Internal;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Latex;

public sealed partial class LatexDocument {
    private static readonly Encoding Utf8WithoutBom = new UTF8Encoding(false);

    /// <summary>Loads and parses a LaTeX document from a caller-owned stream.</summary>
    public static LatexParseResult Load(Stream stream, LatexParseOptions? options = null, Encoding? encoding = null) {
        options ??= new LatexParseOptions();
        byte[] bytes = OfficeStreamReader.ReadAllBytes(stream, options.MaximumInputBytes);
        return Parse(Decode(bytes, encoding ?? Utf8WithoutBom, options), options);
    }

    /// <summary>Asynchronously loads and parses a LaTeX file.</summary>
    public static async Task<LatexParseResult> LoadAsync(
        string path,
        LatexParseOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, true);
        return await LoadAsync(stream, options, encoding, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously loads and parses a LaTeX document from a caller-owned stream.</summary>
    public static async Task<LatexParseResult> LoadAsync(
        Stream stream,
        LatexParseOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) {
        options ??= new LatexParseOptions();
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(
            stream,
            cancellationToken,
            options.MaximumInputBytes).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        string source = Decode(bytes, encoding ?? Utf8WithoutBom, options);
        return Parse(source, options, cancellationToken);
    }

    private static string Decode(byte[] bytes, Encoding encoding, LatexParseOptions options) {
        byte[] preamble = encoding.GetPreamble();
        int offset = StartsWith(bytes, preamble)
            ? preamble.Length
            : encoding.CodePage == Encoding.UTF8.CodePage && StartsWith(bytes, new byte[] { 0xEF, 0xBB, 0xBF })
                ? 3
                : 0;
        int count = bytes.Length - offset;
        int characterCount = encoding.GetCharCount(bytes, offset, count);
        if (options.MaximumInputLength.HasValue && characterCount > options.MaximumInputLength.Value) {
            throw new InvalidDataException(
                $"Decoded LaTeX source exceeds MaximumInputLength ({options.MaximumInputLength.Value} characters).");
        }
        return encoding.GetString(bytes, offset, count);
    }

    private static bool StartsWith(byte[] source, byte[] prefix) {
        if (prefix.Length == 0 || prefix.Length > source.Length) return false;
        for (int index = 0; index < prefix.Length; index++) {
            if (source[index] != prefix[index]) return false;
        }
        return true;
    }

    /// <summary>Encodes the current document text.</summary>
    public byte[] ToBytes(LatexWriterOptions? options = null, Encoding? encoding = null) =>
        (encoding ?? Utf8WithoutBom).GetBytes(ToLatex(options));

    /// <summary>Encodes the current document in a new writable memory stream positioned at the beginning.</summary>
    public MemoryStream ToStream(LatexWriterOptions? options = null, Encoding? encoding = null) =>
        new MemoryStream(ToBytes(options, encoding));

    /// <summary>Saves the current document text to a file.</summary>
    public void Save(string path, LatexWriterOptions? options = null, Encoding? encoding = null) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        OfficeFileCommit.WriteAllBytes(path, ToBytes(options, encoding));
    }

    /// <summary>Writes the current document text to a caller-owned stream.</summary>
    public void Save(Stream stream, LatexWriterOptions? options = null, Encoding? encoding = null) =>
        OfficeStreamWriter.WriteAllBytes(stream, ToBytes(options, encoding));

    /// <summary>Asynchronously saves the current document text to a file.</summary>
    public async Task SaveAsync(
        string path,
        LatexWriterOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        await OfficeFileCommit.WriteAllBytesAsync(
            path,
            ToBytes(options, encoding),
            cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes the current document text to a caller-owned stream.</summary>
    public Task SaveAsync(
        Stream stream,
        LatexWriterOptions? options = null,
        Encoding? encoding = null,
        CancellationToken cancellationToken = default) =>
        OfficeStreamWriter.WriteAllBytesAsync(stream, ToBytes(options, encoding), cancellationToken);
}
