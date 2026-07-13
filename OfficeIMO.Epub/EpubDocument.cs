namespace OfficeIMO.Epub;

using OfficeIMO.Drawing.Internal;
using System.Threading;
using System.Threading.Tasks;

/// <summary>
/// Represents extracted EPUB content.
/// </summary>
public sealed class EpubDocument {
    /// <summary>Loads an EPUB document from a file.</summary>
    public static EpubDocument Load(string path, EpubReadOptions? options = null) => EpubReader.Read(path, options);

    /// <summary>Loads an EPUB document from a caller-owned stream.</summary>
    public static EpubDocument Load(Stream stream, EpubReadOptions? options = null) => EpubReader.Read(stream, options);

    /// <summary>Asynchronously loads an EPUB document from a file.</summary>
    public static async Task<EpubDocument> LoadAsync(
        string path,
        EpubReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("EPUB path cannot be empty.", nameof(path));
        using var stream = new FileStream(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete,
            81920,
            true);
        return await LoadAsync(stream, options, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously loads an EPUB document from a caller-owned stream.</summary>
    public static async Task<EpubDocument> LoadAsync(
        Stream stream,
        EpubReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return EpubReader.ReadBytes(bytes, options);
    }

    /// <summary>
    /// Best-effort document title.
    /// </summary>
    public string? Title { get; internal set; }

    /// <summary>
    /// Package identifier from OPF metadata when available.
    /// </summary>
    public string? Identifier { get; internal set; }

    /// <summary>
    /// Primary language from OPF metadata when available.
    /// </summary>
    public string? Language { get; internal set; }

    /// <summary>
    /// Creator/author from OPF metadata when available.
    /// </summary>
    public string? Creator { get; internal set; }

    /// <summary>
    /// Internal path to the OPF package document when discovered.
    /// </summary>
    public string? OpfPath { get; internal set; }

    /// <summary>
    /// Extracted chapters.
    /// </summary>
    public IReadOnlyList<EpubChapter> Chapters { get; internal set; } = Array.Empty<EpubChapter>();

    /// <summary>
    /// OPF manifest resources in deterministic package order.
    /// </summary>
    public IReadOnlyList<EpubResource> Resources { get; internal set; } = Array.Empty<EpubResource>();

    /// <summary>
    /// Non-fatal warnings encountered during extraction.
    /// </summary>
    public IReadOnlyList<string> Warnings { get; internal set; } = Array.Empty<string>();
}
