using OfficeIMO.Core.Internal;

namespace OfficeIMO.Html;

/// <summary>A bounded ZIP website snapshot with an inert HTML root and archive-owned resource bytes.
/// Loading does not extract files, execute scripts or fetch resources from the host or network.</summary>
public sealed partial class HtmlSiteBundle {
    private readonly Dictionary<string, HtmlResolvedResource> _resources;

    private HtmlSiteBundle(HtmlConversionDocument document, string entryPath, Uri archiveBaseUri,
        Dictionary<string, HtmlResolvedResource> resources, IReadOnlyList<string> entryNames,
        long archiveBytes, long decodedBytes) {
        HtmlDocument = document;
        EntryPath = entryPath;
        ArchiveBaseUri = archiveBaseUri;
        _resources = resources;
        EntryNames = entryNames;
        ArchiveBytes = archiveBytes;
        DecodedBytes = decodedBytes;
    }

    /// <summary>Parsed HTML entry document.</summary>
    public HtmlConversionDocument HtmlDocument { get; }
    /// <summary>Selected normalized, case-sensitive HTML entry path.</summary>
    public string EntryPath { get; }
    /// <summary>Virtual directory URI that maps archive paths.</summary>
    public Uri ArchiveBaseUri { get; }
    /// <summary>Virtual URI of the selected HTML entry.</summary>
    public Uri BaseUri => CreateEntryUri(ArchiveBaseUri, EntryPath);
    /// <summary>Regular-file entry names in ordinal order.</summary>
    public IReadOnlyList<string> EntryNames { get; }
    /// <summary>Encoded ZIP size observed when loading.</summary>
    public long ArchiveBytes { get; }
    /// <summary>Combined decoded size of all regular-file entries.</summary>
    public long DecodedBytes { get; }

    /// <summary>Loads a ZIP site bundle from a file, without extracting it.</summary>
    public static HtmlSiteBundle Load(string path, HtmlSiteBundleOptions? options = null,
        HtmlConversionDocumentOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        cancellationToken.ThrowIfCancellationRequested();
        using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Load(source, options, htmlOptions, cancellationToken);
    }

    /// <summary>Loads a ZIP site bundle from a caller-owned stream. Seekable input is read from the
    /// beginning and its original position is restored; nonseekable input is read forward.</summary>
    public static HtmlSiteBundle Load(Stream stream, HtmlSiteBundleOptions? options = null,
        HtmlConversionDocumentOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        HtmlSiteBundleOptions resolved = (options ?? new HtmlSiteBundleOptions()).Snapshot();
        HtmlConversionDocumentOptions parsing = (htmlOptions ?? new HtmlConversionDocumentOptions()).Clone();
        parsing.Validate();
        byte[] bytes = OfficeStreamReader.ReadAllBytes(stream, cancellationToken, resolved.MaximumArchiveBytes);
        return ReadArchiveAsync(bytes, resolved, parsing, asynchronous: false, cancellationToken)
            .GetAwaiter().GetResult();
    }

    /// <summary>Asynchronously loads a ZIP site bundle from a file, without extracting it.</summary>
    public static async Task<HtmlSiteBundle> LoadAsync(string path, HtmlSiteBundleOptions? options = null,
        HtmlConversionDocumentOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        cancellationToken.ThrowIfCancellationRequested();
        using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, true);
        return await LoadAsync(source, options, htmlOptions, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously loads a ZIP site bundle while leaving the caller's stream open and
    /// restoring its original position when seekable.</summary>
    public static async Task<HtmlSiteBundle> LoadAsync(Stream stream, HtmlSiteBundleOptions? options = null,
        HtmlConversionDocumentOptions? htmlOptions = null, CancellationToken cancellationToken = default) {
        HtmlSiteBundleOptions resolved = (options ?? new HtmlSiteBundleOptions()).Snapshot();
        HtmlConversionDocumentOptions parsing = (htmlOptions ?? new HtmlConversionDocumentOptions()).Clone();
        parsing.Validate();
        byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken,
            resolved.MaximumArchiveBytes).ConfigureAwait(false);
        return await ReadArchiveAsync(bytes, resolved, parsing, asynchronous: true, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Creates a resolver serving only archive snapshots. It never reads local or remote files.</summary>
    public HtmlRenderResourceResolver CreateResourceResolver() => ResolveResourceAsync;

    /// <summary>Snapshots a display-list, SVG or raster request with archive-first resources.
    /// Any explicitly supplied resolver remains the fallback, subject to the renderer's URL and byte limits.
    /// PDF callers use the owning PDF adapter so its embedded-package policy also applies.</summary>
    public HtmlRenderRequest CreateRenderRequest(HtmlRenderRequest request) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (request.Encoder == HtmlRenderEncoder.Pdf) {
            throw new ArgumentException("Use the site-bundle PDF adapter for the Pdf encoder.", nameof(request));
        }
        HtmlRenderOptions options = request.Options;
        options.BaseUri ??= BaseUri;
        HtmlRenderResourceResolver? fallback = options.ResourceResolver;
        options.ResourceResolver = fallback == null ? ResolveResourceAsync : async (resource, token) =>
            await ResolveResourceAsync(resource, token).ConfigureAwait(false)
                ?? await fallback(resource, token).ConfigureAwait(false);
        HtmlRenderSynchronousResourceResolver? synchronousFallback = options.SynchronousResourceResolver;
        options.SynchronousResourceResolver = (HtmlRenderResourceRequest resource, CancellationToken token,
            out HtmlResolvedResource? resolved) => {
            if (TryResolveResource(resource, token, out resolved)) return true;
            return synchronousFallback != null && synchronousFallback(resource, token, out resolved);
        };
        return request.WithOptions(options);
    }

    private Task<HtmlResolvedResource?> ResolveResourceAsync(HtmlRenderResourceRequest request,
        CancellationToken cancellationToken) {
        TryResolveResource(request, cancellationToken, out HtmlResolvedResource? resource);
        return Task.FromResult(resource);
    }

    internal bool TryResolveResource(HtmlRenderResourceRequest request, CancellationToken cancellationToken,
        out HtmlResolvedResource? resource) {
        cancellationToken.ThrowIfCancellationRequested();
        resource = null;
        return request.Kind is not (HtmlResourceKind.Script or HtmlResourceKind.Hyperlink)
            && _resources.TryGetValue(ResourceKey(request.Uri), out resource);
    }

    private static Uri CreateEntryUri(Uri archiveBaseUri, string entryPath) => new Uri(archiveBaseUri,
        string.Join("/", entryPath.Split('/').Select(Uri.EscapeDataString)));

    // ZIP files use path identities; HTML query/fragment cache-busters do not select another entry.
    private static string ResourceKey(Uri uri) => uri.GetComponents(
        UriComponents.SchemeAndServer | UriComponents.Path, UriFormat.UriEscaped);
}
