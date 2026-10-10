namespace OfficeIMO.Chm;

public sealed partial class ChmDocument {
    /// <summary>Creates a private virtual URI for an archive topic or resource. It does not identify a filesystem or network location.</summary>
    public Uri GetTopicUri(string path) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        string canonical = ChmPaths.DirectoryPath(path);
        if (canonical.StartsWith("::", StringComparison.Ordinal)) throw new ArgumentException("A content path is required.", nameof(path));
        return new Uri("chm://archive" + string.Join("/", canonical.Split('/').Select(Uri.EscapeDataString)));
    }

    /// <summary>Creates an embedded-only resolver. Missing and external resources are never fetched.</summary>
    public HtmlRenderResourceResolver CreateResourceResolver() => (request, token) => {
        token.ThrowIfCancellationRequested();
        if (!request.Uri.Scheme.Equals("chm", StringComparison.OrdinalIgnoreCase) || !request.Uri.Host.Equals("archive", StringComparison.OrdinalIgnoreCase))
            return Task.FromResult<HtmlResolvedResource?>(null);
        ChmEntry? entry = FindUriEntry(request.Uri);
        if (entry == null || entry.IsSystem || entry.IsDirectory) return Task.FromResult<HtmlResolvedResource?>(null);
        return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(entry.GetBytes(), GetMediaType(entry.Path)));
    };

    private ChmEntry? FindUriEntry(Uri uri) {
        // A URI carries escaped bytes; decode once before identity lookup. Public
        // FindEntry also accepts exact raw names, including literal percent sequences.
        string path = Uri.UnescapeDataString(uri.AbsolutePath);
        return _entries.TryGetValue(path, out ChmEntry? entry) ? entry : null;
    }

    /// <summary>Configures bounded rendering with this book's virtual base and archive-only resolver.</summary>
    public void ConfigureRenderOptions(HtmlRenderOptions options, string topicPath) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.BaseUri = GetTopicUri(topicPath);
        options.ResourceUrlPolicy = (options.ResourceUrlPolicy ?? options.UrlPolicy).Clone();
        options.ResourceUrlPolicy.AllowedUrlSchemes.Add("chm");
        options.ResourceResolver = CreateResourceResolver();
        options.SynchronousResourceResolver = null;
        options.MaxInputCharacters = Math.Min(options.MaxInputCharacters, _options.MaxEntryBytes);
        options.MaxHtmlNodes = Math.Min(options.MaxHtmlNodes, _options.MaxHtmlNodes);
        options.MaxLayoutDepth = Math.Min(options.MaxLayoutDepth, _options.MaxHtmlDepth);
    }

    /// <summary>Returns the conventional media type for a help-book resource path.</summary>
    public static string GetMediaType(string path) => System.IO.Path.GetExtension(path).ToLowerInvariant() switch {
        ".htm" or ".html" or ".hhc" or ".hhk" => "text/html",
        ".xhtml" => "application/xhtml+xml", ".css" => "text/css", ".js" => "text/javascript",
        ".png" => "image/png", ".jpg" or ".jpeg" => "image/jpeg", ".gif" => "image/gif", ".svg" => "image/svg+xml",
        ".bmp" => "image/bmp", ".webp" => "image/webp", ".ico" => "image/x-icon",
        ".ttf" => "font/ttf", ".otf" => "font/otf", ".woff" => "font/woff", ".woff2" => "font/woff2",
        ".txt" => "text/plain", ".xml" => "application/xml", ".pdf" => "application/pdf", _ => "application/octet-stream"
    };
}
