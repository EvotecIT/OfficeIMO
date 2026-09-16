using OfficeIMO.Html.Runtime;
namespace OfficeIMO.Html.Runtime.Rendering;

// Static resource planning for an offline application snapshot. HTML and CSS
// parsing stay with HtmlResourcePipeline; callers own transport and authority.
internal sealed class HtmlApplicationResourceDiscovery {
    private const int MaximumDiscoveredUrls = 128;
    private readonly HashSet<string> _seen = new(StringComparer.Ordinal);
    private readonly HashSet<string> _stylesheetUrls = new(StringComparer.Ordinal);
    private readonly HashSet<string> _processedStylesheetUrls = new(StringComparer.Ordinal);
    private readonly HashSet<string> _documentUrls = new(StringComparer.Ordinal);
    private readonly HashSet<string> _processedDocumentUrls = new(StringComparer.Ordinal);
    private readonly HtmlResourcePipelineOptions _screenOptions = new() {
        UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
        ResourceUrlPolicy = HtmlUrlPolicy.CreateWebResourceProfile(),
        MediaContext = HtmlCssMediaContext.Screen,
        MediaWidth = 816D,
        MediaHeight = 720D
    };
    private readonly HtmlResourcePipelineOptions _printOptions = new() {
        UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
        ResourceUrlPolicy = HtmlUrlPolicy.CreateWebResourceProfile(),
        MediaContext = HtmlCssMediaContext.Print,
        MediaWidth = 793.7D,
        MediaHeight = 1122.5D
    };

    internal HtmlApplicationResourceDiscovery(double viewportWidth = 816D, double viewportHeight = 720D,
        double devicePixelRatio = 1D) {
        if (!double.IsFinite(viewportWidth) || viewportWidth <= 0D) throw new ArgumentOutOfRangeException(nameof(viewportWidth));
        if (!double.IsFinite(viewportHeight) || viewportHeight <= 0D) throw new ArgumentOutOfRangeException(nameof(viewportHeight));
        if (!double.IsFinite(devicePixelRatio) || devicePixelRatio <= 0D) throw new ArgumentOutOfRangeException(nameof(devicePixelRatio));
        _screenOptions.MediaWidth = viewportWidth;
        _screenOptions.MediaHeight = viewportHeight;
        _screenOptions.DevicePixelRatio = devicePixelRatio;
        _screenOptions.MediaFeatures.ResolutionDpi = devicePixelRatio * HtmlRenderOptions.CssPixelsPerInch;
        _printOptions.DevicePixelRatio = devicePixelRatio;
        _printOptions.MediaFeatures.ResolutionDpi = devicePixelRatio * HtmlRenderOptions.CssPixelsPerInch;
    }

    internal string[] DiscoverDocument(string html, Uri documentUrl, IReadOnlyList<HtmlRuntimeResource> supplied) {
        ArgumentNullException.ThrowIfNull(html);
        ArgumentNullException.ThrowIfNull(documentUrl);
        ArgumentNullException.ThrowIfNull(supplied);
        _screenOptions.BaseUri = documentUrl;
        _printOptions.BaseUri = documentUrl;
        _seen.Add(Key(documentUrl));
        foreach (HtmlRuntimeResource resource in supplied) _seen.Add(Key(resource.Url));
        var pending = new List<string>();
        Append(HtmlResourcePipeline.BuildManifest(html, _screenOptions), pending);
        Append(HtmlResourcePipeline.BuildManifest(html, _printOptions), pending);
        AppendSuppliedResources(supplied, pending);
        return pending.ToArray();
    }

    internal string[] DiscoverResources(IReadOnlyList<HtmlRuntimeResource> resources) {
        ArgumentNullException.ThrowIfNull(resources);
        foreach (HtmlRuntimeResource resource in resources) _seen.Add(Key(resource.Url));
        var pending = new List<string>();
        AppendSuppliedResources(resources, pending);
        return pending.ToArray();
    }

    private void AppendSuppliedResources(IReadOnlyList<HtmlRuntimeResource> resources, List<string> pending) {
        bool processed;
        do {
            processed = AppendStylesheets(resources, pending);
            processed |= AppendDocuments(resources, pending);
        } while (processed);
    }

    private bool AppendDocuments(IReadOnlyList<HtmlRuntimeResource> resources, List<string> pending) {
        bool processed = false;
        foreach (HtmlRuntimeResource resource in resources) {
            string requestedKey = Key(resource.Url);
            if (!_documentUrls.Contains(requestedKey) || !_processedDocumentUrls.Add(requestedKey)) continue;
            processed = true;
            if (resource.StatusCode is < 200 or >= 300)
                throw new HtmlScriptRuntimeException("A discovered frame document response was not successful.");
            if (resource.Length == 0)
                throw new HtmlScriptRuntimeException("A discovered frame document response was empty.");
            string html = HtmlPublicResourceBroker.DecodeUtf8Html(resource, allowXhtml: false);
            _screenOptions.BaseUri = resource.FinalUrl;
            _printOptions.BaseUri = resource.FinalUrl;
            Append(HtmlResourcePipeline.BuildManifest(html, _screenOptions), pending);
            Append(HtmlResourcePipeline.BuildManifest(html, _printOptions), pending);
        }
        return processed;
    }

    private bool AppendStylesheets(IReadOnlyList<HtmlRuntimeResource> resources, List<string> pending) {
        bool processed = false;
        foreach (HtmlRuntimeResource resource in resources) {
            string requestedKey = Key(resource.Url);
            if ((!_stylesheetUrls.Contains(requestedKey) &&
                 !resource.ContentType.StartsWith("text/css", StringComparison.OrdinalIgnoreCase)) ||
                !_processedStylesheetUrls.Add(requestedKey) ||
                resource.StatusCode is < 200 or >= 300 || resource.Length == 0)
                continue;
            processed = true;
            if (!HtmlResourcePipeline.TryDecodeStylesheet(resource.Content, resource.ContentType, out string css))
                throw new HtmlScriptRuntimeException("A discovered stylesheet uses an unsupported encoding.");
            Append(HtmlResourcePipeline.BuildStylesheetManifest(css, resource.FinalUrl, _screenOptions), pending);
            Append(HtmlResourcePipeline.BuildStylesheetManifest(css, resource.FinalUrl, _printOptions), pending);
        }
        return processed;
    }

    private void Append(HtmlResourceManifest manifest, List<string> pending) {
        foreach (HtmlResourceReference reference in manifest.Resources) {
            bool frameDocument = reference.Kind == HtmlResourceKind.Other &&
                reference.ElementName.Equals("iframe", StringComparison.OrdinalIgnoreCase) &&
                reference.AttributeName.Equals("src", StringComparison.OrdinalIgnoreCase);
            if (!reference.IsAllowed || (!frameDocument && reference.Kind is not (HtmlResourceKind.Script or HtmlResourceKind.Stylesheet
                or HtmlResourceKind.Image or HtmlResourceKind.Font)) ||
                !Uri.TryCreate(reference.ResolvedSource, UriKind.Absolute, out Uri? url) ||
                url.Scheme is not ("http" or "https")) continue;
            string key = Key(url);
            if (reference.Kind == HtmlResourceKind.Stylesheet) _stylesheetUrls.Add(key);
            if (frameDocument) _documentUrls.Add(key);
            if (!_seen.Add(key)) continue;
            if (_seen.Count > MaximumDiscoveredUrls)
                throw new HtmlScriptRuntimeException("The document exceeds the pilot's static resource discovery limit.");
            pending.Add(key);
        }
    }

    private static string Key(Uri url) => new UriBuilder(url) { Fragment = string.Empty }.Uri.AbsoluteUri;
}
