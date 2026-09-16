using OfficeIMO.Html.Runtime;

namespace OfficeIMO.Html.Runtime.Rendering;

// Static resource planning for an offline application snapshot. HTML and CSS
// parsing stay with HtmlResourcePipeline; callers own transport and authority.
internal sealed class HtmlApplicationResourceDiscovery {
    private const int MaximumDiscoveredUrls = 128;
    private readonly HashSet<string> _seen = new(StringComparer.Ordinal);
    private readonly HashSet<string> _stylesheetUrls = new(StringComparer.Ordinal);
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
        AppendStylesheets(supplied, pending);
        return pending.ToArray();
    }

    internal string[] DiscoverStylesheets(IReadOnlyList<HtmlRuntimeResource> resources) {
        ArgumentNullException.ThrowIfNull(resources);
        foreach (HtmlRuntimeResource resource in resources) _seen.Add(Key(resource.Url));
        var pending = new List<string>();
        AppendStylesheets(resources, pending);
        return pending.ToArray();
    }

    private void AppendStylesheets(IReadOnlyList<HtmlRuntimeResource> resources, List<string> pending) {
        foreach (HtmlRuntimeResource resource in resources) {
            if ((!_stylesheetUrls.Contains(Key(resource.Url)) &&
                 !resource.ContentType.StartsWith("text/css", StringComparison.OrdinalIgnoreCase)) ||
                resource.StatusCode is < 200 or >= 300 || resource.Length == 0)
                continue;
            if (!HtmlResourcePipeline.TryDecodeStylesheet(resource.Content, resource.ContentType, out string css))
                throw new HtmlScriptRuntimeException("A discovered stylesheet uses an unsupported encoding.");
            Append(HtmlResourcePipeline.BuildStylesheetManifest(css, resource.FinalUrl, _screenOptions), pending);
            Append(HtmlResourcePipeline.BuildStylesheetManifest(css, resource.FinalUrl, _printOptions), pending);
        }
    }

    private void Append(HtmlResourceManifest manifest, List<string> pending) {
        foreach (HtmlResourceReference reference in manifest.Resources) {
            if (!reference.IsAllowed || reference.Kind is not (HtmlResourceKind.Script or HtmlResourceKind.Stylesheet
                or HtmlResourceKind.Image or HtmlResourceKind.Font) ||
                !Uri.TryCreate(reference.ResolvedSource, UriKind.Absolute, out Uri? url) ||
                url.Scheme is not ("http" or "https")) continue;
            string key = Key(url);
            if (reference.Kind == HtmlResourceKind.Stylesheet) _stylesheetUrls.Add(key);
            if (!_seen.Add(key)) continue;
            if (_seen.Count > MaximumDiscoveredUrls)
                throw new HtmlScriptRuntimeException("The document exceeds the pilot's static resource discovery limit.");
            pending.Add(key);
        }
    }

    private static string Key(Uri url) => new UriBuilder(url) { Fragment = string.Empty }.Uri.AbsoluteUri;
}
