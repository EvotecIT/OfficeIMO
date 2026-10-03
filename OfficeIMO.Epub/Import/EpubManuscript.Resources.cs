using AngleSharp.Html.Dom;
using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Epub;

public static partial class EpubManuscript {
    private sealed class ImportedStylesheet {
        internal string Id = string.Empty;
        internal Dictionary<string, string> Attributes = new Dictionary<string, string>();
    }

    private static async Task<List<ImportedStylesheet>> CollectResourcesAsync(IHtmlDocument source, List<Chapter> chapters,
        EpubPublication publication, HtmlConversionDocument manuscript, EpubManuscriptOptions options,
        List<OfficeConversionFidelityDiagnostic> diagnostics, CancellationToken token) {
        var pipelineOptions = new HtmlResourcePipelineOptions {
            BaseUri = manuscript.BaseUri, ResourceUrlPolicy = manuscript.ResourceUrlPolicy, Limits = manuscript.Limits
        };
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildArchiveManifest(source, pipelineOptions);
        var resourceOptions = new HtmlRenderOptions {
            ResourceResolver = options.ResourceResolver, ResourceUrlPolicy = manuscript.ResourceUrlPolicy,
            MaxResourceBytes = options.MaxResourceBytes, MaxTotalResourceBytes = options.MaxTotalResourceBytes,
            MaxResourceCount = options.MaxResourceCount, MaxResourceRequests = Math.Max(512, options.MaxResourceCount),
            MaxConcurrentResourceLoads = 1
        };
        var resourceDiagnostics = new HtmlDiagnosticReport();
        HtmlResourceSession session = await HtmlRenderResourceLoader.LoadArchiveAsync(manifest, resourceOptions, resourceDiagnostics,
            manuscript.Limits, token).ConfigureAwait(false);
        foreach (HtmlDiagnostic diagnostic in resourceDiagnostics.Diagnostics)
            diagnostics.Add(new OfficeConversionFidelityDiagnostic(diagnostic.Code, diagnostic.Message,
                diagnostic.LossKind == OfficeConversionLossKind.None ? OfficeConversionLossKind.Omission : diagnostic.LossKind,
                diagnostic.Component, diagnostic.Source));
        var paths = new Dictionary<HtmlResolvedResource, string>();
        var mediaTypes = new Dictionary<HtmlResolvedResource, string>();
        var declarations = new Dictionary<string, string>(StringComparer.Ordinal);
        var deduplicated = new Dictionary<string, string>(StringComparer.Ordinal);
        int count = 0;
        foreach (HtmlResourceSessionEntry entry in session.Resources) {
            token.ThrowIfCancellationRequested();
            if (!session.TryGet(entry.Source, entry.CanonicalSource, out HtmlResolvedResource resource)) continue;
            string mediaType = entry.ContentType.Split(';')[0].Trim().ToLowerInvariant();
            if (entry.Kind == HtmlResourceKind.Stylesheet) mediaType = "text/css";
            string key = mediaType + ":" + entry.Sha256 + (mediaType == "text/css" || mediaType == "image/svg+xml" ? ":" + entry.CanonicalSource : string.Empty);
            if (!deduplicated.TryGetValue(key, out string? path)) {
                path = "EPUB/resources/resource-" + (++count).ToString("D4") + Extension(mediaType);
                deduplicated.Add(key, path);
            }
            paths[resource] = path; mediaTypes[resource] = mediaType;
        }
        foreach (HtmlResourceSessionEntry entry in session.Resources) {
            token.ThrowIfCancellationRequested();
            if (!session.TryGet(entry.Source, entry.CanonicalSource, out HtmlResolvedResource resource)) continue;
            string path = paths[resource];
            if (declarations.ContainsKey(path)) continue;
            byte[] bytes = resource.Bytes;
            string mediaType = mediaTypes[resource];
            if (mediaType == "text/css") {
                if (!HtmlResourcePipeline.TryDecodeStylesheet(bytes, resource.ContentType, out string css)) {
                    AddDiagnostic(diagnostics, "EPUB_IMPORT_CSS_ENCODING_FAILED", "A stylesheet could not be decoded.", entry.CanonicalSource, OfficeConversionLossKind.Failure);
                    continue;
                }
                css = HtmlResourcePipeline.RewriteCssResourceUrls(css, (url, kind) => Rewrite(url, new Uri(entry.CanonicalSource), path, kind));
                bytes = new UTF8Encoding(false, true).GetBytes(css);
            }
            if (mediaType == "image/svg+xml") {
                try {
                    using var svgStream = new MemoryStream(bytes, false);
                    using var svgReader = XmlReader.Create(svgStream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = options.MaxResourceBytes });
                    var svg = XDocument.Load(svgReader, LoadOptions.PreserveWhitespace);
                    foreach (XElement element in svg.Descendants()) {
                        foreach (XAttribute attribute in element.Attributes().ToArray()) {
                            if (attribute.Name.LocalName == "style") attribute.Value = HtmlResourcePipeline.RewriteCssResourceUrls(attribute.Value,
                                (url, kind) => Rewrite(url, new Uri(entry.CanonicalSource), path, kind));
                            else if (attribute.Name.LocalName is "href" or "src" && element.Name.LocalName != "a")
                                attribute.Value = Rewrite(attribute.Value, new Uri(entry.CanonicalSource), path, HtmlResourceKind.Image);
                        }
                        if (element.Name.LocalName == "style") element.Value = HtmlResourcePipeline.RewriteCssResourceUrls(element.Value,
                            (url, kind) => Rewrite(url, new Uri(entry.CanonicalSource), path, kind));
                    }
                    bytes = new UTF8Encoding(false, true).GetBytes(svg.ToString(SaveOptions.DisableFormatting));
                } catch (XmlException error) {
                    AddDiagnostic(diagnostics, "EPUB_IMPORT_SVG_INVALID", error.Message, entry.CanonicalSource, OfficeConversionLossKind.Failure);
                    continue;
                }
            }
            string id = "manuscript-resource-" + (declarations.Count + 1);
            publication.AddResource(id, path, mediaType, bytes);
            declarations.Add(path, id);
        }
        var styles = new List<ImportedStylesheet>();
        int styleIndex = 0;
        foreach (var style in source.QuerySelectorAll("link[href],style")) {
            token.ThrowIfCancellationRequested();
            if (style.LocalName == "link") {
                if (HtmlResourcePipeline.GetLinkResourceKind(style.GetAttribute("rel"), style.GetAttribute("as")) != HtmlResourceKind.Stylesheet) continue;
                string resourcePath = FindPath(style.GetAttribute("href")!, manuscript.BaseUri, HtmlResourceKind.Stylesheet);
                if (declarations.TryGetValue(resourcePath, out string? resourceId)) {
                    var imported = new ImportedStylesheet { Id = resourceId };
                    foreach (string attribute in new[] { "rel", "media", "title" }) {
                        string? value = style.GetAttribute(attribute);
                        if (value != null) imported.Attributes[attribute] = value;
                    }
                    styles.Add(imported);
                }
                continue;
            }
            string css = style.TextContent;
            string media = style.GetAttribute("media") ?? string.Empty;
            string path = "EPUB/styles/source-" + (++styleIndex).ToString("D4") + ".css";
            css = HtmlResourcePipeline.RewriteCssResourceUrls(css, (url, kind) => Rewrite(url, manuscript.BaseUri, path, kind));
            if (media.Length != 0) css = "@media " + media + "{" + css + "}";
            string id = "manuscript-style-" + styleIndex;
            publication.AddStylesheet(id, path, css); styles.Add(new ImportedStylesheet { Id = id });
        }
        foreach (Chapter chapter in chapters) {
            foreach (XElement element in chapter.Body.DescendantsAndSelf()) {
                token.ThrowIfCancellationRequested();
                foreach (XAttribute attribute in element.Attributes().ToArray()) {
                    string name = attribute.Name.LocalName;
                    if (name == "style") attribute.Value = HtmlResourcePipeline.RewriteCssResourceUrls(attribute.Value,
                        (url, kind) => Rewrite(url, manuscript.BaseUri, chapter.Path, kind));
                    else if (name == "srcset") {
                        var candidates = HtmlSrcSetParser.Parse(attribute.Value, manuscript.Limits.MaxResponsiveImageCandidates);
                        attribute.Value = string.Join(", ", candidates.Select(candidate => Rewrite(candidate.Url, manuscript.BaseUri, chapter.Path, HtmlResourceKind.Image) + " " + candidate.Descriptor));
                    } else if (name == "src" || name == "poster" || name == "background" ||
                        name == "href" && element.Name.NamespaceName == "http://www.w3.org/2000/svg" && element.Name.LocalName != "a") {
                        if (attribute.Value.StartsWith("#", StringComparison.Ordinal)) continue;
                        string rewritten = Rewrite(attribute.Value, manuscript.BaseUri, chapter.Path,
                            name == "poster" || name == "background" || new[] { "img", "image", "use", "feImage" }.Contains(element.Name.LocalName) || element.Parent?.Name == Xhtml + "picture"
                                ? HtmlResourceKind.Image : HtmlResourceKind.Media);
                        if (rewritten.Length == 0) attribute.Remove(); else attribute.Value = rewritten;
                    }
                }
            }
        }
        return styles;

        string FindPath(string url, Uri? baseUri, HtmlResourceKind kind) {
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(url, baseUri, manuscript.ResourceUrlPolicy);
            if (session.TryGet(null, resolved, out HtmlResolvedResource resource) && paths.TryGetValue(resource, out string? path)) return path;
            AddDiagnostic(diagnostics, "EPUB_IMPORT_RESOURCE_MISSING", "A manuscript dependency could not be collected. Supply an authorized resolver or repair the source reference.", url, OfficeConversionLossKind.Failure);
            return string.Empty;
        }
        string Rewrite(string url, Uri? baseUri, string owner, HtmlResourceKind kind) {
            if (url.StartsWith("#", StringComparison.Ordinal)) return url;
            string path = FindPath(url, baseUri, kind);
            if (path.Length == 0) return string.Empty;
            string fragment = Uri.TryCreate(url, UriKind.RelativeOrAbsolute, out Uri? uri) && uri.IsAbsoluteUri ? uri.Fragment :
                url.IndexOf('#') >= 0 ? url.Substring(url.IndexOf('#')) : string.Empty;
            return new Uri("epub://package/" + owner).MakeRelativeUri(new Uri("epub://package/" + path)).OriginalString + fragment;
        }
    }

    private static string Extension(string mediaType) => mediaType switch {
        "text/css" => ".css", "image/png" => ".png", "image/jpeg" => ".jpg", "image/gif" => ".gif", "image/svg+xml" => ".svg",
        "image/webp" => ".webp", "image/avif" => ".avif", "font/woff" => ".woff", "font/woff2" => ".woff2",
        "font/ttf" => ".ttf", "font/otf" => ".otf", "audio/mpeg" => ".mp3", "video/mp4" => ".mp4", "text/vtt" => ".vtt", _ => ".bin"
    };
}
