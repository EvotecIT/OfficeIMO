using AngleSharp.Html.Dom;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    internal static HtmlResourceManifest BuildSvgArchiveManifest(byte[] bytes, Uri uri, HtmlResourcePipelineOptions options) {
        using var stream = new MemoryStream(bytes, false);
        using var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings {
            DtdProcessing = System.Xml.DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = Math.Max(1, bytes.LongLength)
        });
        var svg = System.Xml.Linq.XElement.Load(reader, System.Xml.Linq.LoadOptions.PreserveWhitespace);
        if (svg.Name != System.Xml.Linq.XName.Get("svg", "http://www.w3.org/2000/svg"))
            throw new InvalidDataException("SVG resource requires an SVG document root.");
        HtmlConversionInputGuard.ValidateSource(svg.ToString(System.Xml.Linq.SaveOptions.DisableFormatting), options.Limits);
        System.Xml.Linq.XNamespace xhtml = Dom.HtmlElement.HtmlNamespace;
        var document = HtmlXmlDocumentParser.CreateDocument(new System.Xml.Linq.XDocument(
            new System.Xml.Linq.XElement(xhtml + "html", new System.Xml.Linq.XElement(xhtml + "body", svg))),
            options.Limits, CancellationToken.None);
        return BuildArchiveManifest(document, new HtmlResourcePipelineOptions {
            BaseUri = uri, ResourceUrlPolicy = options.ResourceUrlPolicy, Limits = options.Limits, MediaContext = options.MediaContext
        });
    }
    /// <summary>Rewrites stylesheet imports, URL functions and image-set strings without discarding media, supports or layer conditions.</summary>
    /// <param name="css">Stylesheet or inline declaration source.</param>
    /// <param name="rewrite">Maps a decoded source URI and its resource kind to a replacement URI.</param>
    /// <returns>CSS with changed resource carriers rewritten; source outside those carriers is retained.</returns>
    public static string RewriteCssResourceUrls(string css, Func<string, HtmlResourceKind, string> rewrite) =>
        RewriteCssResourceUrls(css, rewrite, includeFragmentReferences: false);

    /// <summary>Rewrites CSS resource URLs, optionally including local fragment references when relocating document content.</summary>
    public static string RewriteCssResourceUrls(string css, Func<string, HtmlResourceKind, string> rewrite, bool includeFragmentReferences) {
        if (css == null) throw new ArgumentNullException(nameof(css));
        if (rewrite == null) throw new ArgumentNullException(nameof(rewrite));
        // Analyze a length-preserving mask, then edit the original source. Comments can
        // separate CSS tokens without introducing a descendant combinator or whitespace.
        string normalized = MaskCssComments(css).Replace(CssCommentMask, ' ');
        var replacements = new List<(int Start, int End, string Value)>();
        var imports = ExtractCssImports(normalized).ToArray();
        foreach (CssImportReference import in imports) {
            string source = DecodeCssEscapes(import.Source);
            string value = rewrite(source, HtmlResourceKind.Stylesheet);
            if (value == source) continue;
            replacements.Add(value.Length == 0 ? (import.Start, import.End, string.Empty) :
                (import.SourceStart, import.SourceEnd, "url(" + Quote(value) + ")"));
        }
        foreach (Match match in CssUrlExpression.Matches(normalized)) {
            if (!IsValidCssUrlMatch(normalized, match) || !IsCssFunctionNameAt(normalized, match.Index, "url") ||
                IsInsideCssString(normalized, match.Index) || imports.Any(import => match.Index >= import.Start && match.Index < import.End)) continue;
            string source = DecodeCssEscapes(match.Groups["url"].Value.Trim().Trim('\'', '"'));
            if (!includeFragmentReferences && IsFragmentOnlyReference(source)) continue;
            string value = rewrite(source, ClassifyCssUrl(normalized, match.Index));
            if (value != source) replacements.Add((match.Index, match.Index + match.Length, "url(" + Quote(value) + ")"));
        }
        foreach (CssStringUrlReference image in ExtractImageSetStringUrls(normalized)) {
            if (!includeFragmentReferences && IsFragmentOnlyReference(image.Source)) continue;
            string source = DecodeCssEscapes(image.Source);
            string value = rewrite(source, HtmlResourceKind.Image);
            if (value != source) replacements.Add((image.SourceStart - 1, image.End, Quote(value)));
        }
        var output = new StringBuilder(css);
        int lastStart = normalized.Length;
        foreach (var item in replacements.OrderByDescending(item => item.Start)) {
            if (item.End > lastStart) continue;
            output.Remove(item.Start, item.End - item.Start).Insert(item.Start, item.Value);
            lastStart = item.Start;
        }
        return output.ToString();

        static string Quote(string value) => "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\d ").Replace("\n", "\\a ").Replace("\f", "\\c ") + "\"";
    }

    internal static HtmlResourceManifest BuildArchiveManifest(IHtmlDocument document, HtmlResourcePipelineOptions options) {
        HtmlResourceManifest manifest = BuildManifest(document, options);
        Uri? baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(document, options.BaseUri);
        foreach (var element in document.All) {
            string name = element.LocalName;
            if (name == "link" && GetLinkResourceKind(element.GetAttribute("rel"), element.GetAttribute("as")) == HtmlResourceKind.Stylesheet)
                AddAttribute(manifest, HtmlResourceKind.Stylesheet, element, "href", baseUri, options);
            if (name == "source") {
                AddAttribute(manifest, element.ParentElement?.LocalName == "picture" ? HtmlResourceKind.Image : HtmlResourceKind.Media, element, "src", baseUri, options);
                AddSrcSet(manifest, HtmlResourceKind.Image, element, "srcset", baseUri, options);
            }
            if (name == "img") {
                AddAttribute(manifest, HtmlResourceKind.Image, element, "src", baseUri, options);
                AddSrcSet(manifest, HtmlResourceKind.Image, element, "srcset", baseUri, options);
            }
        }
        Uri cssBase = baseUri ?? new Uri("officeimo-unresolved://manuscript/");
        foreach (string css in document.QuerySelectorAll("style").Select(element => element.TextContent)
            .Concat(document.QuerySelectorAll("[style]").Select(element => element.GetAttribute("style") ?? string.Empty))) {
            AddStylesheetArchiveResources(manifest, css, cssBase, options);
        }
        return manifest;
    }

    internal static void AddStylesheetArchiveResources(HtmlResourceManifest manifest, string css, Uri baseUri, HtmlResourcePipelineOptions options) {
        var analysis = AnalyzeExternalStylesheet(css, baseUri, options, includeInactiveResources: true);
        foreach (HtmlResourceReference resource in analysis.Imports.Select(import => import.Reference).Concat(analysis.FontResources).Concat(analysis.ImageResources)) manifest.Add(resource);
    }
}
