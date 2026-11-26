using OfficeIMO.Drawing;

namespace OfficeIMO.Epub;

// Keep Reader and the output adapters on the same spine, resource and navigation model.
internal static class EpubImportedPublication {
    private static readonly XNamespace Xhtml = "http://www.w3.org/1999/xhtml";
    internal static EpubDocument Open(EpubPublication publication, CancellationToken token) {
        using var package = new MemoryStream();
        var saveOptions = new EpubSaveOptions();
        publication.WriteStaged(package, saveOptions, token);
        package.Position = 0;
        EpubDocument document = EpubDocument.Load(package, new EpubLoadOptions {
            IncludeRawHtml = true,
            IncludeResourceData = true,
            ResourceDataFilter = static (_, mediaType) => IsRasterMediaType(mediaType),
            MaxResources = saveOptions.MaxEntries,
            IncludeNonLinearSpineItems = true,
            FallbackToHtmlScan = false,
            MaxChapters = 2_000,
            MaxChapterBytes = new EpubPublicationLoadOptions().MaxEntryBytes,
            MaxTotalRawHtmlBytes = saveOptions.MaxExpandedBytes,
            MaxResourceBytes = 32L * 1024 * 1024,
            MaxTotalResourceBytes = 192L * 1024 * 1024
        }, token);
        if (document.Chapters.Count != publication.Spine.Count || document.Chapters.Any(static chapter => chapter.Html == null)) {
            string? diagnostic = document.Diagnostics.FirstOrDefault(static item => item.Code.StartsWith("epub.chapter.", StringComparison.Ordinal))?.Message;
            throw new InvalidDataException(diagnostic ?? "The imported ebook could not retain all generated chapters.");
        }
        return document;
    }

    internal static void WriteHtml(EpubDocument book, Stream destination, List<OfficeDiagnostic> diagnostics,
        CancellationToken token, bool markdownAnchors = false) {
        var chapters = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (EpubChapter chapter in book.Chapters) {
            if (!chapters.ContainsKey(chapter.Path))
                chapters.Add(chapter.Path, "chapter-" + chapter.Order.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
        var images = new Dictionary<string, Lazy<string>>(StringComparer.Ordinal);
        foreach (EpubResource image in book.Resources) {
            if (image.Data is { Length: > 0 } bytes && image.MediaType is { } mediaType &&
                IsRasterMediaType(mediaType) && !images.ContainsKey(image.Path))
                images.Add(image.Path, new Lazy<string>(() => "data:" + mediaType + ";base64," + Convert.ToBase64String(bytes)));
        }
        using var writer = XmlWriter.Create(destination, new XmlWriterSettings {
            OmitXmlDeclaration = true, Encoding = new UTF8Encoding(false), CloseOutput = false
        });
        writer.WriteDocType("html", null, null, null);
        writer.WriteStartElement("html");
        writer.WriteAttributeString("lang", book.Language ?? "und");
        writer.WriteStartElement("head");
        writer.WriteStartElement("meta");
        writer.WriteAttributeString("charset", "utf-8");
        writer.WriteEndElement();
        writer.WriteElementString("title", book.Title);
        foreach ((string name, string value) in MetadataFields(book)) {
            token.ThrowIfCancellationRequested();
            writer.WriteStartElement("meta");
            writer.WriteAttributeString("name", name);
            writer.WriteAttributeString("content", value);
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
        writer.WriteStartElement("body");
        if (book.TableOfContents.Count > 0) {
            writer.WriteStartElement("nav");
            writer.WriteAttributeString("aria-label", "Table of contents");
            writer.WriteElementString("h2", "Contents");
            WriteNavigation(book.TableOfContents);
            writer.WriteEndElement();
        }
        foreach (EpubChapter chapter in book.Chapters) {
            token.ThrowIfCancellationRequested();
            writer.WriteStartElement("section");
            if (markdownAnchors) WriteAnchor(chapters[chapter.Path]);
            else writer.WriteAttributeString("id", chapters[chapter.Path]);
            if (chapter.Html is { Length: > 0 } markup) {
                using var input = XmlReader.Create(new StringReader(markup), new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Ignore, XmlResolver = null
                });
                XDocument content = XDocument.Load(input, LoadOptions.PreserveWhitespace);
                XElement? body = content.Root?.Element(Xhtml + "body");
                foreach (XNode node in body?.Nodes() ?? Enumerable.Empty<XNode>()) WriteNode(node, chapter.Path);
            } else {
                writer.WriteElementString("p", chapter.Text);
            }
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
        writer.WriteEndElement();

        void WriteNavigation(IReadOnlyList<EpubNavigationItem> items) {
            writer.WriteStartElement("ol");
            foreach (EpubNavigationItem item in items) {
                token.ThrowIfCancellationRequested();
                writer.WriteStartElement("li");
                if (item.Target is { } target && chapters.TryGetValue(target, out string? chapterId)) {
                    writer.WriteStartElement("a");
                    writer.WriteAttributeString("href", "#" + Anchor(chapterId, item.Fragment));
                    writer.WriteString(item.Label);
                    writer.WriteEndElement();
                } else writer.WriteString(item.Label);
                if (item.Children.Count > 0) WriteNavigation(item.Children);
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
        }

        void WriteNode(XNode node, string chapterPath) {
            token.ThrowIfCancellationRequested();
            if (node is XText text) {
                writer.WriteString(text.Value);
                return;
            }
            if (node is not XElement element) return;
            string tag = element.Name.LocalName;
            if (element.Name.Namespace != Xhtml || !IsStaticElement(tag)) {
                Omit("EBOOK_HTML_ELEMENT_OMITTED", "An element is outside the static HTML projection.", tag);
                return;
            }
            if (markdownAnchors && (string?)element.Attribute("id") is { } id)
                WriteAnchor(Anchor(chapters[chapterPath], id));
            writer.WriteStartElement(tag);
            foreach (XAttribute attribute in element.Attributes()) {
                if (attribute.IsNamespaceDeclaration) continue;
                string name = attribute.Name.LocalName;
                string? value = attribute.Value;
                if ((attribute.Name.Namespace != XNamespace.None && attribute.Name != XNamespace.Xml + "lang") ||
                    name is not ("id" or "lang" or "title" or "alt" or "src" or "href" or "width" or "height" or "colspan" or "rowspan")) {
                    Omit("EBOOK_HTML_ATTRIBUTE_OMITTED", "An attribute is outside the static HTML projection.", name);
                    continue;
                }
                if ((name == "src" && tag != "img") || (name == "href" && tag != "a")) continue;
                if (name == "id") {
                    if (markdownAnchors) continue;
                    value = Anchor(chapters[chapterPath], value);
                }
                if (name is "src" or "href") {
                    EpubReference reference = EpubReference.Resolve(chapterPath, value);
                    if (name == "src" && reference.ContainerPath is { } path && images.TryGetValue(path, out var data))
                        value = data.Value;
                    else if (name == "href" && reference.ContainerPath is { } target && chapters.TryGetValue(target, out string? destinationId))
                        value = "#" + Anchor(destinationId, reference.Fragment);
                    else if (name != "href" || !IsWebLink(value)) {
                        Omit("EBOOK_HTML_REFERENCE_OMITTED", "The reference could not be included in the standalone HTML.", value);
                        value = null;
                    }
                }
                if (value != null) writer.WriteAttributeString(name, value);
            }
            foreach (XNode child in element.Nodes()) WriteNode(child, chapterPath);
            // XML empty-element syntax is inappropriate for non-void elements in text/html.
            if (element.Name.LocalName is "img" or "br" or "hr" or "wbr") writer.WriteEndElement();
            else writer.WriteFullEndElement();
        }

        void WriteAnchor(string id) {
            writer.WriteStartElement("a");
            writer.WriteAttributeString("id", id);
            writer.WriteFullEndElement();
        }

        void Omit(string code, string message, string location) {
            if (diagnostics.Count < 1_000) diagnostics.Add(new OfficeDiagnostic(code,
                OfficeDiagnosticSeverity.Warning, message, location, "OfficeIMO.Epub", OfficeConversionLossKind.Omission));
        }
    }

    private static bool IsStaticElement(string tag) => tag is
        "p" or "div" or "span" or "section" or "article" or "aside" or "header" or "footer" or "nav" or
        "h1" or "h2" or "h3" or "h4" or "h5" or "h6" or "a" or "br" or "hr" or "wbr" or "blockquote" or "ul" or "ol" or
        "li" or "em" or "strong" or "b" or "i" or "u" or "s" or "sub" or "sup" or "code" or "pre" or "dfn" or "cite" or
        "table" or "thead" or "tbody" or "tfoot" or "tr" or "td" or "th" or "caption" or "img" or "figure" or "figcaption" or
        "dl" or "dt" or "dd" or "q";

    private static bool IsRasterMediaType(string? mediaType) =>
        mediaType is "image/jpeg" or "image/png" or "image/gif" or "image/webp";

    internal static bool TryReadRaster(byte[] data, string? name, CancellationToken token, out OfficeImageInfo image) =>
        OfficeImageReader.TryIdentifyByContent(data, name, token, out image) &&
        image.Format is (OfficeImageFormat.Png or OfficeImageFormat.Jpeg or OfficeImageFormat.Gif or OfficeImageFormat.Webp) &&
        image.Width > 0 && image.Height > 0 && (long)image.Width * image.Height <= 40_000_000 &&
        OfficeImageReader.TryValidateContent(data, name, token, out image);

    internal static IEnumerable<(string Name, string Value)> MetadataFields(EpubDocument book) {
        foreach (EpubMetadataEntry item in book.Metadata) {
            if (string.IsNullOrWhiteSpace(item.Value)) continue;
            if (item.Kind == EpubMetadataKind.DublinCore)
                yield return ("dc." + item.Name, item.Value);
            else if (item.Property is "belongs-to-collection" or "group-position" or "collection-type")
                yield return ("epub." + item.Property, item.Value);
        }
    }

    private static string Anchor(string chapter, string? fragment) =>
        string.IsNullOrEmpty(fragment) ? chapter : chapter + "-" + fragment;

    internal static bool IsWebLink(string value) =>
        Uri.TryCreate(value, UriKind.Absolute, out Uri? uri) && uri.Scheme is "http" or "https" or "mailto";
}
