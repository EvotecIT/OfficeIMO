namespace OfficeIMO.Epub;

// Keep Reader and the output adapters on the same spine, resource and navigation model.
internal static class EpubImportedPublication {
    internal static EpubDocument Open(EpubPublication publication, CancellationToken token) {
        using var package = new MemoryStream();
        publication.WriteStaged(package, new EpubSaveOptions(), token);
        package.Position = 0;
        return EpubDocument.Load(package, new EpubLoadOptions {
            IncludeRawHtml = true,
            IncludeResourceData = true,
            IncludeNonLinearSpineItems = true,
            FallbackToHtmlScan = false,
            MaxChapters = 2_000,
            MaxChapterBytes = 32L * 1024 * 1024,
            MaxResourceBytes = 32L * 1024 * 1024,
            MaxTotalResourceBytes = 192L * 1024 * 1024
        }, token);
    }

    internal static void WriteHtml(EpubDocument book, Stream destination, List<OfficeDiagnostic> diagnostics,
        CancellationToken token) {
        var chapters = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (EpubChapter chapter in book.Chapters) {
            if (!chapters.ContainsKey(chapter.Path))
                chapters.Add(chapter.Path, "chapter-" + chapter.Order.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
        var images = new Dictionary<string, Lazy<string>>(StringComparer.Ordinal);
        foreach (EpubResource image in book.Resources) {
            if (image.Data is { Length: > 0 } bytes && image.MediaType is { } mediaType &&
                mediaType is "image/jpeg" or "image/png" or "image/gif" or "image/webp" && !images.ContainsKey(image.Path))
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
        foreach (EpubMetadataEntry metadata in book.Metadata) {
            token.ThrowIfCancellationRequested();
            if (metadata.Kind != EpubMetadataKind.DublinCore || string.IsNullOrWhiteSpace(metadata.Value)) continue;
            writer.WriteStartElement("meta");
            writer.WriteAttributeString("name", "dc." + metadata.Name);
            writer.WriteAttributeString("content", metadata.Value);
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
            writer.WriteAttributeString("id", chapters[chapter.Path]);
            if (chapter.Html is { Length: > 0 } markup) {
                using var input = XmlReader.Create(new StringReader(markup), new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Ignore, XmlResolver = null
                });
                XDocument content = XDocument.Load(input, LoadOptions.PreserveWhitespace);
                XElement? body = content.Root?.Element(XName.Get("body", "http://www.w3.org/1999/xhtml"));
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
            writer.WriteStartElement(element.Name.LocalName);
            foreach (XAttribute attribute in element.Attributes()) {
                if (attribute.IsNamespaceDeclaration) continue;
                string name = attribute.Name.LocalName;
                string? value = attribute.Value;
                if (name == "id") value = Anchor(chapters[chapterPath], value);
                if (name is "src" or "href") {
                    EpubReference reference = EpubReference.Resolve(chapterPath, value);
                    if (name == "src" && reference.ContainerPath is { } path && images.TryGetValue(path, out var data))
                        value = data.Value;
                    else if (name == "href" && reference.ContainerPath is { } target && chapters.TryGetValue(target, out string? id))
                        value = "#" + Anchor(id, reference.Fragment);
                    else if (name != "href" || !IsWebLink(value)) {
                        if (diagnostics.Count < 1_000) diagnostics.Add(new OfficeDiagnostic("EBOOK_HTML_REFERENCE_OMITTED",
                            OfficeDiagnosticSeverity.Warning, "The reference could not be included in the standalone HTML.", value,
                            "OfficeIMO.Epub", OfficeConversionLossKind.Omission));
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
    }

    private static string Anchor(string chapter, string? fragment) =>
        string.IsNullOrEmpty(fragment) ? chapter : chapter + "-" + fragment;

    internal static bool IsWebLink(string value) =>
        Uri.TryCreate(value, UriKind.Absolute, out Uri? uri) && uri.Scheme is "http" or "https" or "mailto";
}
