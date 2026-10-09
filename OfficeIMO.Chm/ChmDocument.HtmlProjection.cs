using OfficeIMO.Html.Dom;
using System.Globalization;

namespace OfficeIMO.Chm;

public sealed partial class ChmDocument {
    /// <summary>Projects selected topics into one inert linked HTML book, with topic boundaries and a fidelity report.</summary>
    /// <remarks>Archive images can be embedded. Other resources retain private CHM URIs and require this document's resolver.
    /// Original entry bytes remain available through <see cref="Entries"/>. This projection does not recreate the Windows help viewer.</remarks>
    public ChmConversionResult<HtmlConversionDocument> ToHtmlDocumentResult(ChmConversionOptions? options = null, CancellationToken cancellationToken = default) {
        ChmConversionOptions configured = options?.Clone() ?? new ChmConversionOptions();
        IReadOnlyList<ChmTopic> topics = SelectTopics(configured);
        var diagnostics = SourceDiagnostics();
        HtmlDocument output = new HtmlDocumentEngine(_options.ParserProvider).ParseDocument("<!doctype html><html><head><meta charset=\"utf-8\"></head><body></body></html>").CloneAttached(cancellationToken);
        HtmlElement title = output.CreateElement("title"); title.TextContent = Title ?? "Compiled help"; output.Head!.AppendChild(title);
        var anchors = topics.Select((topic, index) => new { topic.Path, Anchor = "chm-topic-" + (index + 1).ToString(CultureInfo.InvariantCulture) })
            .ToDictionary(item => item.Path, item => item.Anchor, StringComparer.OrdinalIgnoreCase);
        long characters = 0, embeddedBytes = 0;
        int nodes = 8, sourceNodes = 0;
        bool styles = false;
        string? previousStyle = null;
        foreach (ChmTopic topic in topics) {
            cancellationToken.ThrowIfCancellationRequested();
            string html = topic.ReadHtml(cancellationToken);
            ReserveCharacters(ref characters, html.Length, configured);
            HtmlConversionDocument source = ParseConversionTopic(html, topic.Path, configured, ref sourceNodes, cancellationToken);
            if (source.Document.QuerySelectorAll("script,object,embed,iframe,frame,frameset").Count != 0 ||
                source.Document.QuerySelectorAll("*").Any(element => element.Attributes.Any(attribute =>
                    attribute.LocalName.StartsWith("on", StringComparison.OrdinalIgnoreCase) || attribute.LocalName.Equals("srcdoc", StringComparison.OrdinalIgnoreCase))))
                diagnostics.Add(new OfficeConversionFidelityDiagnostic("CHM_ACTIVE_CONTENT_OMITTED", "Scripts, executable attributes, embedded controls and help-viewer frames are not reproduced in the document projection.", OfficeConversionLossKind.Omission, "OfficeIMO.Chm", topic.Path));
            HtmlDocument normalized = source.CreateDocumentForConversion();
            string prefix = anchors[topic.Path];
            foreach (HtmlElement element in normalized.QuerySelectorAll("*")) {
                cancellationToken.ThrowIfCancellationRequested();
                if (++nodes > configured.MaxHtmlNodes) throw ChmBinary.Error("CONVERSION_LIMIT", "The combined book exceeds MaxHtmlNodes.");
                if (element.LocalName == "style" || (element.LocalName == "link" &&
                    element.GetAttribute("rel")?.IndexOf("stylesheet", StringComparison.OrdinalIgnoreCase) >= 0)) styles = true;
                foreach (string attribute in new[] { "id", "name" }) {
                    if (attribute == "name" && element.LocalName != "a" && element.LocalName != "map") continue;
                    string? attributeValue = element.GetAttribute(attribute);
                    if (!string.IsNullOrEmpty(attributeValue)) element.SetAttribute(attribute, prefix + "-" + attributeValue);
                }
                foreach (HtmlAttribute attribute in element.Attributes.Where(attribute => attribute.NamespaceUri.Length == 0 &&
                    HtmlIdentifierReferences.IsReference(element.NamespaceUri, element.LocalName, attribute.LocalName)).ToArray()) {
                    element.SetAttribute(attribute.Name, System.Text.RegularExpressions.Regex.Replace(attribute.Value, @"[^ \t\r\n\f]+",
                        match => prefix + "-" + match.Value));
                }
                string? imageMap = element.GetAttribute("usemap");
                if (imageMap != null && imageMap.StartsWith("#", StringComparison.Ordinal)) element.SetAttribute("usemap", "#" + prefix + "-" + imageMap.Substring(1));
                RewriteSvgReferences(element, prefix, topic.Path, anchors, configured, ref embeddedBytes, diagnostics, cancellationToken);
                string? href = element.GetAttribute("href");
                if (href != null && element.NamespaceUri != "http://www.w3.org/2000/svg" && (element.LocalName == "a" || element.LocalName == "area")) {
                    string? rewritten = RewriteBookLink(href, topic.Path, anchors);
                    if (rewritten != null) element.SetAttribute("href", rewritten);
                    else if (Uri.TryCreate(href, UriKind.Absolute, out Uri? missing) && missing.Scheme == "chm") {
                        element.RemoveAttribute("href");
                        diagnostics.Add(new OfficeConversionFidelityDiagnostic("CHM_LINK_OMITTED", "The link targets a missing or unselected help topic.", OfficeConversionLossKind.Omission, "OfficeIMO.Chm", topic.Path));
                    }
                }
                if (configured.EmbedImages && (element.LocalName == "img" || element.LocalName == "source"))
                    EmbedImages(element, configured, topic.Path, ref embeddedBytes, diagnostics, cancellationToken);
            }
            foreach (HtmlElement element in normalized.Head?.Children ?? Array.Empty<HtmlElement>()) {
                if (element.LocalName != "style" && !(element.LocalName == "link" && element.GetAttribute("rel")?.IndexOf("stylesheet", StringComparison.OrdinalIgnoreCase) >= 0)) continue;
                styles = true;
                string markup = element.OuterHtml;
                // Compilers repeat the same stylesheet on every topic. Adjacent identical
                // declarations have the same inert cascade; retain A/B/A ordering when distinct.
                if (markup == previousStyle) continue;
                output.Head.AppendChild(output.ImportNode(element, cancellationToken: cancellationToken));
                previousStyle = markup;
            }
            HtmlElement section = output.CreateElement("section"); section.SetAttribute("id", prefix);
            section.SetAttribute("data-chm-topic", topic.Path);
            PreserveTopicLanguage(output, section, normalized, topic == topics[0]);
            HtmlElement container = PreserveTopicContainer(output, section, normalized.DocumentElement, topic.Path, diagnostics, ref nodes, configured);
            container = PreserveTopicContainer(output, container, normalized.Body, topic.Path, diagnostics, ref nodes, configured);
            if (nodes > configured.MaxHtmlNodes - 2) throw ChmBinary.Error("CONVERSION_LIMIT", "The combined book exceeds MaxHtmlNodes.");
            nodes += 2;
            HtmlElement heading = output.CreateElement("h1"); heading.TextContent = topic.Title; container.AppendChild(heading);
            foreach (HtmlNode node in normalized.Body?.ChildNodes ?? Array.Empty<HtmlNode>()) {
                HtmlNode copy = output.ImportNode(node, cancellationToken: cancellationToken);
                container.AppendChild(copy);
            }
            // Topic headings own book/chapter boundaries. Demote authored headings one level.
            if (DemoteTopicHeadings(output, section, heading, cancellationToken)) diagnostics.Add(new OfficeConversionFidelityDiagnostic("CHM_HEADING_LEVELS",
                "Authored headings are demoted one level beneath the topic heading. Heading element selectors in retained CSS are not rewritten.",
                OfficeConversionLossKind.Approximation, "OfficeIMO.Chm", topic.Path));
            output.Body!.AppendChild(section);
            diagnostics.AddRange(source.Diagnostics.Select(item => new OfficeConversionFidelityDiagnostic(item.Code, item.Message, item.LossKind, "OfficeIMO.Html", topic.Path)));
        }
        if (styles && topics.Count > 1) diagnostics.Add(new OfficeConversionFidelityDiagnostic("CHM_COMBINED_STYLE_SCOPE", "Independent topic styles share one cascade in the reflowable book projection.", OfficeConversionLossKind.Approximation, "OfficeIMO.Chm"));
        if (styles) diagnostics.Add(new OfficeConversionFidelityDiagnostic("CHM_STYLE_IDENTIFIERS", "Topic anchor identifiers are prefixed for book-wide uniqueness; authored CSS identifier selectors are not rewritten.", OfficeConversionLossKind.Approximation, "OfficeIMO.Chm"));
        if (Index.Count != 0) diagnostics.Add(new OfficeConversionFidelityDiagnostic("CHM_INDEX_METADATA", "The compiled keyword index remains in ChmDocument.Index; it is not a help-viewer index in this projection.", OfficeConversionLossKind.Omission, "OfficeIMO.Chm"));
        HtmlConversionDocumentOptions conversion = CreateHtmlOptions("/book.html");
        conversion.UseBodyContentsOnly = false;
        conversion.Limits.MaxInputCharacters = checked((int)Math.Min(configured.MaxOutputBytes, int.MaxValue));
        conversion.Limits.MaxHtmlNodes = configured.MaxHtmlNodes;
        // The book section and retained html/body wrappers can add three levels.
        conversion.Limits.MaxHtmlDepth = checked(_options.MaxHtmlDepth + 3);
        cancellationToken.ThrowIfCancellationRequested();
        HtmlConversionDocument value = HtmlConversionDocument.FromDocument(output, conversion, cancellationToken);
        EnforceOutput(Encoding.UTF8.GetByteCount(value.SourceHtml), configured);
        cancellationToken.ThrowIfCancellationRequested();
        return new ChmConversionResult<HtmlConversionDocument>(value, new ChmConversionReport(topics.Select(item => item.Path), diagnostics));
    }

    private static bool DemoteTopicHeadings(HtmlDocument output, HtmlElement section, HtmlElement topicHeading, CancellationToken token) {
        var parents = section.QuerySelectorAll("h1,h2,h3,h4,h5")
            .Where(element => element.NamespaceUri == "http://www.w3.org/1999/xhtml" && !ReferenceEquals(element, topicHeading))
            .GroupBy(element => element.Parent!).ToArray();
        // Descendant parents are processed before ancestors so nested content is retained.
        foreach (var group in parents.Reverse()) {
            token.ThrowIfCancellationRequested();
            var replacements = new Dictionary<HtmlNode, HtmlElement>();
            foreach (HtmlElement authored in group) {
                token.ThrowIfCancellationRequested();
                HtmlElement replacement = output.CreateElement("h" + (authored.LocalName[1] - '0' + 1).ToString(CultureInfo.InvariantCulture));
                foreach (HtmlAttribute attribute in authored.Attributes) replacement.SetAttribute(attribute);
                HtmlNode[] children = authored.ChildNodes.ToArray();
                authored.TextContent = string.Empty;
                foreach (HtmlNode child in children) {
                    token.ThrowIfCancellationRequested();
                    replacement.AppendChild(child);
                }
                replacements.Add(authored, replacement);
            }
            HtmlNode parent = group.Key;
            HtmlNode[] siblings = parent.ChildNodes.ToArray();
            // Clearing detaches in one pass; per-heading Remove would repeatedly shift the list.
            parent.TextContent = string.Empty;
            foreach (HtmlNode sibling in siblings) {
                token.ThrowIfCancellationRequested();
                parent.AppendChild(replacements.TryGetValue(sibling, out HtmlElement? replacement) ? replacement : sibling);
            }
        }
        return parents.Length != 0;
    }

    private string? RewriteBookLink(string href, string sourcePath, Dictionary<string, string> anchors) {
        string reference = href;
        ChmEntry? target;
        if (Uri.TryCreate(href, UriKind.Absolute, out Uri? uri)) {
            if (uri.Scheme != "chm" || uri.Host != "archive") return href;
            target = FindUriEntry(uri);
            reference = uri.Fragment;
        } else target = FindEntry(reference, sourcePath);
        if (target == null || !anchors.TryGetValue(target.Path, out string? anchor)) return null;
        int hash = reference.IndexOf('#');
        if (hash >= 0 && hash + 1 < reference.Length) anchor += "-" + Uri.UnescapeDataString(reference.Substring(hash + 1));
        return "#" + Uri.EscapeDataString(anchor);
    }

}
