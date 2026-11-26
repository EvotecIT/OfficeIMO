using System.Globalization;
using System.Text;
using System.Xml.Linq;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using OfficeIMO.Epub;

namespace OfficeIMO.Mobi;

public static partial class MobiReader {
    private sealed partial class Projection {
        private void AddMarkup(EpubPublication book, byte[] text) {
            Kf8Content? kindle = selected.IsKf8 ? new Kf8Content(database, selected, text, token) : null;
            if (kindle?.HasSeparateFlows == true)
                Warn("MOBI_FLOWS_OMITTED", "Separate Kindle flows such as CSS, SVG and fonts are outside the static content projection.");
            byte[][] inputs = kindle?.Parts ?? new[] { text };
            var parts = new MarkupPart[inputs.Length];
            var originalMarkup = new string[inputs.Length];
            var textRuns = new List<MarkupTextRun>[inputs.Length];
            var currentRuns = new List<MarkupTextRun>();
            var pageBreaks = new List<int>();
            var positions = new HashSet<int>[inputs.Length];
            int tokens = 0, elements = 0;
            bool hasMarkup = false;
            int remainingXhtmlTokens = 300_000;
            int generatedEndTags = 0;
            var parser = new HtmlParser(new HtmlParserOptions {
                IsKeepingSourceReferences = true,
                SkipComments = true,
                OnToken = (markupToken, range) => {
                    Charge(ref tokens, 300_000 + generatedEndTags);
                    if (markupToken.Type == HtmlTokenType.StartTag) hasMarkup = true;
                    if (markupToken.Type == HtmlTokenType.Character)
                        currentRuns.Add(new MarkupTextRun(range.Start.Index, range.End.Index, markupToken.Data));
                    if (markupToken.Type == HtmlTokenType.StartTag && markupToken.Name == "mbp:pagebreak")
                        pageBreaks.Add(range.Start.Index);
                },
                OnCreated = (_, _) => Charge(ref elements, 150_000)
            });

            void Charge(ref int count, int maximum) {
                token.ThrowIfCancellationRequested();
                if (++count > maximum) throw new InvalidDataException("Mobipocket HTML exceeds its aggregate token or element limit.");
            }
            for (int i = 0; i < inputs.Length; i++) {
                token.ThrowIfCancellationRequested();
                string html = selected.Encoding.GetString(inputs[i]);
                List<(int Offset, int Length)>? insertedClosers = null;
                if (kindle != null) html = ExpandKindleEmptyElements(html, ref remainingXhtmlTokens, out insertedClosers, token);
                // The bounded XHTML prepass supplies at most one parser token per generated close.
                generatedEndTags += insertedClosers?.Count ?? 0;
                currentRuns = new List<MarkupTextRun>();
                pageBreaks.Clear();
                int tokensBeforePart = tokens, elementsBeforePart = elements;
                var document = parser.ParseDocument(html);
                if (selected.IsHeaderless && !hasMarkup) {
                    AddLiteralText(book, html);
                    return;
                }
                if (pageBreaks.Count != 0) {
                    // A length-preserving token rewrite keeps filepos and Kindle offsets in the original coordinate space.
                    html = NormalizePageBreaks(html, pageBreaks);
                    tokens = tokensBeforePart;
                    elements = elementsBeforePart;
                    currentRuns.Clear();
                    pageBreaks.Clear();
                    document = parser.ParseDocument(html);
                    Warn("MOBI_MARKUP_APPROXIMATED", "A Mobipocket page break was retained as a separator.", "mbp:pagebreak",
                        OfficeConversionLossKind.Approximation);
                }
                originalMarkup[i] = html;
                textRuns[i] = currentRuns;
                if (document.Head is { } head && (head.QuerySelector("link[rel=stylesheet]") != null ||
                    head.QuerySelectorAll("style").Any(style => !string.IsNullOrWhiteSpace(style.TextContent))))
                    Warn("MOBI_STYLE_OMITTED", "Source stylesheets are omitted. Semantic formatting is retained.");
                parts[i] = new MarkupPart(i, document.Body ?? throw new InvalidDataException("Mobipocket markup has no document body."),
                    XmlText(document.Title ?? string.Empty), inputs[i].Length, insertedClosers, token);
                positions[i] = new HashSet<int>();
            }

            var navigation = ReadNavigation(kindle);
            foreach (Navigation item in navigation) {
                if (item.Target is { } target) AddPosition(target);
            }
            foreach (MarkupPart part in parts) {
                foreach (IElement element in part.Elements) {
                    token.ThrowIfCancellationRequested();
                    if (element.LocalName == "a" && LinkPosition(element.GetAttribute("href"), element.GetAttribute("filepos"), kindle) is { } target)
                        AddPosition(target);
                }
            }
            for (int i = 0; i < parts.Length; i++) {
                MarkupPart part = parts[i];
                if (part.BindPositions(positions[i], inputs[i], selected.Encoding, originalMarkup[i], textRuns[i], token))
                    Warn("MOBI_LINK_POSITION_APPROXIMATED", "A Mobipocket text position was moved to the nearest readable anchor.",
                        part.File, OfficeConversionLossKind.Approximation);
            }

            var projected = new XElement[parts.Length];
            for (int i = 0; i < parts.Length; i++) {
                MarkupPart part = parts[i];
                projected[i] = new XElement("div", part.Body.ChildNodes.Select(node => Render(node, i, parts, kindle)));
                if (part.Anchor(part.Body) is { } id) projected[i].SetAttributeValue("id", id);
                projected[i].SetAttributeValue("dir", TextDirection(part.Body) ?? TextDirection(part.Body.ParentElement));
                SetMarkupLanguage(projected[i], MarkupLanguage(part.Body) ?? MarkupLanguage(part.Body.ParentElement));
            }

            int? cover = selected.Integer(201) ?? first.Integer(201);
            if (cover is >= 0 and < int.MaxValue) {
                int imageIndex = cover.Value + 1;
                if (_images.TryGetValue(imageIndex, out EpubManifestItem? image)) {
                    book.SetCoverImage(image.Id);
                    if (!_usedImages.Contains(imageIndex)) {
                        var coverBody = new XElement("div", new XElement("img",
                            new XAttribute("alt", book.Title + " cover"), new XAttribute("src", image.Href)));
                        book.AddChapter("cover", "EPUB/cover.xhtml", "Cover", coverBody.ToString(SaveOptions.DisableFormatting));
                    }
                } else Warn("MOBI_COVER_UNAVAILABLE", "The declared cover was not a supported embedded raster image.");
            }

            for (int i = 0; i < parts.Length; i++) {
                token.ThrowIfCancellationRequested();
                MarkupPart part = parts[i];
                book.AddChapter("part-" + Number(i + 1), "EPUB/" + part.File,
                    string.IsNullOrWhiteSpace(part.Title) ? book.Title : part.Title, projected[i].ToString(SaveOptions.DisableFormatting));
            }
            var contents = BuildNavigation(navigation, parts);
            if (contents.Count > 0) book.SetNavigation(contents);

            void AddPosition((int Part, int Offset) position) {
                if (position.Part >= 0 && position.Part < parts.Length && position.Offset >= 0 &&
                    position.Offset <= parts[position.Part].Length)
                    positions[position.Part].Add(position.Offset);
            }
        }

        private static readonly string[] LegacyImageAttributes = { "recindex", "hirecindex", "lowrecindex" };

        private int LegacyImageIndex(IElement source) {
            foreach (string attribute in LegacyImageAttributes) {
                if (int.TryParse(source.GetAttribute(attribute), NumberStyles.None, CultureInfo.InvariantCulture, out int index) &&
                    index > 0 && _images.ContainsKey(index)) return index;
            }
            return 0;
        }

        private static string? TextDirection(IElement? element) {
            string? direction = element?.GetAttribute("dir")?.Trim().ToLowerInvariant();
            return direction is "ltr" or "rtl" or "auto" ? direction : null;
        }

        private static string? MarkupLanguage(IElement? element) =>
            (element?.GetAttribute("lang") ?? element?.GetAttribute("xml:lang")) is { } language ? XmlText(language) : null;

        private static void SetMarkupLanguage(XElement target, string? language) {
            target.SetAttributeValue("lang", language);
            target.SetAttributeValue(XNamespace.Xml + "lang", language);
        }

        private object? Render(INode node, int partIndex, MarkupPart[] parts, Kf8Content? kindle) {
            token.ThrowIfCancellationRequested();
            if (node is IText text) return parts[partIndex].RenderText(text);
            if (node is not IElement source) return null;
            string tag = source.LocalName;
            MarkupPart part = parts[partIndex];
            if (IsOmittedMarkup(tag)) {
                if (tag != "style" || !string.IsNullOrWhiteSpace(source.TextContent))
                    Warn("MOBI_ELEMENT_OMITTED", "An unsupported active, media or styling element was omitted.", tag);
                // An omitted subtree can still carry a link target in the reading order.
                var ids = source.QuerySelectorAll("*").Prepend(source)
                    .Select(part.Anchor).OfType<string>().ToArray();
                return ids.Length == 0 ? null : new XElement("span",
                    ids.Select(id => new XElement("span", new XAttribute("id", id))));
            }

            string targetTag = tag switch {
                "p" or "div" or "span" or "section" or "article" or "aside" or "header" or "footer" or "nav" or
                "h1" or "h2" or "h3" or "h4" or "h5" or "h6" or "a" or "br" or "hr" or "wbr" or "blockquote" or "ul" or "ol" or
                "li" or "em" or "strong" or "b" or "i" or "u" or "s" or "sub" or "sup" or "code" or "pre" or "dfn" or "cite" or
                "table" or "thead" or "tbody" or "tfoot" or "tr" or "td" or "th" or "caption" or "img" or "figure" or "figcaption" or
                "dl" or "dt" or "dd" or "q" => tag,
                "center" => "div",
                _ => "span"
            };
            if (tag != targetTag || source.HasAttribute("style"))
                Warn("MOBI_MARKUP_APPROXIMATED", "Legacy markup or CSS presentation was approximated.", tag, OfficeConversionLossKind.Approximation);

            var target = new XElement(targetTag);
            if (part.Anchor(source) is { } id) target.SetAttributeValue("id", id);
            target.SetAttributeValue("dir", TextDirection(source));
            SetMarkupLanguage(target, MarkupLanguage(source));
            foreach (string name in new[] { "title", "alt" }) {
                if (source.GetAttribute(name) is { } value) target.SetAttributeValue(name, XmlText(value));
            }

            if (tag is "ol" or "li") {
                string numberAttribute = tag == "ol" ? "start" : "value";
                if (source.GetAttribute(numberAttribute) is { } rawNumber) {
                    if (int.TryParse(rawNumber, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value))
                        target.SetAttributeValue(numberAttribute, Number(value));
                    else
                        Warn("MOBI_LIST_NUMBERING_OMITTED", "An ordered-list number was outside the supported integer range.", numberAttribute);
                }
                if (tag == "ol") {
                    if (source.HasAttribute("reversed")) target.SetAttributeValue("reversed", "reversed");
                    string? markerType = source.GetAttribute("type");
                    if (markerType is "1" or "A" or "a" or "I" or "i")
                        target.SetAttributeValue("type", markerType);
                }
            }

            if (tag == "a") {
                string? href = source.GetAttribute("href");
                string? filepos = source.GetAttribute("filepos");
                string? resolved = null;
                if (LinkPosition(href, filepos, kindle) is { } position)
                    TryResolvePosition(position, parts, out resolved);
                if (resolved != null) target.SetAttributeValue("href", resolved);
                else if (href == "#") target.SetAttributeValue("href", "#" + part.TopAnchor);
                else if (href is { Length: > 1 } && href[0] == '#' && part.ResolveId(href.Substring(1)) is { } local)
                    target.SetAttributeValue("href", "#" + local);
                else if (href != null && EpubImportedPublication.IsWebLink(href)) target.SetAttributeValue("href", XmlText(href));
                else if (href != null || filepos != null)
                    Warn("MOBI_LINK_OMITTED", "A Mobipocket link target could not be resolved safely.", href ?? filepos);
            }
            if (tag == "img") {
                string? reference = source.GetAttribute("src");
                int index = 0;
                const string prefix = "kindle:embed:";
                if (reference != null && reference.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) {
                    int end = reference.IndexOf('?');
                    if (end < 0) end = reference.Length;
                    if (end < prefix.Length || !TryBase32(reference.AsSpan(prefix.Length, end - prefix.Length), out index)) index = 0;
                } else {
                    index = LegacyImageIndex(source);
                }
                if (index <= 0 || !_images.TryGetValue(index, out EpubManifestItem? image)) {
                    Warn("MOBI_IMAGE_UNAVAILABLE", "A Mobipocket image does not resolve to a supported embedded raster.", reference ?? source.GetAttribute("recindex") ?? source.GetAttribute("hirecindex") ?? source.GetAttribute("lowrecindex"));
                    return new XElement("span", target.Attribute("id"), new XText((string?)target.Attribute("alt") ?? ""));
                }
                _usedImages.Add(index);
                target.SetAttributeValue("src", image.Href);
                target.SetAttributeValue("alt", (string?)target.Attribute("alt") ?? "");
            }
            if (tag is "img" or "td" or "th") {
                foreach (string name in tag == "img" ? new[] { "width", "height" } : new[] { "rowspan", "colspan" }) {
                    if (int.TryParse(source.GetAttribute(name), NumberStyles.None, CultureInfo.InvariantCulture, out int value) &&
                        (value > 0 || (name == "rowspan" && value == 0)) &&
                        value <= (tag == "img" ? 10_000 : 100))
                        target.SetAttributeValue(name, Number(value));
                }
            }
            if (tag is not ("img" or "br" or "hr" or "wbr"))
                target.Add(source.ChildNodes.Select(child => Render(child, partIndex, parts, kindle)));
            return target;
        }
    }

    private static bool IsOmittedMarkup(string tag) =>
        tag is "script" or "style" or "object" or "iframe" or "svg" or "audio" or "video" or "form" or "template";

    private static string NormalizePageBreaks(string markup, List<int> pageBreaks) {
        const string tag = "mbp:pagebreak";
        char[] characters = markup.ToCharArray();
        foreach (int tagStart in pageBreaks) {
            characters[tagStart + 1] = 'h';
            characters[tagStart + 2] = 'r';
            // Whitespace preserves token positions and retains any attributes on the source tag.
            for (int index = tagStart + 3; index <= tagStart + tag.Length; index++) characters[index] = ' ';
        }
        return new string(characters);
    }

    private sealed partial class MarkupPart {
        private readonly Dictionary<IElement, string> _anchors = new();
        private readonly Dictionary<string, string> _ids = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _names = new(StringComparer.Ordinal);
        private readonly Dictionary<int, string> _positions = new();
        private readonly List<IElement> _elements = new();
        private readonly List<(int Offset, int Length)>? _insertedClosers;

        internal MarkupPart(int index, IElement body, string title, int length,
            List<(int Offset, int Length)>? insertedClosers, CancellationToken token) {
            Body = body;
            Length = length;
            Title = title;
            File = "part-" + Number(index + 1) + ".xhtml";
            _insertedClosers = insertedClosers;
            if (body.ParentElement?.GetAttribute("id") is { Length: > 0 } rootId)
                _ids.Add(rootId, TopAnchor);
            var pending = new Stack<(IElement Element, int Depth)>();
            pending.Push((body, 1));
            while (pending.Count > 0) {
                token.ThrowIfCancellationRequested();
                var (element, depth) = pending.Pop();
                if (depth > 64 || _elements.Count >= 150_000)
                    throw new InvalidDataException("Mobipocket HTML exceeds its nesting or node count limit.");
                _elements.Add(element);
                string? id = element.GetAttribute("id");
                if (!string.IsNullOrEmpty(id) && !_ids.ContainsKey(id!)) _ids.Add(id!, EnsureAnchor(element));
                string? name = element.LocalName == "a" ? element.GetAttribute("name") : null;
                if (!string.IsNullOrEmpty(name) && !_names.ContainsKey(name!)) _names.Add(name!, EnsureAnchor(element));
                for (IElement? child = element.LastElementChild; child != null; child = child.PreviousElementSibling)
                    pending.Push((child, depth + 1));
            }
        }

        internal string TopAnchor => EnsureAnchor(Body);
        internal string File { get; }
        internal IElement Body { get; }
        internal int Length { get; }
        internal string Title { get; }
        internal IReadOnlyList<IElement> Elements => _elements;
        internal string? AnchorAt(int offset) => _positions.TryGetValue(offset, out string? id) ? id : null;
        internal string? Anchor(IElement element) => _anchors.TryGetValue(element, out string? id) ? id : null;
        internal string? ResolveId(string id) {
            if (FindLocalAnchor(id) is { } anchor) return anchor;
            string decoded = id.IndexOf('%') >= 0 ? Uri.UnescapeDataString(id) : id;
            if (decoded != id && FindLocalAnchor(decoded) is { } decodedAnchor) return decodedAnchor;
            return decoded.Equals("top", StringComparison.OrdinalIgnoreCase) ? TopAnchor : null;
        }

        private string? FindLocalAnchor(string fragment) =>
            _ids.TryGetValue(fragment, out string? anchor) || _names.TryGetValue(fragment, out anchor) ? anchor : null;

        private string EnsureAnchor(IElement element) {
            if (!_anchors.TryGetValue(element, out string? id)) {
                id = "mobi-" + Number(_anchors.Count + 1);
                _anchors.Add(element, id);
            }
            return id;
        }
    }
}
