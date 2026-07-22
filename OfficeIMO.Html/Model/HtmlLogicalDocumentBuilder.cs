using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// Builds the shared OfficeIMO HTML logical document model from parsed or raw HTML.
/// </summary>
internal static class HtmlLogicalDocumentBuilder {
    private static readonly char[] WhitespaceSeparators = { ' ', '\t', '\r', '\n', '\f' };

    /// <summary>
    /// Parses raw HTML and builds a logical document from the conversion root.
    /// </summary>
    public static HtmlLogicalDocument FromHtml(string html, bool useBodyContentsOnly = true) {
        IHtmlDocument document = HtmlConversionDocument.ParseSourceDocumentForAnalysis(html);
        return FromDocument(document, useBodyContentsOnly);
    }

    /// <summary>
    /// Builds a logical document from an AngleSharp HTML document.
    /// </summary>
    public static HtmlLogicalDocument FromDocument(IHtmlDocument document, bool useBodyContentsOnly = true) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        INode rootNode = HtmlDocumentParser.GetConversionRoot(document, useBodyContentsOnly);
        var counts = new Dictionary<HtmlLogicalNodeKind, int>();
        var capabilities = new List<string>();
        HtmlLogicalNode root = Build(rootNode, counts, capabilities, forceRetain: true)!;
        return new HtmlLogicalDocument(root, counts, capabilities);
    }

    private static HtmlLogicalNode? Build(INode source, IDictionary<HtmlLogicalNodeKind, int> counts, ICollection<string> capabilities, bool forceRetain = false) {
        if (!forceRetain && source is IElement sourceElement && IsNonDocumentLogicalElement(sourceElement.TagName)) {
            return null;
        }

        HtmlLogicalNode node = CreateNode(source);

        if (source is IElement element) {
            string accessibleName = HtmlAccessibilitySemantics.GetAccessibleName(
                element,
                includeTextFallback: node.Kind == HtmlLogicalNodeKind.Link);
            node.AccessibleName = accessibleName.Length == 0 ? null : accessibleName;
            foreach (IAttr attribute in element.Attributes) {
                node.AddAttribute(attribute.Name, attribute.Value);
            }
        }

        bool suppressCapturedTextChildren = source is IElement && node.Kind == HtmlLogicalNodeKind.Text && node.Text.Length > 0;
        foreach (INode child in source.ChildNodes) {
            if (child.NodeType == NodeType.Comment) {
                continue;
            }

            if (suppressCapturedTextChildren && child.NodeType == NodeType.Text) {
                continue;
            }

            HtmlLogicalNode? childNode = Build(child, counts, capabilities);
            if (childNode != null) {
                node.AddChild(childNode);
            }
        }

        if (!forceRetain && !ShouldRetain(node)) {
            return null;
        }

        Count(counts, node.Kind);
        foreach (string capability in InferCapabilities(node)) {
            node.AddCapability(capability);
            if (!capabilities.Contains(capability)) {
                capabilities.Add(capability);
            }
        }

        return node;
    }

    private static bool ShouldRetain(HtmlLogicalNode node) {
        return node.Kind != HtmlLogicalNodeKind.Unknown || node.Children.Count > 0 || node.Text.Length > 0;
    }

    private static bool IsNonDocumentLogicalElement(string name) {
        return string.Equals(name, "script", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "style", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "template", StringComparison.OrdinalIgnoreCase);
    }

    private static HtmlLogicalNode CreateNode(INode source) {
        if (source.NodeType == NodeType.Text) {
            string text = NormalizeText(source.TextContent);
            return new HtmlLogicalNode(text.Length == 0 ? HtmlLogicalNodeKind.Unknown : HtmlLogicalNodeKind.Text, "#text", text);
        }

        if (source is IHtmlDocument) {
            return new HtmlLogicalNode(HtmlLogicalNodeKind.Document, "#document", string.Empty);
        }

        if (!(source is IElement element)) {
            return new HtmlLogicalNode(HtmlLogicalNodeKind.Unknown, source.NodeName ?? string.Empty, NormalizeText(source.TextContent));
        }

        string name = element.TagName.ToLowerInvariant();
        HtmlLogicalNodeKind kind = MapKind(name, element);
        string capturedText = kind == HtmlLogicalNodeKind.Text
            ? NormalizeText(element.TextContent)
            : CaptureText(name, element, kind);
        return new HtmlLogicalNode(kind, name, capturedText);
    }

    private static HtmlLogicalNodeKind MapKind(string name, IElement element) {
        if (HtmlAccessibilitySemantics.TryGetHeadingLevel(element, out _)) {
            return HtmlLogicalNodeKind.Heading;
        }
        if (IsFootnoteDefinition(element)) return HtmlLogicalNodeKind.Footnote;
        if (HtmlAccessibilitySemantics.HasRole(element, "list")) return HtmlLogicalNodeKind.List;
        if (HtmlAccessibilitySemantics.HasRole(element, "listitem")) return HtmlLogicalNodeKind.ListItem;
        if (HtmlAccessibilitySemantics.HasRole(element, "table")) return HtmlLogicalNodeKind.Table;
        if (HtmlAccessibilitySemantics.HasRole(element, "row")) return HtmlLogicalNodeKind.TableRow;
        if (HtmlAccessibilitySemantics.HasRole(element, "cell")
            || HtmlAccessibilitySemantics.HasRole(element, "columnheader")
            || HtmlAccessibilitySemantics.HasRole(element, "rowheader")) {
            return HtmlLogicalNodeKind.TableCell;
        }
        if (HtmlAccessibilitySemantics.HasRole(element, "img")) return HtmlLogicalNodeKind.Image;
        if (HtmlAccessibilitySemantics.HasRole(element, "link")) return HtmlLogicalNodeKind.Link;

        if (name == "body" || name == "main" || name == "article" || name == "section" || name == "aside" || name == "header" || name == "footer") {
            return HtmlLogicalNodeKind.Section;
        }

        if (name.Length == 2 && name[0] == 'h' && name[1] >= '1' && name[1] <= '6') {
            return HtmlLogicalNodeKind.Heading;
        }

        switch (name) {
            case "p":
                return HtmlLogicalNodeKind.Paragraph;
            case "pre":
                return HtmlLogicalNodeKind.Code;
            case "blockquote":
                return HtmlLogicalNodeKind.Quote;
            case "ul":
            case "ol":
            case "dl":
                return HtmlLogicalNodeKind.List;
            case "li":
            case "dt":
            case "dd":
                return HtmlLogicalNodeKind.ListItem;
            case "table":
                return HtmlLogicalNodeKind.Table;
            case "tr":
                return HtmlLogicalNodeKind.TableRow;
            case "td":
            case "th":
                return HtmlLogicalNodeKind.TableCell;
            case "caption":
                return HtmlLogicalNodeKind.TableCaption;
            case "figure":
                return HtmlLogicalNodeKind.Figure;
            case "figcaption":
                return HtmlLogicalNodeKind.Text;
            case "img":
            case "image":
            case "svg":
                return HtmlLogicalNodeKind.Image;
            case "picture":
                return HtmlLogicalNodeKind.Inline;
            case "source":
                return string.Equals(element.ParentElement?.TagName, "picture", StringComparison.OrdinalIgnoreCase)
                    ? HtmlLogicalNodeKind.Image
                    : HtmlLogicalNodeKind.Media;
            case "video":
            case "audio":
            case "track":
                return HtmlLogicalNodeKind.Media;
            case "a":
            case "area":
                return HtmlLogicalNodeKind.Link;
            case "form":
                return HtmlLogicalNodeKind.Form;
            case "input":
            case "select":
            case "textarea":
            case "button":
            case "option":
                return HtmlLogicalNodeKind.FormControl;
            case "title":
            case "meta":
            case "base":
            case "link":
            case "style":
                return HtmlLogicalNodeKind.Metadata;
            case "span":
            case "strong":
            case "em":
            case "b":
            case "i":
            case "u":
            case "small":
            case "sub":
            case "sup":
            case "code":
                return HtmlLogicalNodeKind.Inline;
            default:
                return element.Children.Length == 0 && NormalizeText(element.TextContent).Length > 0
                    ? HtmlLogicalNodeKind.Text
                    : HtmlLogicalNodeKind.Unknown;
        }
    }

    private static string CaptureText(string name, IElement element, HtmlLogicalNodeKind kind) {
        if (kind == HtmlLogicalNodeKind.Heading) return NormalizeText(element.TextContent);
        if (kind == HtmlLogicalNodeKind.Code) return NormalizePreformattedText(element.TextContent);
        if (kind == HtmlLogicalNodeKind.Quote) return NormalizeText(element.TextContent);
        if (kind == HtmlLogicalNodeKind.Footnote) return CaptureFootnoteText(element);
        switch (name) {
            case "h1":
            case "h2":
            case "h3":
            case "h4":
            case "h5":
            case "h6":
            case "title":
            case "p":
            case "figcaption":
            case "caption":
            case "label":
                return NormalizeText(element.TextContent);
            default:
                return string.Empty;
        }
    }

    private static IEnumerable<string> InferCapabilities(HtmlLogicalNode node) {
        if (!string.IsNullOrWhiteSpace(node.AccessibleName)
            || node.Attributes.ContainsKey("role")
            || node.Attributes.Keys.Any(static key => key.StartsWith("aria-", StringComparison.OrdinalIgnoreCase))) {
            yield return "accessibility";
        }
        if (HasFootnoteSemantic(node)) yield return "footnotes";

        switch (node.Kind) {
            case HtmlLogicalNodeKind.Heading:
                yield return "headings";
                break;
            case HtmlLogicalNodeKind.Code:
                yield return "code";
                break;
            case HtmlLogicalNodeKind.Quote:
                yield return "quotes";
                break;
            case HtmlLogicalNodeKind.Footnote:
                yield return "footnotes";
                break;
            case HtmlLogicalNodeKind.Table:
            case HtmlLogicalNodeKind.TableCell:
            case HtmlLogicalNodeKind.TableCaption:
                yield return "tables";
                break;
            case HtmlLogicalNodeKind.Image:
                yield return "images";
                break;
            case HtmlLogicalNodeKind.Media:
                yield return "media";
                break;
            case HtmlLogicalNodeKind.Form:
            case HtmlLogicalNodeKind.FormControl:
                yield return "forms";
                break;
            case HtmlLogicalNodeKind.Link:
                yield return "links";
                break;
            case HtmlLogicalNodeKind.List:
            case HtmlLogicalNodeKind.ListItem:
                yield return "lists";
                break;
            case HtmlLogicalNodeKind.Figure:
                yield return "figures";
                break;
        }
    }

    private static bool HasFootnoteSemantic(HtmlLogicalNode node) {
        if (node.Attributes.TryGetValue("role", out string? role)
            && (HtmlAccessibilitySemantics.ContainsToken(role, "doc-noteref")
                || HtmlAccessibilitySemantics.ContainsToken(role, "doc-footnote")
                || HtmlAccessibilitySemantics.ContainsToken(role, "doc-endnote")
                || HtmlAccessibilitySemantics.ContainsToken(role, "doc-endnotes")
                || HtmlAccessibilitySemantics.ContainsToken(role, "doc-backlink"))) {
            return true;
        }
        return node.Attributes.TryGetValue("epub:type", out string? epubType)
               && new[] { "noteref", "footnote", "footnotes", "endnote", "endnotes", "rearnote", "rearnotes", "backlink" }
                   .Any(token => HtmlAccessibilitySemantics.ContainsToken(epubType, token));
    }

    private static bool IsFootnoteDefinition(IElement element) =>
        HtmlAccessibilitySemantics.HasRole(element, "doc-footnote")
        || HtmlAccessibilitySemantics.HasRole(element, "doc-endnote")
        || HtmlAccessibilitySemantics.HasEpubType(element, "footnote")
        || HtmlAccessibilitySemantics.HasEpubType(element, "endnote")
        || HtmlAccessibilitySemantics.HasEpubType(element, "rearnote")
        || IsFootnoteCollectionMember(element);

    private static bool IsFootnoteCollectionMember(IElement element) {
        if (!element.TagName.Equals("LI", StringComparison.OrdinalIgnoreCase)) return false;

        string? id = element.Id;
        if (string.IsNullOrWhiteSpace(id)) return false;

        string footnoteId = id!;
        if (!footnoteId.StartsWith("fn:", StringComparison.OrdinalIgnoreCase)
            && !footnoteId.StartsWith("fn-", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        for (IElement? current = element.ParentElement; current != null; current = current.ParentElement) {
            if (IsFootnoteCollection(current)) return true;
        }
        return false;
    }

    private static bool IsFootnoteCollection(IElement element) =>
        element.HasAttribute("data-footnotes")
        || element.ClassList.Contains("footnotes")
        || element.ClassList.Contains("endnotes")
        || HtmlAccessibilitySemantics.HasEpubType(element, "footnotes")
        || HtmlAccessibilitySemantics.HasEpubType(element, "endnotes")
        || HtmlAccessibilitySemantics.HasEpubType(element, "rearnotes")
        || HtmlAccessibilitySemantics.HasRole(element, "doc-endnotes");

    private static string CaptureFootnoteText(IElement element) {
        var builder = new StringBuilder();
        AppendFootnoteText(element, builder);
        return NormalizeText(builder.ToString());
    }

    private static void AppendFootnoteText(INode node, StringBuilder builder) {
        foreach (INode child in node.ChildNodes) {
            if (child is IElement element) {
                if (HtmlAccessibilitySemantics.HasRole(element, "doc-backlink")
                    || HtmlAccessibilitySemantics.HasEpubType(element, "backlink")
                    || element.HasAttribute("data-footnote-backref")
                    || element.ClassList.Contains("footnote-backref")) {
                    continue;
                }
                AppendFootnoteText(element, builder);
            } else if (child is IText text) {
                builder.Append(text.Data);
            }
        }
    }

    private static string NormalizePreformattedText(string? text) {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        return text!.Replace("\r\n", "\n").Replace('\r', '\n').Trim('\n');
    }

    private static void Count(IDictionary<HtmlLogicalNodeKind, int> counts, HtmlLogicalNodeKind kind) {
        if (!counts.ContainsKey(kind)) {
            counts[kind] = 0;
        }

        counts[kind]++;
    }

    private static string NormalizeText(string text) {
        if (string.IsNullOrWhiteSpace(text)) {
            return string.Empty;
        }

        return string.Join(" ", text.Split(WhitespaceSeparators, StringSplitOptions.RemoveEmptyEntries));
    }
}
