using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// Builds the shared OfficeIMO HTML logical document model from parsed or raw HTML.
/// </summary>
public static class HtmlLogicalDocumentBuilder {
    private static readonly char[] WhitespaceSeparators = { ' ', '\t', '\r', '\n', '\f' };

    /// <summary>
    /// Parses raw HTML and builds a logical document from the conversion root.
    /// </summary>
    public static HtmlLogicalDocument FromHtml(string html, bool useBodyContentsOnly = true) {
        IHtmlDocument document = HtmlDocumentParser.ParseDocument(html);
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
            || string.Equals(name, "template", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "noscript", StringComparison.OrdinalIgnoreCase);
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
        string capturedText = kind == HtmlLogicalNodeKind.Text ? NormalizeText(element.TextContent) : CaptureText(name, element);
        return new HtmlLogicalNode(kind, name, capturedText);
    }

    private static HtmlLogicalNodeKind MapKind(string name, IElement element) {
        if (name == "body" || name == "main" || name == "article" || name == "section" || name == "aside" || name == "header" || name == "footer") {
            return HtmlLogicalNodeKind.Section;
        }

        if (name.Length == 2 && name[0] == 'h' && name[1] >= '1' && name[1] <= '6') {
            return HtmlLogicalNodeKind.Heading;
        }

        switch (name) {
            case "p":
            case "blockquote":
            case "pre":
                return HtmlLogicalNodeKind.Paragraph;
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
            case "figcaption":
                return HtmlLogicalNodeKind.Figure;
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

    private static string CaptureText(string name, IElement element) {
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
        switch (node.Kind) {
            case HtmlLogicalNodeKind.Heading:
                yield return "headings";
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
