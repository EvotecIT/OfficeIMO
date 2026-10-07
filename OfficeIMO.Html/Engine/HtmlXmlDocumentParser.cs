using System.Xml.Linq;
using AngleSharp.Html.Dom;
using OfficeIMO.Html.Providers;

namespace OfficeIMO.Html;

/// <summary>Materializes XML content structurally; HTML tokenization would change empty elements and raw text.</summary>
internal static class HtmlXmlDocumentParser {
    internal static IHtmlDocument CreateDocument(XDocument xml, HtmlConversionLimits limits, CancellationToken token, bool skipProcessingInstructions = false) {
        limits.Validate();
        token.ThrowIfCancellationRequested();
        var owned = new Dom.HtmlDocument(AngleSharpDomServices.Instance, AngleSharpHtmlParser.Instance.Id);
        var tracker = HtmlDomLimitTracker.Create(limits.MaxHtmlNodes, limits.MaxHtmlDepth);
        var pending = new Stack<(IEnumerator<XNode> Nodes, Dom.HtmlNode Parent, int Depth)>();
        pending.Push((xml.Nodes().GetEnumerator(), owned, 0));
        try {
            while (pending.Count != 0) {
                token.ThrowIfCancellationRequested();
                var level = pending.Peek();
                if (!level.Nodes.MoveNext()) { level.Nodes.Dispose(); pending.Pop(); continue; }
                XNode source = level.Nodes.Current;
                if (source is XElement element) {
                    tracker?.RecordElementStart(level.Depth + 1);
                    string ns = element.Name.NamespaceName;
                    if (ns == Dom.HtmlElement.HtmlNamespace && element.Name.LocalName != element.Name.LocalName.ToLowerInvariant())
                        throw new InvalidDataException("XHTML conversion requires canonical lowercase element names.");
                    var target = owned.CreateElement(element.Name.LocalName, ns, element.GetPrefixOfNamespace(element.Name.Namespace));
                    foreach (XAttribute attribute in element.Attributes()) {
                        token.ThrowIfCancellationRequested();
                        string name;
                        if (attribute.IsNamespaceDeclaration) {
                            name = attribute.Name.LocalName == "xmlns" ? "xmlns" : "xmlns:" + attribute.Name.LocalName;
                        } else {
                            if (ns == Dom.HtmlElement.HtmlNamespace && attribute.Name.NamespaceName.Length == 0 &&
                                attribute.Name.LocalName != attribute.Name.LocalName.ToLowerInvariant())
                                throw new InvalidDataException("XHTML conversion requires canonical lowercase unqualified attribute names.");
                            string? prefix = element.GetPrefixOfNamespace(attribute.Name.Namespace);
                            name = string.IsNullOrEmpty(prefix) ? attribute.Name.LocalName : prefix + ":" + attribute.Name.LocalName;
                        }
                        target.SetAttribute(name, attribute.Value, attribute.IsNamespaceDeclaration
                            ? "http://www.w3.org/2000/xmlns/" : attribute.Name.NamespaceName);
                    }
                    level.Parent.AppendChild(target);
                    // XML template children are ordinary nodes, not HTML parser-created template fragments.
                    pending.Push((element.Nodes().GetEnumerator(), target, level.Depth + 1));
                } else {
                    tracker?.RecordNode();
                    if (source is XText text) {
                        // XML permits formatting whitespace around its root; the HTML DOM does not.
                        if (!ReferenceEquals(level.Parent, owned) || !string.IsNullOrWhiteSpace(text.Value))
                            level.Parent.AppendChild(owned.CreateTextNode(text.Value));
                    } else if (source is XProcessingInstruction && skipProcessingInstructions) {
                        // Package editing retains these separately; they are not HTML resources.
                    } else if (source is XComment comment) level.Parent.AppendChild(owned.CreateComment(comment.Value));
                    else if (source is XDocumentType type) level.Parent.AppendChild(owned.CreateDocumentType(type.Name, type.PublicId ?? "", type.SystemId ?? ""));
                    else throw new InvalidDataException("XML conversion encountered an unsupported node.");
                }
            }
        } finally {
            while (pending.Count != 0) pending.Pop().Nodes.Dispose();
        }
        IHtmlDocument document = NativeDomBridge.GetNativeDocument(owned, token);
        HtmlConversionInputGuard.ValidateDocument(document, limits, token);
        return document;
    }
}
