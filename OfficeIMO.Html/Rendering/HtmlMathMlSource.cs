using System.IO;
using System.Xml;
using System.Text;
using System.Xml.Linq;
using AngleSharp.Dom;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Html;

/// <summary>Recoverable namespace-aware MathML for a rendered formula.</summary>
public sealed class HtmlMathMlSource {
    private HtmlMathMlSource(string mathMl, bool isOriginalMarkup) {
        MathMl = mathMl;
        IsOriginalMarkup = isOriginalMarkup;
    }
    /// <summary>Standalone MathML XML. Exact caller markup is retained when it represents the current formula.</summary>
    public string MathMl { get; }
    /// <summary>Whether MathMl retains the caller's exact lexical markup, including quotes and whitespace.</summary>
    /// <remarks>Edited, recovered or non-XML HTML MathML uses the current namespace-aware serialization instead.</remarks>
    public bool IsOriginalMarkup { get; }

    internal static HtmlMathMlSource Create(IElement element, int maxDepth, int maxNodes, CancellationToken cancellationToken, out XElement layoutRoot) {
        int nodes = 0;
        XElement current = Export(element, 1, maxDepth, maxNodes, ref nodes, cancellationToken);
        layoutRoot = current;
        string? original = NativeSourceMarkup.Get(element)?.Markup;
        if (original != null) {
            try {
                using var input = new StringReader(original);
                using var reader = XmlReader.Create(input, new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
                    MaxCharactersInDocument = original.Length
                });
                using var bounded = new OfficeXmlLimitingReader(reader, "MathML source", maxDepth, maxNodes,
                    (int)Math.Min(int.MaxValue, 16L * maxNodes), cancellationToken);
                XElement parsed = XElement.Load(bounded, LoadOptions.PreserveWhitespace);
                Normalize(parsed);
                Normalize(current);
                if (XNode.DeepEquals(parsed, current)) return new HtmlMathMlSource(original, true);
            } catch (Exception exception) when (exception is XmlException or InvalidDataException) {
                // Valid HTML foreign content need not be standalone XML (named HTML entities,
                // unquoted attributes, omitted namespace). A stale source may also exceed
                // comparison limits after editing. Preserve the bounded current DOM instead.
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new HtmlMathMlSource(current.ToString(SaveOptions.DisableFormatting), false);
    }

    private static XElement Export(IElement element, int depth, int maxDepth, int maxNodes, ref int nodes, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        CountNode(depth, maxDepth, maxNodes, ref nodes);
        var result = new XElement(XName.Get(element.LocalName, element.NamespaceUri ?? string.Empty));
        result.AddAnnotation(element);
        foreach (IAttr attribute in element.Attributes) {
            cancellationToken.ThrowIfCancellationRequested();
            if (attribute.NamespaceUri == "http://www.w3.org/2000/xmlns/" || attribute.Name == "xmlns" || attribute.Prefix == "xmlns") continue;
            result.Add(new XAttribute(XName.Get(attribute.LocalName, attribute.NamespaceUri ?? string.Empty), attribute.Value));
        }
        foreach (INode child in element.ChildNodes) {
            cancellationToken.ThrowIfCancellationRequested();
            if (child is IElement nested) result.Add(Export(nested, depth + 1, maxDepth, maxNodes, ref nodes, cancellationToken));
            else {
                CountNode(depth, maxDepth, maxNodes, ref nodes);
                if (child.NodeType == NodeType.Text) result.Add(new XText(child.TextContent));
                else if (child.NodeType == NodeType.Comment) result.Add(new XComment(child.TextContent));
                else throw new XmlException("The MathML DOM contains a node that cannot be represented as XML.");
            }
        }
        return result;
    }

    private static void CountNode(int depth, int maxDepth, int maxNodes, ref int nodes) {
        if (depth > maxDepth || ++nodes > maxNodes)
            throw new InvalidDataException("MathML source exceeds the configured DOM depth or node budget.");
    }

    private static void Normalize(XElement root) {
        foreach (XElement element in root.DescendantsAndSelf()) {
            element.Attributes().Where(attribute => attribute.IsNamespaceDeclaration).Remove();
            // XML CDATA and HTML character data represent the same text. Coalesce adjacent
            // text nodes so lexical CDATA boundaries cannot turn an unchanged formula stale.
            XText? previous = null;
            StringBuilder? joined = null;
            foreach (XNode node in element.Nodes().ToArray()) {
                if (node is XText text) {
                    if (previous == null) {
                        if (text is XCData) { previous = new XText(text.Value); text.ReplaceWith(previous); }
                        else previous = text;
                    } else {
                        joined ??= new StringBuilder(previous.Value);
                        joined.Append(text.Value);
                        text.Remove();
                    }
                } else {
                    if (joined != null) previous!.Value = joined.ToString();
                    previous = null;
                    joined = null;
                }
            }
            if (joined != null) previous!.Value = joined.ToString();
        }
    }
}
