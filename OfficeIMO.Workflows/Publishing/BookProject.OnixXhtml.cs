using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static XElement ReadOnixXhtmlFragment(string text, CancellationToken cancellationToken) {
        // ONIX's chameleon XHTML schema places inline markup in the ONIX namespace.
        // Parse bounded XML, not forgiving HTML; never retrieve a DTD, entity or linked resource.
        string wrapped = "<fragment xmlns=\"" + OnixNamespace + "\">" + text + "</fragment>";
        XElement fragment;
        try {
            using var input = new StringReader(wrapped);
            using var reader = XmlReader.Create(input, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 65536 + 256
            });
            int nodes = 0;
            while (reader.Read()) {
                cancellationToken.ThrowIfCancellationRequested();
                if ((reader.NodeType == XmlNodeType.Element && reader.Depth > 32) || (reader.Depth > 0 && reader.NodeType != XmlNodeType.EndElement && ++nodes > 4096))
                    throw new ArgumentException("ONIX XHTML exceeds 32 levels or 4096 XML nodes.", nameof(text));
                if (reader.NodeType is XmlNodeType.Comment or XmlNodeType.ProcessingInstruction or XmlNodeType.DocumentType)
                    throw new ArgumentException("ONIX XHTML does not accept comments, processing instructions or declarations.", nameof(text));
            }
            // The immutable string has passed depth/node bounds before tree materialization.
            fragment = XElement.Parse(wrapped, LoadOptions.PreserveWhitespace);
        } catch (XmlException error) {
            throw new ArgumentException("Supply a well-formed XHTML fragment without declarations or external entities.", nameof(text), error);
        }
        if (string.IsNullOrWhiteSpace(fragment.Value)) throw new ArgumentException("ONIX XHTML requires nonblank textual content.", nameof(text));
        XNamespace ns = OnixNamespace;
        foreach (var element in fragment.Descendants()) {
            cancellationToken.ThrowIfCancellationRequested();
            string tag = element.Name.LocalName;
            string namespaceName = element.Name.NamespaceName;
            if ((namespaceName != "" && namespaceName != OnixNamespace && namespaceName != "http://www.w3.org/1999/xhtml") || !IsOnixCollateralTag(tag))
                throw new ArgumentException("Unsupported ONIX XHTML element: " + element.Name, nameof(text));
            foreach (var attribute in element.Attributes().ToArray()) {
                if (attribute.IsNamespaceDeclaration) { attribute.Remove(); continue; }
                if (attribute.Name == XNamespace.Xml + "lang") {
                    var existing = element.Attribute("lang");
                    if (existing != null && existing.Value != attribute.Value)
                        throw new ArgumentException("XHTML lang and xml:lang must agree.", nameof(text));
                    if (existing == null) element.SetAttributeValue("lang", attribute.Value);
                    attribute.Remove();
                    continue;
                }
                string name = attribute.Name.LocalName;
                if (attribute.Name.NamespaceName.Length != 0 || !IsOnixCollateralAttribute(tag, name))
                    throw new ArgumentException("Unsupported ONIX XHTML attribute: " + attribute.Name, nameof(text));
                if (name is "href" or "cite") RequireOnixHttpUrl(attribute.Value, name);
            }
            element.Name = ns + tag;
        }
        return fragment;
    }

    private static bool IsOnixCollateralTag(string tag) => tag is
        "div" or "p" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6" or
        "ul" or "ol" or "li" or "dl" or "dt" or "dd" or "blockquote" or "pre" or
        "a" or "span" or "bdo" or "br" or "hr" or "em" or "strong" or "b" or "i" or
        "dfn" or "code" or "samp" or "kbd" or "var" or "cite" or "abbr" or "acronym" or "q" or "sub" or "sup" or "tt" or
        "table" or "caption" or "thead" or "tbody" or "tfoot" or "tr" or "th" or "td" or "colgroup" or "col";

    private static bool IsOnixCollateralAttribute(string tag, string name) => name is "title" or "lang" or "dir" ||
        (tag == "a" && name is "href" or "hreflang") ||
        (tag is "q" or "blockquote" && name == "cite") ||
        (tag == "ol" && name is "start" or "type") ||
        (tag is "td" or "th" && name is "colspan" or "rowspan") ||
        (tag == "th" && name == "scope") || (tag is "col" or "colgroup" && name == "span");
}
