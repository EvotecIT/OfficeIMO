using AngleSharp.Dom;
using OfficeIMO.Html;
using System.Threading;

namespace OfficeIMO.Epub;

public static partial class EpubManuscript {
    private static readonly HashSet<string> OmittedElements = new HashSet<string>(new[] {
        "script", "iframe", "object", "embed", "form", "input", "button", "select", "textarea", "base", "link", "meta", "style"
    }, StringComparer.OrdinalIgnoreCase);

    private static XElement? ConvertElement(IElement source, List<OfficeConversionFidelityDiagnostic> diagnostics,
        HtmlConversionDocument manuscript, HashSet<string> anchorIds, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        string name = source.LocalName;
        if (name == "input" && string.Equals(source.GetAttribute("type"), "checkbox", StringComparison.OrdinalIgnoreCase) && source.HasAttribute("disabled"))
            return new XElement(Xhtml + "span", new XAttribute("class", "task-state"), source.HasAttribute("checked") ? "[x] " : "[ ] ");
        if (OmittedElements.Contains(name)) {
            if (name != "style" && name != "base" && name != "link" && name != "meta")
                AddDiagnostic(diagnostics, "EPUB_IMPORT_ACTIVE_CONTENT_OMITTED", "Interactive or executable content was omitted from the reflowable manuscript.", name);
            return null;
        }
        XNamespace ns = source.NamespaceUri ?? Xhtml.NamespaceName;
        var element = new XElement(ns + name);
        foreach (var attribute in source.Attributes) {
            string localName = attribute.LocalName;
            if (localName.StartsWith("on", StringComparison.OrdinalIgnoreCase) || localName == "srcdoc") {
                AddDiagnostic(diagnostics, "EPUB_IMPORT_ACTIVE_ATTRIBUTE_OMITTED", "An executable attribute was omitted.", name + "@" + attribute.Name);
                continue;
            }
            if (attribute.Name == "xmlns" || attribute.Prefix == "xmlns") continue;
            XName attributeName;
            if ((attribute.NamespaceUri == null || attribute.NamespaceUri.Length == 0) && localName.IndexOf(':') >= 0) {
                string[] parts = localName.Split(':');
                string? attributeNamespace = parts.Length != 2 ? null : parts[0] switch {
                    "epub" => "http://www.idpf.org/2007/ops", "xml" => XNamespace.Xml.NamespaceName,
                    "xlink" => "http://www.w3.org/1999/xlink", _ => null
                };
                if (attributeNamespace == null) {
                    AddDiagnostic(diagnostics, "EPUB_IMPORT_ATTRIBUTE_OMITTED", "An undeclared prefixed attribute could not be represented in XHTML.", name + "@" + attribute.Name);
                    continue;
                }
                attributeName = XName.Get(parts[1], attributeNamespace);
            } else attributeName = attribute.NamespaceUri == null || attribute.NamespaceUri.Length == 0
                ? XName.Get(localName) : XName.Get(localName, attribute.NamespaceUri);
            try {
                XmlConvert.VerifyXmlChars(attribute.Value);
                element.SetAttributeValue(attributeName, attribute.Value);
            } catch (XmlException) {
                AddDiagnostic(diagnostics, "EPUB_IMPORT_ATTRIBUTE_OMITTED", "An attribute could not be represented in XML.", name + "@" + attribute.Name);
            }
        }
        foreach (INode node in source.ChildNodes) {
            token.ThrowIfCancellationRequested();
            if (node is IElement child) {
                XElement? converted = ConvertElement(child, diagnostics, manuscript, anchorIds, token);
                if (converted != null) element.Add(converted);
            } else if (node is IText text) {
                XmlConvert.VerifyXmlChars(text.Data);
                element.Add(new XText(text.Data));
            }
        }
        NormalizeLegacyHtml(element, diagnostics);
        // HTML 4 named anchors are common in compiled help and older manuals. EPUB
        // navigation addresses XML IDs; keep both destinations when id and name differ.
        if (name == "a" && element.Attribute("name") is XAttribute namedAnchor) {
            string anchor = namedAnchor.Value;
            namedAnchor.Remove();
            if (anchor.Length != 0 && (string?)element.Attribute("id") != anchor) {
                if (!anchorIds.Add(anchor)) {
                    AddDiagnostic(diagnostics, "EPUB_IMPORT_ID_REFERENCE_INVALID", "A legacy named anchor collides with an existing destination; its alias was omitted.", anchor, OfficeConversionLossKind.Failure);
                }
                else if (element.Attribute("id") == null) element.SetAttributeValue("id", anchor);
                else
                    element.AddFirst(new XElement(Xhtml + "span", new XAttribute("id", anchor)));
            }
        }
        if (name == "img" && element.Attribute("alt") == null) {
            string role = source.GetAttribute("role")?.Trim() ?? string.Empty;
            string accessibleName = HtmlAccessibilitySemantics.GetImageAccessibleName(source);
            if (accessibleName.Length > 0) element.SetAttributeValue("alt", accessibleName);
            else if (role.Equals("presentation", StringComparison.OrdinalIgnoreCase) || role.Equals("none", StringComparison.OrdinalIgnoreCase))
                element.SetAttributeValue("alt", string.Empty);
            else AddDiagnostic(diagnostics, "EPUB_IMPORT_IMAGE_ALT_MISSING", "The source image needs alternative text or an explicit decorative role. Repair the source and import it again.", source.GetAttribute("src"), OfficeConversionLossKind.Failure);
        }
        return element;
    }
}
