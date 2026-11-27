using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word.OpenXml;

/// <summary>
/// Reads WordprocessingML through <see cref="OfficeXmlNode"/> with the results the SDK's typed members give, so the
/// converters read one implementation whether OfficeIMO's own tree or the SDK holds the document.
/// </summary>
/// <remarks>
/// Enumerated values are read as the XML token the attribute holds (<c>"center"</c>, <c>"auto"</c>), the same text the
/// SDK's enumeration values carry. On/off values follow the SDK: a missing <c>w:val</c> is on.
/// </remarks>
internal static class WordXml {
    internal const int W = OfficeXmlNamespaces.WordprocessingML;
    internal const int A = OfficeXmlNamespaces.DrawingML;
    internal const int R = OfficeXmlNamespaces.Relationships;
    internal const int M = OfficeXmlNamespaces.Math;
    internal const int Wp = OfficeXmlNamespaces.WordprocessingDrawing;
    internal const int Pic = OfficeXmlNamespaces.Picture;
    internal const int C = OfficeXmlNamespaces.Chart;
    internal const int Dgm = OfficeXmlNamespaces.Diagram;
    internal const int Mc = OfficeXmlNamespaces.MarkupCompatibility;
    internal const int V = OfficeXmlNamespaces.Vml;
    internal const int O = OfficeXmlNamespaces.VmlOffice;
    internal const int Wvml = OfficeXmlNamespaces.VmlWord;
    internal const int Xml = OfficeXmlNamespaces.Xml;

    internal const string W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    internal const string W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
    internal const string Wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    internal const string Wpg = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
    internal const string Wpc = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
    internal const string Wp14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
    internal const string A14 = "http://schemas.microsoft.com/office/drawing/2010/main";

    // ---- Navigation -----------------------------------------------------------------------------------------------

    /// <summary>The child elements, in document order.</summary>
    internal static IEnumerable<OfficeXmlNode> Elements(this OfficeXmlNode element) {
        for (OfficeXmlNode child = element.FirstChild; !child.IsNone; child = child.NextSibling) yield return child;
    }

    /// <summary>The child elements with this name, in document order.</summary>
    internal static IEnumerable<OfficeXmlNode> Elements(this OfficeXmlNode element, int namespaceId, string localName) {
        for (OfficeXmlNode child = element.FirstChild; !child.IsNone; child = child.NextSibling) {
            if (child.Is(namespaceId, localName)) yield return child;
        }
    }

    /// <summary>The child elements with this name in a namespace outside the fixed ones.</summary>
    internal static IEnumerable<OfficeXmlNode> Elements(this OfficeXmlNode element, string namespaceUri, string localName) {
        for (OfficeXmlNode child = element.FirstChild; !child.IsNone; child = child.NextSibling) {
            if (child.Is(namespaceUri, localName)) yield return child;
        }
    }

    /// <summary>The <c>w:</c> child elements with this name.</summary>
    internal static IEnumerable<OfficeXmlNode> WElements(this OfficeXmlNode element, string localName) => element.Elements(W, localName);

    /// <summary>The first <c>w:</c> child with this name, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode WChild(this OfficeXmlNode element, string localName) =>
        element.IsNone ? default : element.Child(W, localName);

    /// <summary>The last child element with this name, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode LastChild(this OfficeXmlNode element, int namespaceId, string localName) {
        for (OfficeXmlNode child = element.LastChild; !child.IsNone; child = child.PreviousSibling) {
            if (child.Is(namespaceId, localName)) return child;
        }
        return default;
    }

    /// <summary>Every element below this one, depth first in document order, as the SDK's <c>Descendants()</c>.</summary>
    internal static IEnumerable<OfficeXmlNode> DescendantElements(this OfficeXmlNode element) {
        if (element.IsNone) yield break;
        OfficeXmlNode node = element.FirstChild;
        while (!node.IsNone) {
            yield return node;
            OfficeXmlNode next = node.FirstChild;
            if (next.IsNone) {
                while (!node.IsNone && !node.Equals(element)) {
                    next = node.NextSibling;
                    if (!next.IsNone) break;
                    node = node.Parent;
                }
                if (node.IsNone || node.Equals(element)) yield break;
            }
            node = next;
        }
    }

    /// <summary>The elements with this name below this one, depth first in document order, read lazily.</summary>
    internal static IEnumerable<OfficeXmlNode> DescendantElements(this OfficeXmlNode element, int namespaceId, string localName) {
        foreach (OfficeXmlNode node in element.DescendantElements()) {
            if (node.Is(namespaceId, localName)) yield return node;
        }
    }

    /// <summary>The elements with this name, in a namespace outside the fixed ones, below this one.</summary>
    internal static IEnumerable<OfficeXmlNode> DescendantElements(this OfficeXmlNode element, string namespaceUri, string localName) {
        foreach (OfficeXmlNode node in element.DescendantElements()) {
            if (node.Is(namespaceUri, localName)) yield return node;
        }
    }

    /// <summary>The first element with this name below this one, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode FirstDescendant(this OfficeXmlNode element, int namespaceId, string localName) {
        foreach (OfficeXmlNode node in element.DescendantElements()) {
            if (node.Is(namespaceId, localName)) return node;
        }
        return default;
    }

    /// <summary>The first element with this name, in a namespace outside the fixed ones, below this one.</summary>
    internal static OfficeXmlNode FirstDescendant(this OfficeXmlNode element, string namespaceUri, string localName) {
        foreach (OfficeXmlNode node in element.DescendantElements()) {
            if (node.Is(namespaceUri, localName)) return node;
        }
        return default;
    }

    /// <summary>Whether an element with this name is below this one.</summary>
    internal static bool HasDescendant(this OfficeXmlNode element, int namespaceId, string localName) =>
        !element.FirstDescendant(namespaceId, localName).IsNone;

    /// <summary>The ancestors, nearest first, without the element itself.</summary>
    internal static IEnumerable<OfficeXmlNode> Ancestors(this OfficeXmlNode element) {
        for (OfficeXmlNode node = element.IsNone ? default : element.Parent; !node.IsNone; node = node.Parent) yield return node;
    }

    /// <summary>The nearest ancestor with this name, without the element itself, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode Ancestor(this OfficeXmlNode element, int namespaceId, string localName) =>
        element.IsNone ? default : element.Parent.AncestorOrSelf(namespaceId, localName);

    /// <summary>The following sibling elements, nearest first.</summary>
    internal static IEnumerable<OfficeXmlNode> ElementsAfter(this OfficeXmlNode element) {
        for (OfficeXmlNode node = element.IsNone ? default : element.NextSibling; !node.IsNone; node = node.NextSibling) yield return node;
    }

    /// <summary>The preceding sibling elements, nearest first.</summary>
    internal static IEnumerable<OfficeXmlNode> ElementsBefore(this OfficeXmlNode element) {
        for (OfficeXmlNode node = element.IsNone ? default : element.PreviousSibling; !node.IsNone; node = node.PreviousSibling) yield return node;
    }

    /// <summary>Whether the element is a <c>w:</c> element with this name.</summary>
    internal static bool IsW(this OfficeXmlNode element, string localName) => !element.IsNone && element.Is(W, localName);

    /// <summary>Whether the element has child elements.</summary>
    internal static bool HasChildren(this OfficeXmlNode element) => !element.IsNone && !element.FirstChild.IsNone;

    // ---- Values ---------------------------------------------------------------------------------------------------

    /// <summary>A <c>w:</c> attribute, or null when the element or the attribute is missing.</summary>
    internal static string? WAttr(this OfficeXmlNode element, string localName) =>
        element.IsNone ? null : element.Attribute(W, localName);

    /// <summary>The element's <c>w:val</c>, or null when the element or the value is missing.</summary>
    internal static string? Val(this OfficeXmlNode element) => element.IsNone ? null : element.Attribute(W, "val");

    /// <summary>The <c>w:val</c> of the first <c>w:</c> child with this name.</summary>
    internal static string? ChildVal(this OfficeXmlNode element, string localName) => element.WChild(localName).Val();

    /// <summary>An attribute with no namespace, or null.</summary>
    internal static string? Attr(this OfficeXmlNode element, string localName) =>
        element.IsNone ? null : element.Attribute(OfficeXmlNamespaces.None, localName);

    /// <summary>An attribute in a namespace outside the fixed ones, or null.</summary>
    internal static string? Attr(this OfficeXmlNode element, string namespaceUri, string localName) {
        if (element.IsNone) return null;
        if (OfficeXmlNamespaces.TryGetKnownId(namespaceUri, out int id)) return element.Attribute(id, localName);
        foreach (OfficeXmlAttribute attribute in element.Attributes()) {
            if (attribute.LocalName == localName && attribute.NamespaceUri == namespaceUri) return attribute.Value;
        }
        return null;
    }

    /// <summary>
    /// Whether an on/off element is on, as the SDK reads it: present without <c>w:val</c>, or with a true value. A
    /// missing element is off; a value the SDK can't parse throws as the SDK does.
    /// </summary>
    internal static bool IsOn(this OfficeXmlNode element) => WordXmlEdits.ReadOnOff(element);

    /// <summary>An on/off element's state: null when it is missing, otherwise as <see cref="IsOn"/>.</summary>
    internal static bool? OnOff(this OfficeXmlNode element) => element.IsNone ? null : WordXmlEdits.ReadOnOff(element);

    /// <summary>
    /// An on/off attribute value as the SDK's <c>OnOffValue</c> reads it: true, false, or null when missing. A value the
    /// SDK can't parse throws as the SDK does.
    /// </summary>
    internal static bool? ParseOnOff(string? value) => value switch {
        null => null,
        "true" or "1" or "on" => true,
        "false" or "0" or "off" => false,
        _ => throw new FormatException("The text value is invalid. It can be only 'true', 'false', 'on', 'off', '0', '1'.")
    };

    /// <summary>An on/off attribute, or null when it or the element is missing.</summary>
    internal static bool? OnOffAttr(this OfficeXmlNode element, int namespaceId, string localName) =>
        element.IsNone ? null : ParseOnOff(element.Attribute(namespaceId, localName));

    /// <summary>A <c>xsd:boolean</c> attribute as the SDK's <c>BooleanValue</c> reads it, or null.</summary>
    internal static bool? ParseBoolean(string? value) => value switch {
        null => null,
        "true" or "1" => true,
        "false" or "0" => false,
        _ => value.Trim() switch {
            "true" or "1" => true,
            "false" or "0" => false,
            _ => null
        }
    };

    /// <summary>An integer as the SDK's <c>Int32Value</c> reads it, or null when missing or not an integer.</summary>
    internal static int? ParseInt32(string? value) =>
        value != null && int.TryParse(value.Trim(), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int parsed) ? parsed : null;

    /// <summary>An unsigned integer as the SDK's <c>UInt32Value</c> reads it, or null.</summary>
    internal static uint? ParseUInt32(string? value) =>
        value != null && uint.TryParse(value.Trim(), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out uint parsed) ? parsed : null;

    /// <summary>A long integer as the SDK's <c>Int64Value</c> reads it, or null.</summary>
    internal static long? ParseInt64(string? value) =>
        value != null && long.TryParse(value.Trim(), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out long parsed) ? parsed : null;

    /// <summary>A double as the SDK's <c>DoubleValue</c> reads it, or null.</summary>
    internal static double? ParseDouble(string? value) => OfficeXmlViewValues.ParseDouble(value);

    /// <summary>A <c>w:</c> integer attribute, or null.</summary>
    internal static int? WInt32(this OfficeXmlNode element, string localName) => ParseInt32(element.WAttr(localName));

    /// <summary>A <c>w:</c> unsigned integer attribute, or null.</summary>
    internal static uint? WUInt32(this OfficeXmlNode element, string localName) => ParseUInt32(element.WAttr(localName));

    /// <summary>An unqualified integer attribute, or null.</summary>
    internal static int? Int32Attr(this OfficeXmlNode element, string localName) => ParseInt32(element.Attr(localName));

    /// <summary>An unqualified long integer attribute, or null.</summary>
    internal static long? Int64Attr(this OfficeXmlNode element, string localName) => ParseInt64(element.Attr(localName));

    /// <summary>The element's text and its descendants', as the SDK's <c>InnerText</c>.</summary>
    internal static string Text(this OfficeXmlNode element) => element.IsNone ? string.Empty : element.InnerText;
}
