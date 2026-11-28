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
/// SDK's enumeration values carry. On/off values follow the SDK: a missing <c>w:val</c> is on. Names that read a
/// <c>w:</c> attribute start with <c>W</c> (<see cref="WVal"/>, <see cref="WAttr"/>), so they never meet the shared chart
/// readers' unqualified <c>Val()</c>; <c>FirstDescendant</c> and <c>Ancestor</c> with a namespace id come from the shared
/// <see cref="OfficeChartXml"/> helpers.
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

    /// <summary>
    /// The first element with this name below this one, in document order, or <see cref="OfficeXmlNode.None"/>. It
    /// walks every element, both branches of <c>mc:AlternateContent</c> included, as the SDK's typed
    /// <c>Descendants&lt;T&gt;()</c> does for a document opened without markup compatibility processing. (The shared
    /// <c>OfficeChartXml.FirstDescendant</c> leaves alternate content out, which suits chart parts, not stories.)
    /// </summary>
    internal static OfficeXmlNode FirstDescendantElement(this OfficeXmlNode element, int namespaceId, string localName) {
        foreach (OfficeXmlNode node in element.DescendantElements()) {
            if (node.Is(namespaceId, localName)) return node;
        }
        return default;
    }

    /// <summary>The first element with this name, in a namespace outside the fixed ones, below this one.</summary>
    internal static OfficeXmlNode FirstDescendantElement(this OfficeXmlNode element, string namespaceUri, string localName) {
        foreach (OfficeXmlNode node in element.DescendantElements()) {
            if (node.Is(namespaceUri, localName)) return node;
        }
        return default;
    }

    /// <summary>Whether an element with this name is below this one, alternate content included.</summary>
    internal static bool HasDescendant(this OfficeXmlNode element, int namespaceId, string localName) =>
        !element.FirstDescendantElement(namespaceId, localName).IsNone;

    /// <summary>
    /// Every element below this one, depth first in document order, without entering extension lists
    /// (<c>extLst</c>). Extension lists carry producer metadata, such as filtered chart series or properties of a
    /// later Office version, that can repeat the very elements a reader looks for. The SDK reads an extension it
    /// doesn't know as untyped elements, so its typed lookups never matched inside one; DrawingML and chart reads
    /// use this walk to stay on the content itself. Unlike the shared <c>OfficeXmlNode.ContentDescendants</c>, it
    /// keeps <c>mc:AlternateContent</c>: the SDK typed both branches, and Word stories hold their drawings there.
    /// </summary>
    internal static IEnumerable<OfficeXmlNode> DescendantsSkippingExtLst(this OfficeXmlNode element) {
        if (element.IsNone) yield break;
        OfficeXmlNode node = element.FirstChild;
        while (!node.IsNone) {
            bool extension = node.LocalName == "extLst";
            if (!extension) yield return node;
            OfficeXmlNode next = extension ? default : node.FirstChild;
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

    /// <summary>The elements with this name below this one, outside extension lists (see <see cref="DescendantsSkippingExtLst(OfficeXmlNode)"/>).</summary>
    internal static IEnumerable<OfficeXmlNode> DescendantsSkippingExtLst(this OfficeXmlNode element, int namespaceId, string localName) {
        foreach (OfficeXmlNode node in element.DescendantsSkippingExtLst()) {
            if (node.Is(namespaceId, localName)) yield return node;
        }
    }

    /// <summary>The elements with this name, in a namespace outside the fixed ones, below this one and outside extension lists.</summary>
    internal static IEnumerable<OfficeXmlNode> DescendantsSkippingExtLst(this OfficeXmlNode element, string namespaceUri, string localName) {
        foreach (OfficeXmlNode node in element.DescendantsSkippingExtLst()) {
            if (node.Is(namespaceUri, localName)) yield return node;
        }
    }

    /// <summary>The first element with this name below this one and outside extension lists, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode FirstDescendantSkippingExtLst(this OfficeXmlNode element, int namespaceId, string localName) {
        foreach (OfficeXmlNode node in element.DescendantsSkippingExtLst()) {
            if (node.Is(namespaceId, localName)) return node;
        }
        return default;
    }

    /// <summary>The first element with this name, in a namespace outside the fixed ones, below this one and outside extension lists.</summary>
    internal static OfficeXmlNode FirstDescendantSkippingExtLst(this OfficeXmlNode element, string namespaceUri, string localName) {
        foreach (OfficeXmlNode node in element.DescendantsSkippingExtLst()) {
            if (node.Is(namespaceUri, localName)) return node;
        }
        return default;
    }

    /// <summary>Whether an element with this name is below this one, outside extension lists.</summary>
    internal static bool HasDescendantSkippingExtLst(this OfficeXmlNode element, int namespaceId, string localName) =>
        !element.FirstDescendantSkippingExtLst(namespaceId, localName).IsNone;

    /// <summary>The ancestors, nearest first, without the element itself.</summary>
    internal static IEnumerable<OfficeXmlNode> Ancestors(this OfficeXmlNode element) {
        for (OfficeXmlNode node = element.IsNone ? default : element.Parent; !node.IsNone; node = node.Parent) yield return node;
    }

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

    // ---- Classification -------------------------------------------------------------------------------------------

    /// <summary>
    /// The level of a <c>w:sdt</c> by where it sits, as the SDK types it (<c>SdtRun</c>, <c>SdtBlock</c>,
    /// <c>SdtRow</c>, <c>SdtCell</c>): content controls, custom XML and markup-compatibility branches pass their own
    /// context on to their content. <see cref="WordContentControlLevel.None"/> where the SDK leaves it untyped.
    /// </summary>
    internal static WordContentControlLevel ContentControlLevel(this OfficeXmlNode sdt) {
        if (!sdt.IsW("sdt")) return WordContentControlLevel.None;
        OfficeXmlNode parent = sdt.Parent;
        while (true) {
            while (parent.Is(Mc, "Choice") || parent.Is(Mc, "Fallback") || parent.Is(Mc, "AlternateContent")) parent = parent.Parent;
            if (parent.IsW("sdtContent")) return parent.Parent.ContentControlLevel();
            if (!parent.IsW("customXml")) break;
            parent = parent.Parent;
        }
        if (parent.IsNone) return WordContentControlLevel.None;
        if (parent.NamespaceId == M) {
            switch (parent.LocalName) {
                case "oMath": case "e": case "num": case "den": case "sub": case "sup": case "deg": case "lim": case "fName":
                    return WordContentControlLevel.Run;
                default:
                    return WordContentControlLevel.None;
            }
        }
        // Run-level conflict marks from merged revisions hold runs, so the SDK types their controls as SdtRun too.
        if (parent.Is(W14, "conflictIns") || parent.Is(W14, "conflictDel")) return WordContentControlLevel.Run;
        if (parent.NamespaceId != W) return WordContentControlLevel.None;
        switch (parent.LocalName) {
            case "p": case "hyperlink": case "fldSimple": case "dir": case "bdo":
            case "ins": case "del": case "moveFrom": case "moveTo":
                return WordContentControlLevel.Run;
            case "body": case "txbxContent": case "tc": case "hdr": case "ftr": case "footnote": case "endnote":
            case "comment": case "docPartBody":
                return WordContentControlLevel.Block;
            case "tr":
                return WordContentControlLevel.Cell;
            case "tbl":
                return WordContentControlLevel.Row;
            default:
                return WordContentControlLevel.None;
        }
    }

    /// <summary>Whether the element is an inline <c>w:sdt</c>, one the SDK types as <c>SdtRun</c>.</summary>
    internal static bool IsRunContentControl(this OfficeXmlNode element) => element.ContentControlLevel() == WordContentControlLevel.Run;

    /// <summary>
    /// Whether the element is a tracked change that holds content (<c>w:ins</c>, <c>w:del</c>, <c>w:moveFrom</c> or
    /// <c>w:moveTo</c>, the SDK's <c>InsertedRun</c> family), not the same-named revision mark inside run, row,
    /// numbering or math-control properties.
    /// </summary>
    internal static bool IsTrackedChangeContainer(this OfficeXmlNode element) {
        if (element.IsNone || element.NamespaceId != W) return false;
        switch (element.LocalName) {
            case "ins": case "del": case "moveFrom": case "moveTo":
                OfficeXmlNode parent = element.Parent;
                return !(parent.IsW("rPr") || parent.IsW("trPr") || parent.IsW("numPr") || parent.IsW("smartTag") ||
                    parent.Is(M, "ctrlPr"));
            default:
                return false;
        }
    }

    // ---- Values ---------------------------------------------------------------------------------------------------

    /// <summary>A <c>w:</c> attribute, or null when the element or the attribute is missing.</summary>
    internal static string? WAttr(this OfficeXmlNode element, string localName) =>
        element.IsNone ? null : element.Attribute(W, localName);

    /// <summary>The element's <c>w:val</c>, or null when the element or the value is missing.</summary>
    internal static string? WVal(this OfficeXmlNode element) => element.IsNone ? null : element.Attribute(W, "val");

    /// <summary>The <c>w:val</c> of the first <c>w:</c> child with this name.</summary>
    internal static string? ChildVal(this OfficeXmlNode element, string localName) => element.WChild(localName).WVal();

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

    /// <summary>A short integer as the SDK's <c>Int16Value</c> reads it, or null when missing or out of range.</summary>
    internal static short? ParseInt16(string? value) =>
        value != null && short.TryParse(value.Trim(), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out short parsed) ? parsed : null;

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

/// <summary>The level of a <c>w:sdt</c> content control, as <see cref="WordXml.ContentControlLevel"/> reads it.</summary>
internal enum WordContentControlLevel {
    /// <summary>Not a content control the SDK types by position.</summary>
    None,
    /// <summary>An inline control in paragraph content (<c>SdtRun</c>).</summary>
    Run,
    /// <summary>A block-level control (<c>SdtBlock</c>).</summary>
    Block,
    /// <summary>A control holding table rows (<c>SdtRow</c>).</summary>
    Row,
    /// <summary>A control holding table cells (<c>SdtCell</c>).</summary>
    Cell
}
