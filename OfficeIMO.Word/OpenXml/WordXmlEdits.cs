using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word.OpenXml;

/// <summary>
/// WordprocessingML edits over <see cref="OfficeXmlNode"/>, with the results the SDK's typed members give, so one
/// implementation serves a document in OfficeIMO's own tree and one the SDK holds.
/// </summary>
internal static class WordXmlEdits {
    private const int W = OfficeXmlNamespaces.WordprocessingML;

    /// <summary>
    /// Puts a new, empty element in place of the first child with this name, or inserts one in schema order, as
    /// assigning a new element to a typed SDK property does, and returns it.
    /// </summary>
    internal static OfficeXmlNode ReplaceChild(OfficeXmlNode parent, string localName, OfficeXmlChildOrder order) {
        OfficeXmlNode replacement = parent.CreateElement(W, localName);
        OfficeXmlNode existing = parent.Child(W, localName);
        if (existing.IsNone) return parent.InsertOrdered(replacement, order);
        replacement = parent.InsertBefore(replacement, existing);
        existing.Remove();
        return replacement;
    }

    /// <summary>Like <see cref="ReplaceChild(OfficeXmlNode, string, OfficeXmlChildOrder)"/>, with a <c>w:val</c> value.</summary>
    internal static OfficeXmlNode ReplaceChild(OfficeXmlNode parent, string localName, OfficeXmlChildOrder order, string value) {
        OfficeXmlNode replacement = ReplaceChild(parent, localName, order);
        replacement.SetAttribute(W, "val", value);
        return replacement;
    }

    /// <summary>Removes the first child with this name, as clearing a typed SDK property does.</summary>
    internal static void RemoveChild(OfficeXmlNode parent, string localName) {
        if (parent.IsNone) return;
        OfficeXmlNode existing = parent.Child(W, localName);
        if (!existing.IsNone) existing.Remove();
    }

    /// <summary>
    /// Reads an on/off property the way OfficeIMO's SDK path does: present with no value, or a true value, is on. A
    /// value the SDK doesn't parse throws as the SDK does.
    /// </summary>
    internal static bool ReadOnOff(OfficeXmlNode element) {
        if (element.IsNone) return false;
        string? token = element.Attribute(W, "val");
        switch (token) {
            case null: case "true": case "1": case "on":
                return true;
            case "false": case "0": case "off":
                return false;
            default:
                throw new FormatException("The text value is invalid. It can be only 'true', 'false', 'on', 'off', '0', '1'.");
        }
    }

    /// <summary>The <c>w:val</c> of a child element, or null when there is no such child or value.</summary>
    internal static string? ChildValue(OfficeXmlNode parent, string localName) {
        if (parent.IsNone) return null;
        OfficeXmlNode child = parent.Child(W, localName);
        return child.IsNone ? null : child.Attribute(W, "val");
    }

    /// <summary>A parsed token, or the exception the SDK path throws for a token it doesn't map.</summary>
    internal static TEnum Token<TEnum>(bool parsed, TEnum value, string token) where TEnum : struct, Enum =>
        parsed ? value : throw new ArgumentOutOfRangeException(nameof(token), token, $"Unsupported {typeof(TEnum).Name} token.");
}
