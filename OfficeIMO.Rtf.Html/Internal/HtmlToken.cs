using AngleSharp.Dom;

namespace OfficeIMO.Rtf.Html;

internal sealed class HtmlToken {
    private HtmlToken(string value, IReadOnlyDictionary<string, string>? attributes = null, IElement? element = null) {
        Value = value;
        Attributes = attributes ?? EmptyAttributes;
        Element = element;
    }

    internal static HtmlToken FromElement(IElement element) {
        var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (IAttr attribute in element.Attributes) {
            attributes[attribute.Name] = attribute.Value;
        }

        return new HtmlToken(element.LocalName.ToLowerInvariant(), attributes, element);
    }

    internal string Value { get; }

    internal IReadOnlyDictionary<string, string> Attributes { get; }

    internal IElement? Element { get; }

    private static readonly IReadOnlyDictionary<string, string> EmptyAttributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
}
