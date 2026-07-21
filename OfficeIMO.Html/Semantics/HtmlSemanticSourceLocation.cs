using AngleSharp.Dom;
using System.Globalization;

namespace OfficeIMO.Html;

/// <summary>Stable provenance for a semantic value interpreted from source HTML.</summary>
public sealed class HtmlSemanticSourceLocation {
    internal HtmlSemanticSourceLocation(string selector, string elementName, int line, int column, int index) {
        Selector = selector ?? string.Empty;
        ElementName = elementName ?? string.Empty;
        Line = line;
        Column = column;
        Index = index;
    }

    /// <summary>Deterministic CSS-like path identifying the source element.</summary>
    public string Selector { get; }

    /// <summary>Lower-case source element name.</summary>
    public string ElementName { get; }

    /// <summary>One-based source line, or zero when the parser has no source reference.</summary>
    public int Line { get; }

    /// <summary>One-based source column, or zero when the parser has no source reference.</summary>
    public int Column { get; }

    /// <summary>Zero-based character index, or -1 when unavailable.</summary>
    public int Index { get; }

    /// <inheritdoc />
    public override string ToString() => Line > 0
        ? Selector + " (line " + Line + ", column " + Column + ")"
        : Selector;

    internal static HtmlSemanticSourceLocation FromElement(IElement element) {
        if (element == null) throw new ArgumentNullException(nameof(element));
        string selector = BuildSelector(element);
        ISourceReference? source = element.SourceReference;
        if (source == null) return new HtmlSemanticSourceLocation(selector, element.LocalName, 0, 0, -1);
        return new HtmlSemanticSourceLocation(
            selector,
            element.LocalName,
            source.Position.Line,
            source.Position.Column,
            source.Position.Index);
    }

    private static string BuildSelector(IElement element) {
        var segments = new Stack<string>();
        for (IElement? current = element; current != null; current = current.ParentElement) {
            string segment = current.LocalName.ToLowerInvariant();
            if (!string.IsNullOrWhiteSpace(current.Id)) {
                segment += "#" + current.Id;
                segments.Push(segment);
                break;
            }
            if (current.ParentElement != null) {
                int ordinal = current.ParentElement.Children
                    .Where(sibling => string.Equals(sibling.LocalName, current.LocalName, StringComparison.OrdinalIgnoreCase))
                    .TakeWhile(sibling => !ReferenceEquals(sibling, current)).Count() + 1;
                segment += ":nth-of-type(" + ordinal.ToString(CultureInfo.InvariantCulture) + ")";
            }
            segments.Push(segment);
        }
        return string.Join(" > ", segments);
    }
}
