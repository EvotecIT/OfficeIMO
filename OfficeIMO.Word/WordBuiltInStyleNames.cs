namespace OfficeIMO.Word;

/// <summary>
/// Word's built-in styles by their English identifier and their culture-independent built-in name. Word stores
/// built-in styles under the English built-in name in <c>w:name</c>, but localizes <c>w:styleId</c>: a French
/// template defines <c>heading 1</c> as <c>Titre1</c> and <c>Quote</c> as <c>Citation</c>.
/// </summary>
internal static class WordBuiltInStyleNames {
    private static readonly Dictionary<string, string> NameById = CreateNames();
    private static readonly Dictionary<string, string> IdByName = NameById.ToDictionary(pair => pair.Value, pair => pair.Key, StringComparer.OrdinalIgnoreCase);

    private static Dictionary<string, string> CreateNames() {
        var names = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["Normal"] = "Normal",
            ["Title"] = "Title",
            ["Subtitle"] = "Subtitle",
            ["Quote"] = "Quote",
            ["IntenseQuote"] = "Intense Quote",
            ["ListParagraph"] = "List Paragraph",
            ["Caption"] = "caption",
            ["NoSpacing"] = "No Spacing",
            ["TOCHeading"] = "TOC Heading",
            ["Header"] = "header",
            ["Footer"] = "footer",
            ["FootnoteText"] = "footnote text",
            ["EndnoteText"] = "endnote text",
            ["Hyperlink"] = "Hyperlink",
            ["TableGrid"] = "Table Grid"
        };
        for (int level = 1; level <= 9; level++) {
            string digit = level.ToString(System.Globalization.CultureInfo.InvariantCulture);
            names["Heading" + digit] = "heading " + digit;
            names["TOC" + digit] = "toc " + digit;
        }
        return names;
    }

    /// <summary>Returns the <c>w:type</c> a built-in style has: <c>character</c>, <c>table</c> or <c>paragraph</c>.</summary>
    internal static string GetStyleType(string builtInId) => builtInId switch {
        "Hyperlink" => "character",
        "TableGrid" => "table",
        _ => "paragraph"
    };

    /// <summary>Returns the built-in name of an English built-in style identifier.</summary>
    internal static bool TryGetName(string styleId, out string name) => NameById.TryGetValue(styleId, out name!);

    /// <summary>Returns the English identifier of a built-in style name, ignoring case and surrounding spaces.</summary>
    internal static string? GetId(string? name) =>
        name != null && IdByName.TryGetValue(name.Trim(), out string? id) ? id : null;

    /// <summary>
    /// Returns the English identifier of the built-in style a definition is, through its built-in name, or its
    /// identifier when that is an English built-in identifier and the name names no other built-in style.
    /// </summary>
    internal static string? GetBuiltInId(DocumentFormat.OpenXml.Wordprocessing.Style style) =>
        Classify(style.StyleId?.Value, style.StyleName?.Val?.Value, StyleType(style));

    /// <summary>
    /// Returns the English identifier of the built-in style a definition with this identifier, name and type is. The
    /// type must match the built-in style's type, so a character style named <c>Quote</c> is not the quote style.
    /// </summary>
    internal static string? Classify(string? styleId, string? name, string? type) {
        string effectiveType = string.IsNullOrEmpty(type) ? "paragraph" : type!;
        if (GetId(name) is { } byName && string.Equals(GetStyleType(byName), effectiveType, StringComparison.Ordinal)) return byName;
        return IsBuiltInId(styleId) && string.Equals(GetStyleType(styleId!), effectiveType, StringComparison.Ordinal) ? styleId : null;
    }

    /// <summary>The <c>w:type</c> of a style definition as written, or null when it has none.</summary>
    internal static string? StyleType(DocumentFormat.OpenXml.Wordprocessing.Style style) =>
        style.Type == null ? null : style.Type.InnerText;

    /// <summary>Returns whether an identifier is the English identifier of a built-in style.</summary>
    internal static bool IsBuiltInId(string? styleId) => styleId != null && NameById.ContainsKey(styleId);

    /// <summary>Returns the heading level, 1 to 9, of an English heading identifier; otherwise 0.</summary>
    internal static int GetHeadingLevel(string? builtInId) =>
        builtInId is { Length: 8 } && builtInId.StartsWith("Heading", StringComparison.Ordinal) && builtInId[7] is >= '1' and <= '9'
            ? builtInId[7] - '0'
            : 0;

    /// <summary>Returns the level, 1 to 9, of an English table of contents entry identifier; otherwise 0.</summary>
    internal static int GetTocLevel(string? builtInId) =>
        builtInId is { Length: 4 } && builtInId.StartsWith("TOC", StringComparison.Ordinal) && builtInId[3] is >= '1' and <= '9'
            ? builtInId[3] - '0'
            : 0;
}
