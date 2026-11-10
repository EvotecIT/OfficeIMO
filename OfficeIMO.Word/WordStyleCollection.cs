using System.Collections;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>
/// The styles defined in one <see cref="WordDocument"/>: list, inspect, create, change, remove and import
/// paragraph, character and table styles. Every change is made in the document's own style definitions.
/// </summary>
public sealed partial class WordStyleCollection : IEnumerable<WordStyle> {
    private readonly WordDocument _document;

    internal WordStyleCollection(WordDocument document) {
        _document = document;
    }

    /// <summary>Gets the number of styles the document defines.</summary>
    public int Count => Definitions.Count();

    /// <summary>Gets a style by identifier.</summary>
    /// <exception cref="KeyNotFoundException">The document defines no style with this identifier.</exception>
    public WordStyle this[string styleId] => FindById(styleId)
        ?? throw new KeyNotFoundException($"The document defines no style with the identifier '{styleId}'.");

    /// <summary>Finds a style by identifier, or by name when no identifier matches. Names ignore case.</summary>
    /// <returns>The style, or <see langword="null"/> when the document defines none.</returns>
    public WordStyle? Get(string styleIdOrName) {
        ArgumentNullException.ThrowIfNull(styleIdOrName);
        return FindById(styleIdOrName) ?? FindByName(styleIdOrName);
    }

    /// <summary>Returns whether the document defines a style with this identifier.</summary>
    public bool Contains(string styleId) => FindById(styleId) != null;

    /// <summary>Adds a paragraph style, based on the document's default paragraph style.</summary>
    /// <param name="styleId">The identifier paragraphs store to use the style.</param>
    /// <param name="name">The name shown in Word; defaults to <paramref name="styleId"/>.</param>
    /// <returns>The new style, ready to configure.</returns>
    /// <exception cref="ArgumentException">The identifier or name is already in use.</exception>
    public WordStyle AddParagraphStyle(string styleId, string? name = null) => Add(StyleValues.Paragraph, styleId, name);

    /// <summary>Adds a character style, based on the document's default character style.</summary>
    /// <inheritdoc cref="AddParagraphStyle(string, string?)"/>
    public WordStyle AddCharacterStyle(string styleId, string? name = null) => Add(StyleValues.Character, styleId, name);

    /// <summary>Adds a table style, based on the document's default table style.</summary>
    /// <inheritdoc cref="AddParagraphStyle(string, string?)"/>
    public WordStyle AddTableStyle(string styleId, string? name = null) => Add(StyleValues.Table, styleId, name);

    /// <summary>
    /// Returns the document's definition of a built-in paragraph style, adding OfficeIMO's definition first when
    /// the document has none. A definition the document already has is never replaced.
    /// </summary>
    public WordStyle GetOrAdd(WordParagraphStyles style) {
        if (style == WordParagraphStyles.Custom) throw new ArgumentException("Custom is a placeholder, not a style.", nameof(style));
        return GetOrAddLibraryStyle(style.ToStringStyle());
    }

    /// <summary>
    /// Returns the document's definition of a built-in table style, adding OfficeIMO's definition first when
    /// the document has none. A definition the document already has is never replaced.
    /// </summary>
    public WordStyle GetOrAdd(WordTableStyle style) => GetOrAddLibraryStyle(WordTableStyles.GetStyle(style).Val!.Value!);

    /// <summary>
    /// Removes a style. Content that used it falls back to the default style, styles based on it inherit from its
    /// base style instead, and next-style and linked-style references to it are removed.
    /// </summary>
    /// <returns><see langword="true"/> when the style existed and was removed.</returns>
    /// <exception cref="InvalidOperationException">The style is the default style of its type.</exception>
    public bool Remove(string styleId) {
        WordStyle? style = FindById(styleId);
        if (style == null) return false;
        style.EnsureEditable();
        if (style.IsDefault) throw new InvalidOperationException($"'{styleId}' is the default {style.Type.ToString().ToLowerInvariant()} style and cannot be removed.");

        string? parent = style.BasedOnStyleId;
        foreach (WordStyle other in this) {
            if (ReferenceEquals(other.Element, style.Element)) continue;
            if (string.Equals(other.BasedOnStyleId, styleId, StringComparison.Ordinal)) other.Element.BasedOn = parent == null ? null : new BasedOn { Val = parent };
            if (string.Equals(other.NextStyleId, styleId, StringComparison.Ordinal)) other.Element.NextParagraphStyle = null;
            if (string.Equals(other.LinkedStyleId, styleId, StringComparison.Ordinal)) other.Element.LinkedStyle = null;
        }
        style.Element.Remove();

        foreach (OpenXmlElement root in EnumerateContentRoots()) {
            foreach (OpenXmlElement reference in root.Descendants().Where(IsStyleReference).ToList()) {
                if (string.Equals(ReadStyleReference(reference), styleId, StringComparison.Ordinal)) reference.Remove();
            }
        }
        return true;
    }

    /// <summary>Returns an enumerator over the document's styles in document order.</summary>
    public IEnumerator<WordStyle> GetEnumerator() => Definitions.Select(style => new WordStyle(_document, style)).GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    private IEnumerable<Style> Definitions =>
        _document._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles?.Elements<Style>()
        ?? Enumerable.Empty<Style>();

    internal WordStyle? FindById(string styleId) {
        Style? style = Definitions.FirstOrDefault(definition => string.Equals(definition.StyleId?.Value, styleId, StringComparison.Ordinal));
        return style == null ? null : new WordStyle(_document, style);
    }

    internal WordStyle? FindByName(string name) {
        string trimmed = name.Trim();
        Style? style = Definitions.FirstOrDefault(definition =>
            string.Equals(definition.StyleName?.Val?.Value?.Trim(), trimmed, StringComparison.OrdinalIgnoreCase));
        return style == null ? null : new WordStyle(_document, style);
    }

    private Styles EnsureStyles() {
        if (_document.FileOpenAccess == FileAccess.Read) throw new InvalidOperationException("A read-only document's styles cannot be changed.");
        MainDocumentPart main = _document._wordprocessingDocument.MainDocumentPart
            ?? throw new InvalidOperationException("The document has no main part.");
        StyleDefinitionsPart part = main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>();
        return part.Styles ??= new Styles();
    }

    private WordStyle Add(StyleValues type, string styleId, string? name) {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleId);
        string displayName = string.IsNullOrWhiteSpace(name) ? styleId : name!;
        Styles styles = EnsureStyles();
        if (Contains(styleId)) throw new ArgumentException($"The document already has a style with the identifier '{styleId}'.", nameof(styleId));
        if (FindByName(displayName) != null) throw new ArgumentException($"The document already has a style named '{displayName}'.", nameof(name));

        var style = new Style { Type = type, StyleId = styleId, CustomStyle = true };
        style.StyleName = new StyleName { Val = displayName };
        // Word bases new styles on the default style of their type.
        string? baseId = Definitions.FirstOrDefault(definition => definition.Type?.Value == type && definition.Default?.Value == true)?.StyleId?.Value;
        if (baseId != null) style.BasedOn = new BasedOn { Val = baseId };
        if (type == StyleValues.Paragraph) style.NextParagraphStyle = new NextParagraphStyle { Val = styleId };
        if (type != StyleValues.Table) style.PrimaryStyle = new PrimaryStyle();
        styles.Append(style);
        return new WordStyle(_document, style);
    }

    private WordStyle GetOrAddLibraryStyle(string styleId) {
        EnsureStyles();
        _document.EnsureStyleDefinition(styleId);
        return FindById(styleId) ?? throw new InvalidOperationException($"OfficeIMO has no definition of '{styleId}'.");
    }

    internal IEnumerable<OpenXmlElement> EnumerateContentRoots() {
        MainDocumentPart? main = _document._wordprocessingDocument?.MainDocumentPart;
        if (main == null) yield break;
        if (main.Document != null) yield return main.Document;
        foreach (HeaderPart part in main.HeaderParts) if (part.Header != null) yield return part.Header;
        foreach (FooterPart part in main.FooterParts) if (part.Footer != null) yield return part.Footer;
        if (main.FootnotesPart?.Footnotes is { } footnotes) yield return footnotes;
        if (main.EndnotesPart?.Endnotes is { } endnotes) yield return endnotes;
        if (main.WordprocessingCommentsPart?.Comments is { } comments) yield return comments;
        if (main.NumberingDefinitionsPart?.Numbering is { } numbering) yield return numbering;
    }

    internal static bool IsStyleReference(OpenXmlElement element) => element is ParagraphStyleId or RunStyle or TableStyle;

    internal static string? ReadStyleReference(OpenXmlElement element) => element switch {
        ParagraphStyleId paragraph => paragraph.Val?.Value,
        RunStyle run => run.Val?.Value,
        TableStyle table => table.Val?.Value,
        _ => null
    };
}
