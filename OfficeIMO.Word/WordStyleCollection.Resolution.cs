namespace OfficeIMO.Word;

public sealed partial class WordStyleCollection {
    /// <summary>Finds a paragraph style the document defines, by identifier or by name.</summary>
    /// <returns>The style identifier, or <see langword="null"/> when the document defines no such paragraph style.</returns>
    internal string? FindParagraphStyleId(string styleIdOrName) {
        if (string.IsNullOrWhiteSpace(styleIdOrName)) return null;
        if (FindById(styleIdOrName) is { Type: WordStyleType.Paragraph } byId) return byId.StyleId;
        // The first referenceable paragraph style with the name, even when a style of another type shares it.
        return WordStyleCatalog.For(_document).FindReferenceableByName(styleIdOrName, definition =>
            string.IsNullOrEmpty(definition.Type) || string.Equals(definition.Type, "paragraph", StringComparison.Ordinal))?.StyleId;
    }

    /// <summary>
    /// Resolves a paragraph style requested for generated content: a paragraph style the document defines, by
    /// identifier or name, or an OfficeIMO built-in paragraph style, which is added to the document when it is used.
    /// </summary>
    /// <returns>The style identifier, or <see langword="null"/> when neither the document nor OfficeIMO defines it.</returns>
    internal string? ResolveParagraphStyleForContent(string styleIdOrName) {
        string? styleId = FindParagraphStyleId(styleIdOrName);
        if (styleId != null) return styleId;
        // A localized template stores built-in styles under translated identifiers.
        if (WordBuiltInStyleNames.IsBuiltInId(styleIdOrName) &&
            WordBuiltInStyleNames.GetStyleType(styleIdOrName) == "paragraph" &&
            _document.FindBuiltInStyleId(styleIdOrName) is { } localized) return localized;
        // The library definition is added only where no style of another type already holds the identifier.
        if (Get(styleIdOrName) != null) return null;
        foreach (WordParagraphStyles style in global::OfficeIMO.Internal.EnumCompat.GetValues<WordParagraphStyles>()) {
            if (style != WordParagraphStyles.Custom && string.Equals(style.ToStringStyle(), styleIdOrName, StringComparison.Ordinal)) {
                return styleIdOrName;
            }
        }
        return null;
    }
}
