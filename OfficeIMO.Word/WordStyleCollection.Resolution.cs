namespace OfficeIMO.Word;

public sealed partial class WordStyleCollection {
    /// <summary>Finds a paragraph style the document defines, by identifier or by name.</summary>
    /// <returns>The style identifier, or <see langword="null"/> when the document defines no such paragraph style.</returns>
    internal string? FindParagraphStyleId(string styleIdOrName) {
        if (string.IsNullOrWhiteSpace(styleIdOrName)) return null;
        WordStyle? style = Get(styleIdOrName);
        return style?.Type == WordStyleType.Paragraph ? style.StyleId : null;
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
        if (WordBuiltInStyleNames.IsBuiltInId(styleIdOrName) && _document.FindBuiltInStyleId(styleIdOrName) is { } localized) return localized;
        foreach (WordParagraphStyles style in global::OfficeIMO.Internal.EnumCompat.GetValues<WordParagraphStyles>()) {
            if (style != WordParagraphStyles.Custom && string.Equals(style.ToStringStyle(), styleIdOrName, StringComparison.Ordinal)) {
                return styleIdOrName;
            }
        }
        return null;
    }
}
