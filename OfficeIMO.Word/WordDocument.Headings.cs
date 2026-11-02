namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Adds a heading paragraph using the document's built-in heading style.</summary>
    /// <param name="text">Heading text.</param>
    /// <param name="level">Heading level from one through nine.</param>
    /// <returns>The created heading paragraph.</returns>
    public WordParagraph AddHeading(string text, int level = 1) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        WordParagraphStyles style = level switch {
            1 => WordParagraphStyles.Heading1,
            2 => WordParagraphStyles.Heading2,
            3 => WordParagraphStyles.Heading3,
            4 => WordParagraphStyles.Heading4,
            5 => WordParagraphStyles.Heading5,
            6 => WordParagraphStyles.Heading6,
            7 => WordParagraphStyles.Heading7,
            8 => WordParagraphStyles.Heading8,
            9 => WordParagraphStyles.Heading9,
            _ => throw new ArgumentOutOfRangeException(nameof(level), "Heading levels range from one through nine.")
        };
        return AddParagraph(text).SetStyle(style);
    }
}
