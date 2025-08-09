namespace OfficeIMO.Word {
    /// <summary>
    /// Provides helper methods for mapping between heading levels and paragraph styles.
    /// </summary>
    public static class HeadingStyleMapper {
        /// <summary>
        /// Returns the corresponding paragraph style for the specified heading level.
        /// </summary>
        /// <param name="level">Heading level starting at 1.</param>
        /// <returns>The matching <see cref="WordParagraphStyles"/> value.</returns>
        public static WordParagraphStyles GetHeadingStyleForLevel(int level) {
            return level switch {
                1 => WordParagraphStyles.Heading1,
                2 => WordParagraphStyles.Heading2,
                3 => WordParagraphStyles.Heading3,
                4 => WordParagraphStyles.Heading4,
                5 => WordParagraphStyles.Heading5,
                6 => WordParagraphStyles.Heading6,
                7 => WordParagraphStyles.Heading7,
                8 => WordParagraphStyles.Heading8,
                9 => WordParagraphStyles.Heading9,
                _ => WordParagraphStyles.Heading1,
            };
        }

        /// <summary>
        /// Resolves the heading level for the provided paragraph style.
        /// </summary>
        /// <param name="style">Paragraph style to evaluate.</param>
        /// <returns>The heading level or 0 when the style is not a heading.</returns>
        public static int GetLevelForHeadingStyle(WordParagraphStyles style) {
            return style switch {
                WordParagraphStyles.Heading1 => 1,
                WordParagraphStyles.Heading2 => 2,
                WordParagraphStyles.Heading3 => 3,
                WordParagraphStyles.Heading4 => 4,
                WordParagraphStyles.Heading5 => 5,
                WordParagraphStyles.Heading6 => 6,
                WordParagraphStyles.Heading7 => 7,
                WordParagraphStyles.Heading8 => 8,
                WordParagraphStyles.Heading9 => 9,
                _ => 0,
            };
        }
    }
}
