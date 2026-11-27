using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Predefined character styles available in Word documents.
    /// </summary>
    public enum WordCharacterStyles {
        /// <summary>
        /// Default character formatting used for paragraphs.
        /// </summary>
        DefaultParagraphFont,
        /// <summary>
        /// Character style associated with Heading 1 paragraphs.
        /// </summary>
        Heading1Char,
        /// <summary>
        /// Character style associated with Heading 2 paragraphs.
        /// </summary>
        Heading2Char,
        /// <summary>
        /// Character style associated with Heading 3 paragraphs.
        /// </summary>
        Heading3Char,
        /// <summary>
        /// Character style associated with Heading 4 paragraphs.
        /// </summary>
        Heading4Char,
        /// <summary>
        /// Character style associated with Heading 5 paragraphs.
        /// </summary>
        Heading5Char,
        /// <summary>
        /// Character style associated with Heading 6 paragraphs.
        /// </summary>
        Heading6Char,
        /// <summary>
        /// Character style associated with Heading 7 paragraphs.
        /// </summary>
        Heading7Char,
        /// <summary>
        /// Character style associated with Heading 8 paragraphs.
        /// </summary>
        Heading8Char,
        /// <summary>
        /// Character style associated with Heading 9 paragraphs.
        /// </summary>
        Heading9Char,
    }

    /// <summary>
    /// Helper methods for working with <see cref="WordCharacterStyles"/> values.
    /// </summary>
    internal static class WordCharacterStyle {
        /// <summary>
        /// Gets the <see cref="Style"/> definition for the specified character style.
        /// </summary>
        /// <param name="style">Character style to retrieve.</param>
        /// <returns>The Open XML <see cref="Style"/> definition.</returns>
        internal static Style GetStyleDefinition(WordCharacterStyles style) {
            switch (style) {
                case WordCharacterStyles.DefaultParagraphFont: return DefaultParagraphFont;
                case WordCharacterStyles.Heading1Char: return WordThemeDefaults.CreateHeadingCharacter(1, OfficeTheme.Default);
                case WordCharacterStyles.Heading2Char: return WordThemeDefaults.CreateHeadingCharacter(2, OfficeTheme.Default);
                case WordCharacterStyles.Heading3Char: return WordThemeDefaults.CreateHeadingCharacter(3, OfficeTheme.Default);
                case WordCharacterStyles.Heading4Char: return WordThemeDefaults.CreateHeadingCharacter(4, OfficeTheme.Default);
                case WordCharacterStyles.Heading5Char: return WordThemeDefaults.CreateHeadingCharacter(5, OfficeTheme.Default);
                case WordCharacterStyles.Heading6Char: return WordThemeDefaults.CreateHeadingCharacter(6, OfficeTheme.Default);
                case WordCharacterStyles.Heading7Char: return StyleHeading7;
                case WordCharacterStyles.Heading8Char: return StyleHeading8;
                case WordCharacterStyles.Heading9Char: return StyleHeading9;
            }

            throw new ArgumentOutOfRangeException(nameof(style));
        }

        /// <summary>
        /// Converts the style enumeration to its XML string representation.
        /// </summary>
        /// <param name="style">Style to convert.</param>
        /// <returns>String identifier used in the document.</returns>
        public static string ToStringStyle(this WordCharacterStyles style) {
            switch (style) {
                case WordCharacterStyles.DefaultParagraphFont: return "DefaultParagraphFont";
                case WordCharacterStyles.Heading1Char: return "Heading1Char";
                case WordCharacterStyles.Heading2Char: return "Heading2Char";
                case WordCharacterStyles.Heading3Char: return "Heading3Char";
                case WordCharacterStyles.Heading4Char: return "Heading4Char";
                case WordCharacterStyles.Heading5Char: return "Heading5Char";
                case WordCharacterStyles.Heading6Char: return "Heading6Char";
                case WordCharacterStyles.Heading7Char: return "Heading7Char";
                case WordCharacterStyles.Heading8Char: return "Heading8Char";
                case WordCharacterStyles.Heading9Char: return "Heading9Char";
            }

            throw new ArgumentOutOfRangeException(nameof(style));
        }

        /// <summary>
        /// Converts the XML string identifier to its corresponding <see cref="WordCharacterStyles"/> value.
        /// </summary>
        /// <param name="style">String identifier.</param>
        /// <returns>The matching enumeration value.</returns>
        public static WordCharacterStyles GetStyle(string style) {
            switch (style) {
                case "DefaultParagraphFont": return WordCharacterStyles.DefaultParagraphFont;
                case "Heading1Char": return WordCharacterStyles.Heading1Char;
                case "Heading2Char": return WordCharacterStyles.Heading2Char;
                case "Heading3Char": return WordCharacterStyles.Heading3Char;
                case "Heading4Char": return WordCharacterStyles.Heading4Char;
                case "Heading5Char": return WordCharacterStyles.Heading5Char;
                case "Heading6Char": return WordCharacterStyles.Heading6Char;
                case "Heading7Char": return WordCharacterStyles.Heading7Char;
                case "Heading8Char": return WordCharacterStyles.Heading8Char;
                case "Heading9Char": return WordCharacterStyles.Heading9Char;
            }
            throw new ArgumentOutOfRangeException(nameof(style));
        }

        private static Style DefaultParagraphFont => WordEmbeddedMarkup.CreateDefaultStyleElement("DefaultParagraphFont", xml => new Style(xml));

        private static Style StyleHeading7 => WordEmbeddedMarkup.CreateDefaultStyleElement("Heading7Char", xml => new Style(xml));

        private static Style StyleHeading8 => WordEmbeddedMarkup.CreateDefaultStyleElement("Heading8Char", xml => new Style(xml));

        private static Style StyleHeading9 => WordEmbeddedMarkup.CreateDefaultStyleElement("Heading9Char", xml => new Style(xml));
    }
}
