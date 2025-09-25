namespace OfficeIMO.Word {
    /// <summary>
    /// Specifies reference types for cross references.
    /// </summary>
    public enum WordCrossReferenceType {
        /// <summary>Reference to a bookmark.</summary>
        Bookmark,
        /// <summary>Reference to a heading bookmark.</summary>
        Heading
    }

    /// <summary>
    /// Provides helper methods for inserting cross reference fields.
    /// </summary>
    public class WordCrossReference : WordElement {
        /// <summary>
        /// Inserts a REF field referencing the specified bookmark or heading.
        /// </summary>
        /// <param name="paragraph">Paragraph to insert the field into.</param>
        /// <param name="referenceId">Bookmark or heading identifier.</param>
        /// <param name="type">Type of reference.</param>
        /// <returns>The paragraph that this method was called on.</returns>
        public static WordParagraph AddCrossReference(WordParagraph paragraph, string referenceId, WordCrossReferenceType type) {
            if (paragraph == null) throw new ArgumentNullException(nameof(paragraph));
            if (string.IsNullOrWhiteSpace(referenceId)) throw new ArgumentNullException(nameof(referenceId));

            var builder = new WordFieldBuilder(WordFieldType.Ref)
                .AddInstruction(referenceId)
                .AddSwitch("\\h");
            WordField.AddField(paragraph, builder, false);
            return paragraph;
        }
    }
}
