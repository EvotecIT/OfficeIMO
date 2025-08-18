using System;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Provides data for the <see cref="WordHtmlConverterExtensions.StyleMissing"/> event.
    /// </summary>
    public class StyleMissingEventArgs : EventArgs {
        /// <summary>
        /// Initializes a new instance of the <see cref="StyleMissingEventArgs"/> class.
        /// </summary>
        /// <param name="paragraph">Paragraph where the missing style was referenced.</param>
        /// <param name="className">Name of the CSS class that was not found.</param>
        public StyleMissingEventArgs(WordParagraph paragraph, string className) {
            Paragraph = paragraph ?? throw new ArgumentNullException(nameof(paragraph));
            ClassName = className ?? throw new ArgumentNullException(nameof(className));
        }

        /// <summary>Gets the paragraph where the missing style was referenced.</summary>
        public WordParagraph Paragraph { get; }

        /// <summary>Gets the name of the missing CSS class.</summary>
        public string ClassName { get; }

        /// <summary>Gets or sets the style to apply for the missing class.</summary>
        public WordParagraphStyles? Style { get; set; }

        /// <summary>Gets or sets the identifier of the applied style.</summary>
        public string? StyleId { get; set; }
    }
}
