using System;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for paragraphs.
    /// </summary>
    public class ParagraphBuilder {
        private readonly WordFluentDocument _fluent;
        private readonly WordParagraph _paragraph;

        internal ParagraphBuilder(WordFluentDocument fluent, WordParagraph paragraph) {
            _fluent = fluent;
            _paragraph = paragraph;
        }

        public WordParagraph Paragraph => _paragraph;

        public ParagraphBuilder AddParagraph(string text = "") {
            var paragraph = _fluent.Document.AddParagraph(text);
            return new ParagraphBuilder(_fluent, paragraph);
        }

        public ParagraphBuilder Text(string text, Action<TextBuilder>? configure = null) {
            var run = _paragraph.AddText(text);
            configure?.Invoke(new TextBuilder(run));
            return this;
        }

        public ParagraphBuilder Run(string text, Action<TextBuilder>? configure = null) => Text(text, configure);
    }
}
