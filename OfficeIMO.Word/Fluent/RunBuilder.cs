using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for text runs.
    /// </summary>
    public class RunBuilder {
        private readonly WordFluentDocument _fluent;
        private readonly WordParagraph? _paragraph;

        internal RunBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        internal RunBuilder(WordFluentDocument fluent, WordParagraph paragraph) {
            _fluent = fluent;
            _paragraph = paragraph;
        }

        public WordParagraph? Paragraph => _paragraph;

        public RunBuilder AddRun(string text, bool bold = false) {
            var paragraph = _fluent.Document.AddParagraph(text);
            if (bold) {
                paragraph.SetBold();
            }
            return new RunBuilder(_fluent, paragraph);
        }
    }
}
