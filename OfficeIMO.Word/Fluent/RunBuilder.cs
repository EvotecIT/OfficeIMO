namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for text runs.
    /// </summary>
    public class RunBuilder {
        private readonly WordFluentDocument _fluent;

        internal RunBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        public WordFluentDocument AddRun(string text, bool bold = false) {
            var paragraph = _fluent.Document.AddParagraph(text);
            if (bold) {
                paragraph.SetBold();
            }
            return _fluent;
        }
    }
}
