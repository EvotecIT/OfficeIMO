using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for runs.
    /// </summary>
    public class RunBuilder {
        private readonly WordParagraph _run;

        internal RunBuilder(WordParagraph run) {
            _run = run;
        }

        public WordParagraph Run => _run;
    }
}
