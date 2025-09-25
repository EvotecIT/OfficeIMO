namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for runs.
    /// </summary>
    public class RunBuilder {
        private readonly WordParagraph _run;

        internal RunBuilder(WordParagraph run) {
            _run = run;
        }

        /// <summary>
        /// Gets the underlying run.
        /// </summary>
        public WordParagraph Run => _run;
    }
}
