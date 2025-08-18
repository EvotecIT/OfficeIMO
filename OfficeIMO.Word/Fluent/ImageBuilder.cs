namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for images.
    /// </summary>
    public class ImageBuilder {
        private readonly WordFluentDocument _fluent;

        internal ImageBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        public WordFluentDocument AddImage(string url) {
            _fluent.Document.AddImageFromUrl(url);
            return _fluent;
        }
    }
}
