using OfficeIMO.Word;

namespace OfficeIMO.Word.Fluent {
    /// <summary>
    /// Builder for images.
    /// </summary>
    public class ImageBuilder {
        private readonly WordFluentDocument _fluent;
        private readonly WordImage? _image;

        internal ImageBuilder(WordFluentDocument fluent) {
            _fluent = fluent;
        }

        internal ImageBuilder(WordFluentDocument fluent, WordImage image) {
            _fluent = fluent;
            _image = image;
        }

        public WordImage? Image => _image;

        public ImageBuilder AddImage(string url) {
            var image = _fluent.Document.AddImageFromUrl(url);
            return new ImageBuilder(_fluent, image);
        }
    }
}
