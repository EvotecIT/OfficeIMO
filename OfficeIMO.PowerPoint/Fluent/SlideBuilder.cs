namespace OfficeIMO.PowerPoint.Fluent {
    /// <summary>
    /// Builder for slide content.
    /// </summary>
    public class SlideBuilder {
        private readonly PowerPointSlide _slide;

        internal SlideBuilder(PowerPointSlide slide) {
            _slide = slide;
        }

        /// <summary>
        /// Adds a title textbox to the slide.
        /// </summary>
        public SlideBuilder Title(string text) {
            _slide.AddTextBox(text);
            return this;
        }

        /// <summary>
        /// Adds a textbox with the specified text.
        /// </summary>
        public SlideBuilder Text(string text) {
            _slide.AddTextBox(text);
            return this;
        }

        /// <summary>
        /// Adds a bulleted list to the slide.
        /// </summary>
        public SlideBuilder Bullets(params string[] bullets) {
            PowerPointTextBox box = _slide.AddTextBox(string.Empty);
            foreach (string bullet in bullets) {
                box.AddBullet(bullet);
            }
            return this;
        }

        /// <summary>
        /// Adds an image from the given file path.
        /// </summary>
        public SlideBuilder Image(string imagePath) {
            _slide.AddPicture(imagePath);
            return this;
        }

        /// <summary>
        /// Adds a table to the slide.
        /// </summary>
        public SlideBuilder Table(int rows, int columns) {
            _slide.AddTable(rows, columns);
            return this;
        }

        /// <summary>
        /// Sets notes text for the slide.
        /// </summary>
        public SlideBuilder Notes(string text) {
            _slide.Notes.Text = text;
            return this;
        }
    }
}