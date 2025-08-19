namespace OfficeIMO.PowerPoint.Fluent {
    /// <summary>
    /// Builder for slide content.
    /// </summary>
    public class PowerPointSlideBuilder {
        private readonly PowerPointSlide _slide;

        internal PowerPointSlideBuilder(PowerPointSlide slide) {
            _slide = slide;
        }

        /// <summary>
        /// Adds a title textbox to the slide.
        /// </summary>
        public PowerPointSlideBuilder Title(string text) {
            _slide.AddTextBox(text);
            return this;
        }

        /// <summary>
        /// Adds a textbox with the specified text.
        /// </summary>
        public PowerPointSlideBuilder Text(string text) {
            _slide.AddTextBox(text);
            return this;
        }

        /// <summary>
        /// Adds a bulleted list to the slide.
        /// </summary>
        public PowerPointSlideBuilder Bullets(params string[] bullets) {
            PowerPointTextBox box = _slide.AddTextBox(string.Empty);
            foreach (string bullet in bullets) {
                box.AddBullet(bullet);
            }
            return this;
        }

        /// <summary>
        /// Adds an image from the given file path.
        /// </summary>
        public PowerPointSlideBuilder Image(string imagePath) {
            _slide.AddPicture(imagePath);
            return this;
        }

        /// <summary>
        /// Adds a table to the slide.
        /// </summary>
        public PowerPointSlideBuilder Table(int rows, int columns) {
            _slide.AddTable(rows, columns);
            return this;
        }

        /// <summary>
        /// Sets notes text for the slide.
        /// </summary>
        public PowerPointSlideBuilder Notes(string text) {
            _slide.Notes.Text = text;
            return this;
        }
    }
}