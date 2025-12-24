using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint.Fluent {
    /// <summary>
    ///     Builder for slide content.
    /// </summary>
    public class PowerPointSlideBuilder {
        private readonly PowerPointFluentPresentation _presentation;
        private readonly PowerPointSlide _slide;

        internal PowerPointSlideBuilder(PowerPointFluentPresentation presentation, PowerPointSlide slide) {
            _presentation = presentation;
            _slide = slide;
        }

        /// <summary>
        ///     Adds a title textbox to the slide.
        /// </summary>
        public PowerPointSlideBuilder Title(string text, Action<PowerPointTextBox>? configure = null) {
            PowerPointTextBox box = _slide.AddTitle(text);
            configure?.Invoke(box);
            return this;
        }

        /// <summary>
        ///     Adds a textbox with the specified text.
        /// </summary>
        public PowerPointSlideBuilder TextBox(string text, Action<PowerPointTextBox>? configure = null) {
            PowerPointTextBox box = _slide.AddTextBox(text);
            configure?.Invoke(box);
            return this;
        }

        /// <summary>
        ///     Sets the slide layout.
        /// </summary>
        public PowerPointSlideBuilder Layout(int masterIndex, int layoutIndex) {
            _slide.SetLayout(masterIndex, layoutIndex);
            return this;
        }

        /// <summary>
        ///     Adds a bulleted list to the slide.
        /// </summary>
        public PowerPointSlideBuilder Bullets(params string[] bullets) {
            PowerPointTextBox box = _slide.AddTextBox(string.Empty);
            foreach (string bullet in bullets) {
                box.AddBullet(bullet);
            }

            return this;
        }

        /// <summary>
        ///     Adds a bulleted list to the slide and applies configuration.
        /// </summary>
        public PowerPointSlideBuilder Bullets(Action<PowerPointTextBox> configure, params string[] bullets) {
            PowerPointTextBox box = _slide.AddTextBox(string.Empty);
            foreach (string bullet in bullets) {
                box.AddBullet(bullet);
            }
            configure?.Invoke(box);
            return this;
        }

        /// <summary>
        ///     Adds a numbered list to the slide.
        /// </summary>
        public PowerPointSlideBuilder Numbered(params string[] items) {
            PowerPointTextBox box = _slide.AddTextBox(string.Empty);
            foreach (string item in items) {
                box.AddNumberedItem(item);
            }
            return this;
        }

        /// <summary>
        ///     Adds a numbered list to the slide and applies configuration.
        /// </summary>
        public PowerPointSlideBuilder Numbered(Action<PowerPointTextBox> configure, params string[] items) {
            PowerPointTextBox box = _slide.AddTextBox(string.Empty);
            foreach (string item in items) {
                box.AddNumberedItem(item);
            }
            configure?.Invoke(box);
            return this;
        }

        /// <summary>
        ///     Adds an image from the given file path.
        /// </summary>
        public PowerPointSlideBuilder Image(string imagePath, Action<PowerPointPicture>? configure = null) {
            PowerPointPicture picture = _slide.AddPicture(imagePath);
            configure?.Invoke(picture);
            return this;
        }

        /// <summary>
        ///     Adds a table to the slide.
        /// </summary>
        public PowerPointSlideBuilder Table(int rows, int columns, Action<PowerPointTable>? configure = null) {
            PowerPointTable table = _slide.AddTable(rows, columns);
            configure?.Invoke(table);
            return this;
        }

        /// <summary>
        ///     Adds an auto shape to the slide.
        /// </summary>
        public PowerPointSlideBuilder Shape(A.ShapeTypeValues shapeType, long left, long top, long width, long height,
            Action<PowerPointAutoShape>? configure = null) {
            PowerPointAutoShape shape = _slide.AddShape(shapeType, left, top, width, height);
            configure?.Invoke(shape);
            return this;
        }

        /// <summary>
        ///     Sets notes text for the slide.
        /// </summary>
        public PowerPointSlideBuilder Notes(string text) {
            _slide.Notes.Text = text;
            return this;
        }

        /// <summary>
        ///     Ends slide configuration and returns to the presentation builder.
        /// </summary>
        public PowerPointFluentPresentation End() {
            return _presentation;
        }
    }
}
