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
        ///     Adds a title textbox to the slide at a specific position (EMU units).
        /// </summary>
        public PowerPointSlideBuilder Title(string text, long left, long top, long width, long height,
            Action<PowerPointTextBox>? configure = null) {
            PowerPointTextBox box = _slide.AddTitle(text, left, top, width, height);
            configure?.Invoke(box);
            return this;
        }

        /// <summary>
        ///     Adds a title textbox to the slide at a specific position (centimeters).
        /// </summary>
        public PowerPointSlideBuilder TitleCm(string text, double leftCm, double topCm, double widthCm,
            double heightCm, Action<PowerPointTextBox>? configure = null) {
            PowerPointTextBox box = _slide.AddTitleCm(text, leftCm, topCm, widthCm, heightCm);
            configure?.Invoke(box);
            return this;
        }

        /// <summary>
        ///     Adds a title textbox to the slide at a specific position (inches).
        /// </summary>
        public PowerPointSlideBuilder TitleInches(string text, double leftInches, double topInches,
            double widthInches, double heightInches, Action<PowerPointTextBox>? configure = null) {
            PowerPointTextBox box = _slide.AddTitleInches(text, leftInches, topInches, widthInches, heightInches);
            configure?.Invoke(box);
            return this;
        }

        /// <summary>
        ///     Adds a title textbox to the slide at a specific position (points).
        /// </summary>
        public PowerPointSlideBuilder TitlePoints(string text, double leftPoints, double topPoints, double widthPoints,
            double heightPoints, Action<PowerPointTextBox>? configure = null) {
            PowerPointTextBox box = _slide.AddTitlePoints(text, leftPoints, topPoints, widthPoints, heightPoints);
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
        ///     Adds a textbox with the specified text at a position (EMU units).
        /// </summary>
        public PowerPointSlideBuilder TextBox(string text, long left, long top, long width, long height,
            Action<PowerPointTextBox>? configure = null) {
            PowerPointTextBox box = _slide.AddTextBox(text, left, top, width, height);
            configure?.Invoke(box);
            return this;
        }

        /// <summary>
        ///     Adds a textbox with the specified text at a position (centimeters).
        /// </summary>
        public PowerPointSlideBuilder TextBoxCm(string text, double leftCm, double topCm, double widthCm,
            double heightCm, Action<PowerPointTextBox>? configure = null) {
            PowerPointTextBox box = _slide.AddTextBoxCm(text, leftCm, topCm, widthCm, heightCm);
            configure?.Invoke(box);
            return this;
        }

        /// <summary>
        ///     Adds a textbox with the specified text at a position (inches).
        /// </summary>
        public PowerPointSlideBuilder TextBoxInches(string text, double leftInches, double topInches,
            double widthInches, double heightInches, Action<PowerPointTextBox>? configure = null) {
            PowerPointTextBox box = _slide.AddTextBoxInches(text, leftInches, topInches, widthInches, heightInches);
            configure?.Invoke(box);
            return this;
        }

        /// <summary>
        ///     Adds a textbox with the specified text at a position (points).
        /// </summary>
        public PowerPointSlideBuilder TextBoxPoints(string text, double leftPoints, double topPoints,
            double widthPoints, double heightPoints, Action<PowerPointTextBox>? configure = null) {
            PowerPointTextBox box = _slide.AddTextBoxPoints(text, leftPoints, topPoints, widthPoints, heightPoints);
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
        ///     Adds an image from the given file path at a position (EMU units).
        /// </summary>
        public PowerPointSlideBuilder Image(string imagePath, long left, long top, long width, long height,
            Action<PowerPointPicture>? configure = null) {
            PowerPointPicture picture = _slide.AddPicture(imagePath, left, top, width, height);
            configure?.Invoke(picture);
            return this;
        }

        /// <summary>
        ///     Adds an image from the given file path at a position (centimeters).
        /// </summary>
        public PowerPointSlideBuilder ImageCm(string imagePath, double leftCm, double topCm, double widthCm,
            double heightCm, Action<PowerPointPicture>? configure = null) {
            PowerPointPicture picture = _slide.AddPictureCm(imagePath, leftCm, topCm, widthCm, heightCm);
            configure?.Invoke(picture);
            return this;
        }

        /// <summary>
        ///     Adds an image from the given file path at a position (inches).
        /// </summary>
        public PowerPointSlideBuilder ImageInches(string imagePath, double leftInches, double topInches,
            double widthInches, double heightInches, Action<PowerPointPicture>? configure = null) {
            PowerPointPicture picture = _slide.AddPictureInches(imagePath, leftInches, topInches, widthInches,
                heightInches);
            configure?.Invoke(picture);
            return this;
        }

        /// <summary>
        ///     Adds an image from the given file path at a position (points).
        /// </summary>
        public PowerPointSlideBuilder ImagePoints(string imagePath, double leftPoints, double topPoints,
            double widthPoints, double heightPoints, Action<PowerPointPicture>? configure = null) {
            PowerPointPicture picture = _slide.AddPicturePoints(imagePath, leftPoints, topPoints, widthPoints,
                heightPoints);
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
        ///     Adds a table to the slide at a position (EMU units).
        /// </summary>
        public PowerPointSlideBuilder Table(int rows, int columns, long left, long top, long width, long height,
            Action<PowerPointTable>? configure = null) {
            PowerPointTable table = _slide.AddTable(rows, columns, left, top, width, height);
            configure?.Invoke(table);
            return this;
        }

        /// <summary>
        ///     Adds a table to the slide at a position (centimeters).
        /// </summary>
        public PowerPointSlideBuilder TableCm(int rows, int columns, double leftCm, double topCm, double widthCm,
            double heightCm, Action<PowerPointTable>? configure = null) {
            PowerPointTable table = _slide.AddTableCm(rows, columns, leftCm, topCm, widthCm, heightCm);
            configure?.Invoke(table);
            return this;
        }

        /// <summary>
        ///     Adds a table to the slide at a position (inches).
        /// </summary>
        public PowerPointSlideBuilder TableInches(int rows, int columns, double leftInches, double topInches,
            double widthInches, double heightInches, Action<PowerPointTable>? configure = null) {
            PowerPointTable table = _slide.AddTableInches(rows, columns, leftInches, topInches, widthInches,
                heightInches);
            configure?.Invoke(table);
            return this;
        }

        /// <summary>
        ///     Adds a table to the slide at a position (points).
        /// </summary>
        public PowerPointSlideBuilder TablePoints(int rows, int columns, double leftPoints, double topPoints,
            double widthPoints, double heightPoints, Action<PowerPointTable>? configure = null) {
            PowerPointTable table = _slide.AddTablePoints(rows, columns, leftPoints, topPoints, widthPoints,
                heightPoints);
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
        ///     Adds an auto shape to the slide using centimeter measurements.
        /// </summary>
        public PowerPointSlideBuilder ShapeCm(A.ShapeTypeValues shapeType, double leftCm, double topCm, double widthCm,
            double heightCm, Action<PowerPointAutoShape>? configure = null) {
            PowerPointAutoShape shape = _slide.AddShapeCm(shapeType, leftCm, topCm, widthCm, heightCm);
            configure?.Invoke(shape);
            return this;
        }

        /// <summary>
        ///     Adds an auto shape to the slide using inch measurements.
        /// </summary>
        public PowerPointSlideBuilder ShapeInches(A.ShapeTypeValues shapeType, double leftInches, double topInches,
            double widthInches, double heightInches, Action<PowerPointAutoShape>? configure = null) {
            PowerPointAutoShape shape = _slide.AddShapeInches(shapeType, leftInches, topInches, widthInches,
                heightInches);
            configure?.Invoke(shape);
            return this;
        }

        /// <summary>
        ///     Adds an auto shape to the slide using point measurements.
        /// </summary>
        public PowerPointSlideBuilder ShapePoints(A.ShapeTypeValues shapeType, double leftPoints, double topPoints,
            double widthPoints, double heightPoints, Action<PowerPointAutoShape>? configure = null) {
            PowerPointAutoShape shape = _slide.AddShapePoints(shapeType, leftPoints, topPoints, widthPoints,
                heightPoints);
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
