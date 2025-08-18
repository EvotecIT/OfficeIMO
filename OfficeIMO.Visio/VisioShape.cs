namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a shape on a Visio page.
    /// </summary>
    public class VisioShape {
        public VisioShape(string id) {
            Id = id;
        }

        public VisioShape(string id, double pinX, double pinY, double width, double height, string text) : this(id) {
            PinX = pinX;
            PinY = pinY;
            Width = width;
            Height = height;
            Text = text;
        }

        /// <summary>
        /// Identifier of the shape.
        /// </summary>
        public string Id { get; }

        public string? NameU { get; set; }

        public double PinX { get; set; }

        public double PinY { get; set; }

        public double Width { get; set; }

        public double Height { get; set; }

        public string? Text { get; set; }
    }
}

