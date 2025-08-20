namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a shape on a Visio page.
    /// </summary>
    public class VisioShape {
        public VisioShape(string id) {
            Id = id;
            LineWeight = 0.0138889;
            Angle = 0;
        }

        public VisioShape(string id, double pinX, double pinY, double width, double height, string text) : this(id) {
            PinX = pinX;
            PinY = pinY;
            Width = width;
            Height = height;
            LocPinX = width / 2;
            LocPinY = height / 2;
            Text = text;
        }

        /// <summary>
        /// Identifier of the shape.
        /// </summary>
        public string Id { get; }

        public string? Name { get; set; }

        public string? NameU { get; set; }

        public VisioMaster? Master { get; set; }

        public string? MasterNameU => Master?.NameU;

        public double PinX { get; set; }

        public double PinY { get; set; }

        public double Width { get; set; }

        public double Height { get; set; }

        public double LineWeight { get; set; }

        public double LocPinX { get; set; }

        public double LocPinY { get; set; }

        public double Angle { get; set; }

        public string? Text { get; set; }

        public IList<VisioConnectionPoint> ConnectionPoints { get; } = new List<VisioConnectionPoint>();
    }
}
