using SixLabors.ImageSharp;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a shape on a Visio page.
    /// </summary>
    public class VisioShape {
        public VisioShape(string id) {
            Id = id;
            LineWeight = 0.0138889;
            Angle = 0;
            LineColor = Color.Black;
            FillColor = Color.White;
            LinePattern = 1; // Solid line
            FillPattern = 1; // Solid fill
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

        /// <summary>
        /// Height of the shape.
        /// </summary>
        public double Height { get; set; }

        /// <summary>
        /// Thickness of the shape's outline.
        /// </summary>
        public double LineWeight { get; set; }

        /// <summary>
        /// Local X coordinate of the pin (anchor) point.
        /// </summary>
        public double LocPinX { get; set; }

        /// <summary>
        /// Local Y coordinate of the pin (anchor) point.
        /// </summary>
        public double LocPinY { get; set; }

        /// <summary>
        /// Rotation angle in radians.
        /// </summary>
        public double Angle { get; set; }

        /// <summary>
        /// Text displayed by the shape.
        /// </summary>
        public string? Text { get; set; }
        
        /// <summary>
        /// Line (border) color of the shape.
        /// </summary>
        public Color LineColor { get; set; }
        
        /// <summary>
        /// Fill color of the shape.
        /// </summary>
        public Color FillColor { get; set; }
        
        /// <summary>
        /// Line pattern (0=None, 1=Solid, 2=Dashed, etc.).
        /// </summary>
        public int LinePattern { get; set; }
        
        /// <summary>
        /// Fill pattern (0=None, 1=Solid, etc.).
        /// </summary>
        public int FillPattern { get; set; }

        /// <summary>
        /// Connection points associated with the shape.
        /// </summary>
        public IList<VisioConnectionPoint> ConnectionPoints { get; } = new List<VisioConnectionPoint>();

        /// <summary>
        /// Custom data stored on the shape.
        /// </summary>
        public Dictionary<string, string> Data { get; } = new();

        /// <summary>
        /// Transforms a point from the shape's local coordinate system to the page coordinate system.
        /// </summary>
        public (double X, double Y) GetAbsolutePoint(double x, double y) {
            double cos = Math.Cos(Angle);
            double sin = Math.Sin(Angle);
            double dx = x - LocPinX;
            double dy = y - LocPinY;
            double absX = PinX + cos * dx - sin * dy;
            double absY = PinY + sin * dx + cos * dy;
            return (absX, absY);
        }

        /// <summary>
        /// Computes the absolute bounds of the shape on the page.
        /// </summary>
        public (double Left, double Bottom, double Right, double Top) GetBounds() {
            (double x1, double y1) = GetAbsolutePoint(0, 0);
            (double x2, double y2) = GetAbsolutePoint(Width, 0);
            (double x3, double y3) = GetAbsolutePoint(0, Height);
            (double x4, double y4) = GetAbsolutePoint(Width, Height);
            double left = Math.Min(Math.Min(x1, x2), Math.Min(x3, x4));
            double right = Math.Max(Math.Max(x1, x2), Math.Max(x3, x4));
            double bottom = Math.Min(Math.Min(y1, y2), Math.Min(y3, y4));
            double top = Math.Max(Math.Max(y1, y2), Math.Max(y3, y4));
            return (left, bottom, right, top);
        }
    }
}
