using System.Collections.Generic;
using SixLabors.ImageSharp;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a shape on a Visio page.
    /// </summary>
    public class VisioShape {
        /// <summary>
        /// Initializes a new instance of the <see cref="VisioShape"/> class with the specified identifier.
        /// </summary>
        /// <param name="id">Identifier of the shape.</param>
        public VisioShape(string id) {
            Id = id;
            LineWeight = 0.0138889;
            Angle = 0;
            LineColor = Color.Black;
            FillColor = Color.White;
            LinePattern = 1; // Solid line
            FillPattern = 1; // Solid fill
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="VisioShape"/> class.
        /// </summary>
        /// <param name="id">Identifier of the shape.</param>
        /// <param name="pinX">X coordinate of the pin.</param>
        /// <param name="pinY">Y coordinate of the pin.</param>
        /// <param name="width">Width of the shape.</param>
        /// <param name="height">Height of the shape.</param>
        /// <param name="text">Text contained within the shape.</param>
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

        /// <summary>
        /// Gets or sets the shape name.
        /// </summary>
        public string? Name { get; set; }

        /// <summary>
        /// Gets or sets the universal name of the shape.
        /// </summary>
        public string? NameU { get; set; }

        /// <summary>
        /// Gets or sets the Visio type of the shape (for example "Group").
        /// </summary>
        public string? Type { get; internal set; }

        /// <summary>
        /// Gets or sets the master associated with the shape.
        /// </summary>
        public VisioMaster? Master { get; set; }

        /// <summary>
        /// Gets the identifier of the referenced master shape when <see cref="Master"/> is defined.
        /// </summary>
        public string? MasterShapeId { get; internal set; }

        /// <summary>
        /// Gets the master shape instance referenced by <see cref="MasterShapeId"/>, if any.
        /// </summary>
        public VisioShape? MasterShape { get; internal set; }

        /// <summary>
        /// Gets the universal name of the master.
        /// </summary>
        public string? MasterNameU => Master?.NameU;

        /// <summary>
        /// Gets or sets the X coordinate of the pin.
        /// </summary>
        public double PinX { get; set; }

        /// <summary>
        /// Gets or sets the Y coordinate of the pin.
        /// </summary>
        public double PinY { get; set; }

        /// <summary>
        /// Gets or sets the width of the shape.
        /// </summary>
        public double Width { get; set; }

        /// <summary>
        /// Gets or sets the height of the shape.
        /// </summary>
        public double Height { get; set; }

        /// <summary>
        /// Gets or sets the line weight of the shape.
        /// </summary>
        public double LineWeight { get; set; }

        /// <summary>
        /// Gets or sets the X coordinate of the local pin.
        /// </summary>
        public double LocPinX { get; set; }

        /// <summary>
        /// Gets or sets the Y coordinate of the local pin.
        /// </summary>
        public double LocPinY { get; set; }

        /// <summary>
        /// Gets or sets the rotation angle of the shape in radians.
        /// </summary>
        public double Angle { get; set; }

        /// <summary>
        /// Gets or sets the text contained in the shape.
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
        /// Parent shape when part of a group hierarchy.
        /// </summary>
        public VisioShape? Parent { get; internal set; }

        /// <summary>
        /// Child shapes when this shape represents a group.
        /// </summary>
        public IList<VisioShape> Children { get; } = new List<VisioShape>();

        /// <summary>
        /// Connection points associated with the shape.
        /// </summary>
        public IList<VisioConnectionPoint> ConnectionPoints { get; } = new List<VisioConnectionPoint>();

        /// <summary>
        /// Arbitrary data associated with the shape.
        /// </summary>
        public Dictionary<string, string> Data { get; } = new();

        /// <summary>
        /// Recursively searches the shape hierarchy for a shape with the provided identifier.
        /// </summary>
        /// <param name="id">Identifier to locate.</param>
        /// <returns>The matching shape when found; otherwise <c>null</c>.</returns>
        public VisioShape? FindDescendantById(string id) {
            if (Id == id) {
                return this;
            }

            foreach (VisioShape child in Children) {
                VisioShape? result = child.FindDescendantById(id);
                if (result != null) {
                    return result;
                }
            }

            return null;
        }

        /// <summary>
        /// Ensures the shape has four side connection points (Left, Right, Bottom, Top).
        /// If they already exist (>=4), nothing is added.
        /// Internal: users should not call this; side points are added automatically
        /// by the connector API when explicit sides are requested.
        /// </summary>
        internal void EnsureSideConnectionPoints() {
            if (ConnectionPoints.Count >= 4) return;
            ConnectionPoints.Add(new VisioConnectionPoint(0,       Height / 2,  1, 0));   // Left
            ConnectionPoints.Add(new VisioConnectionPoint(Width,   Height / 2, -1, 0));   // Right
            ConnectionPoints.Add(new VisioConnectionPoint(Width/2, 0,           0, 1));   // Bottom
            ConnectionPoints.Add(new VisioConnectionPoint(Width/2, Height,      0,-1));   // Top
        }

        /// <summary>
        /// Transforms a point from the shape's local coordinate system to the page coordinate system.
        /// </summary>
        /// <param name="x">X coordinate of the point relative to the shape's local coordinate system.</param>
        /// <param name="y">Y coordinate of the point relative to the shape's local coordinate system.</param>
        /// <returns>The point's absolute coordinates on the page.</returns>
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
