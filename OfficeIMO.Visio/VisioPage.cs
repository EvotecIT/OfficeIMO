using System.Collections.Generic;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a single page within a Visio document.
    /// </summary>
    public class VisioPage {
        private readonly List<VisioShape> _shapes = new();
        private readonly List<VisioConnector> _connectors = new();
        private double _width = 8.26771653543307; // A4 width in inches
        private double _height = 11.69291338582677; // A4 height in inches
        private bool _gridVisible;
        private bool _snap = true;
        private VisioMeasurementUnit _defaultUnit = VisioMeasurementUnit.Inches;

        /// <summary>
        /// Initializes a new instance of the <see cref="VisioPage"/> class with default A4 size.
        /// </summary>
        /// <param name="name">Name of the page.</param>
        public VisioPage(string name) : this(name, 8.26771653543307, 11.69291338582677) {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="VisioPage"/> class.
        /// </summary>
        /// <param name="name">Name of the page.</param>
        /// <param name="widthInches">Page width in inches.</param>
        /// <param name="heightInches">Page height in inches.</param>
        public VisioPage(string name, double widthInches, double heightInches) {
            Name = name;
            NameU = name;
            _width = widthInches;
            _height = heightInches;
            ViewScale = -1;
            ViewCenterX = widthInches / 2;
            ViewCenterY = heightInches / 2;
        }

        /// <summary>
        /// Gets the identifier of the page within the document.
        /// </summary>
        public int Id { get; internal set; }

        /// <summary>
        /// Gets the page name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets or sets the universal name of the page.
        /// </summary>
        public string? NameU { get; set; }

        /// <summary>
        /// Gets or sets the view scale of the page.
        /// </summary>
        public double ViewScale { get; set; }

        /// <summary>
        /// Gets or sets the horizontal center of the view.
        /// </summary>
        public double ViewCenterX { get; set; }

        /// <summary>
        /// Gets or sets the vertical center of the view.
        /// </summary>
        public double ViewCenterY { get; set; }

        /// <summary>
        /// Gets or sets the page width in inches.
        /// </summary>
        public double Width {
            get => _width;
            set {
                _width = value;
                ViewCenterX = value / 2;
            }
        }

        /// <summary>
        /// Gets or sets the page width in centimeters.
        /// </summary>
        public double WidthCentimeters {
            get => _width.FromInches(VisioMeasurementUnit.Centimeters);
            set => Width = value.ToInches(VisioMeasurementUnit.Centimeters);
        }

        /// <summary>
        /// Gets or sets the page height in inches.
        /// </summary>
        public double Height {
            get => _height;
            set {
                _height = value;
                ViewCenterY = value / 2;
            }
        }

        /// <summary>
        /// Gets or sets the page height in centimeters.
        /// </summary>
        public double HeightCentimeters {
            get => _height.FromInches(VisioMeasurementUnit.Centimeters);
            set => Height = value.ToInches(VisioMeasurementUnit.Centimeters);
        }

        /// <summary>
        /// Default measurement unit for positions and sizes on this page.
        /// New shape-adding overloads use this unit implicitly.
        /// </summary>
        public VisioMeasurementUnit DefaultUnit {
            get => _defaultUnit;
            set => _defaultUnit = value;
        }

        /// <summary>
        /// Gets or sets the page width. Use <see cref="Width"/> instead.
        /// </summary>
        [System.Obsolete("Use Width instead")]
        public double PageWidth {
            get => Width;
            set => Width = value;
        }

        /// <summary>
        /// Gets or sets the page height. Use <see cref="Height"/> instead.
        /// </summary>
        [System.Obsolete("Use Height instead")]
        public double PageHeight {
            get => Height;
            set => Height = value;
        }

        /// <summary>
        /// Gets or sets a value indicating whether the grid is visible.
        /// </summary>
        public bool GridVisible {
            get => _gridVisible;
            set => _gridVisible = value;
        }

        /// <summary>
        /// Gets or sets a value indicating whether snapping to grid is enabled.
        /// </summary>
        public bool Snap {
            get => _snap;
            set => _snap = value;
        }

        /// <summary>
        /// Shapes placed on the page.
        /// </summary>
        public IList<VisioShape> Shapes => _shapes;

        /// <summary>
        /// Connectors placed on the page.
        /// </summary>
        public IList<VisioConnector> Connectors => _connectors;

        private string NextId() {
            int n = _shapes.Count + _connectors.Count + 1;
            return n.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static void ApplyUnits(ref double x, ref double y, ref double w, ref double h, VisioMeasurementUnit unit) {
            x = x.ToInches(unit); y = y.ToInches(unit); w = w.ToInches(unit); h = h.ToInches(unit);
        }

        // Normal, typed API — consistent with OfficeIMO style
        public VisioShape AddRectangle(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Rectangle" };
            _shapes.Add(s);
            return s;
        }

        // Overloads that respect page DefaultUnit to avoid manual conversions.
        public VisioShape AddRectangle(double x, double y, double width, double height, string? text = null) =>
            AddRectangle(x, y, width, height, text, DefaultUnit);

        public VisioShape AddSquare(double x, double y, double size, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            var s = AddRectangle(x, y, size, size, text, unit);
            s.NameU = "Square";
            return s;
        }

        public VisioShape AddSquare(double x, double y, double size, string? text = null) =>
            AddSquare(x, y, size, text, DefaultUnit);

        public VisioShape AddCircle(double x, double y, double diameter, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            // Avoid double-converting when passing the same variable for width/height
            double w = diameter, h = diameter;
            ApplyUnits(ref x, ref y, ref w, ref h, unit);
            var s = new VisioShape(NextId(), x, y, w, h, text ?? string.Empty) { NameU = "Circle" };
            _shapes.Add(s);
            return s;
        }

        public VisioShape AddCircle(double x, double y, double diameter, string? text = null) =>
            AddCircle(x, y, diameter, text, DefaultUnit);

        public VisioShape AddEllipse(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Ellipse" };
            _shapes.Add(s);
            return s;
        }

        public VisioShape AddEllipse(double x, double y, double width, double height, string? text = null) =>
            AddEllipse(x, y, width, height, text, DefaultUnit);

        public VisioShape AddDiamond(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Diamond" };
            _shapes.Add(s);
            return s;
        }

        public VisioShape AddDiamond(double x, double y, double width, double height, string? text = null) =>
            AddDiamond(x, y, width, height, text, DefaultUnit);

        public VisioShape AddTriangle(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Triangle" };
            _shapes.Add(s);
            return s;
        }

        public VisioShape AddTriangle(double x, double y, double width, double height, string? text = null) =>
            AddTriangle(x, y, width, height, text, DefaultUnit);

        private static int SideToIndex(VisioSide side) => side switch {
            VisioSide.Left => 0,
            VisioSide.Right => 1,
            VisioSide.Bottom => 2,
            VisioSide.Top => 3,
            _ => -1
        };

        public VisioConnector AddConnector(VisioShape from, VisioShape to, ConnectorKind kind = ConnectorKind.Straight, VisioSide fromSide = VisioSide.Auto, VisioSide toSide = VisioSide.Auto) {
            var conn = new VisioConnector(NextId(), from, to) { Kind = kind };
            // ensure side CPs when sides requested
            if (fromSide != VisioSide.Auto) from.EnsureSideConnectionPoints();
            if (toSide   != VisioSide.Auto) to.EnsureSideConnectionPoints();
            int fi = SideToIndex(fromSide), ti = SideToIndex(toSide);
            if (fi >= 0) conn.FromConnectionPoint = from.ConnectionPoints[fi];
            if (ti >= 0) conn.ToConnectionPoint   = to.ConnectionPoints[ti];
            _connectors.Add(conn);
            return conn;
        }

        /// <summary>
        /// Sets the page size.
        /// </summary>
        /// <param name="w">Width of the page.</param>
        /// <param name="h">Height of the page.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The current page.</returns>
        public VisioPage Size(double w, double h, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            Width = w.ToInches(unit);
            Height = h.ToInches(unit);
            return this;
        }

        /// <summary>
        /// Configures grid visibility and snapping.
        /// </summary>
        /// <param name="visible">Whether the grid is visible.</param>
        /// <param name="snap">Whether snapping is enabled.</param>
        /// <returns>The current page.</returns>
        public VisioPage Grid(bool visible, bool snap) {
            GridVisible = visible;
            Snap = snap;
            return this;
        }

        /// <summary>
        /// Adds a shape to the page.
        /// </summary>
        /// <param name="id">Identifier of the shape.</param>
        /// <param name="master">Master associated with the shape.</param>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="w">Width of the shape.</param>
        /// <param name="h">Height of the shape.</param>
        /// <param name="text">Optional text.</param>
        /// <returns>The created shape.</returns>
        public VisioShape AddShape(string id, VisioMaster master, double x, double y, double w, double h, string? text = null) {
            VisioShape shape = new VisioShape(id, x, y, w, h, text ?? string.Empty) { Master = master };
            _shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// Adds a connector between two shapes.
        /// </summary>
        /// <param name="id">Identifier of the connector.</param>
        /// <param name="from">Shape from which the connector starts.</param>
        /// <param name="to">Shape to which the connector ends.</param>
        /// <param name="kind">Type of connector.</param>
        /// <returns>The created connector.</returns>
        public VisioConnector AddConnector(string id, VisioShape from, VisioShape to, ConnectorKind kind) {
            VisioConnector connector = new VisioConnector(id, from, to) { Kind = kind };
            _connectors.Add(connector);
            return connector;
        }
    }
}

