using System;
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
        private VisioMeasurementUnit _scaleMeasurementUnit = VisioMeasurementUnit.Inches;
        private double _viewScale = 1;
        private VisioScaleSetting? _pageScaleOverride;
        private VisioScaleSetting? _drawingScaleOverride;

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
        public double ViewScale {
            get => _viewScale;
            set {
                if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0) {
                    _viewScale = 1;
                } else {
                    _viewScale = value;
                }
            }
        }

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
        /// Measurement unit used to compute page and drawing scales when explicit overrides are not supplied.
        /// Defaults to inches and typically mirrors <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioMeasurementUnit ScaleMeasurementUnit {
            get => _scaleMeasurementUnit;
            set => _scaleMeasurementUnit = value;
        }

        /// <summary>
        /// Gets or sets the page scale (the ratio between page units and real-world units).
        /// </summary>
        public VisioScaleSetting PageScale {
            get {
                VisioScaleSetting scale = _pageScaleOverride ?? VisioScaleSetting.FromUnit(ScaleMeasurementUnit);
                return scale.Normalized();
            }
            set => _pageScaleOverride = value.Normalized();
        }

        /// <summary>
        /// Gets or sets the drawing scale (the ratio between drawing units and real-world units).
        /// </summary>
        public VisioScaleSetting DrawingScale {
            get {
                VisioScaleSetting scale = _drawingScaleOverride ?? VisioScaleSetting.FromUnit(ScaleMeasurementUnit);
                return scale.Normalized();
            }
            set => _drawingScaleOverride = value.Normalized();
        }

        /// <summary>
        /// Removes any custom page scale override and reverts to <see cref="ScaleMeasurementUnit"/>.
        /// </summary>
        public void ResetPageScale() => _pageScaleOverride = null;

        /// <summary>
        /// Removes any custom drawing scale override and reverts to <see cref="ScaleMeasurementUnit"/>.
        /// </summary>
        public void ResetDrawingScale() => _drawingScaleOverride = null;

        internal VisioScaleSetting GetEffectivePageScale() {
            VisioScaleSetting scale = _pageScaleOverride ?? VisioScaleSetting.FromUnit(ScaleMeasurementUnit);
            return scale.Normalized();
        }

        internal VisioScaleSetting GetEffectiveDrawingScale() {
            VisioScaleSetting scale = _drawingScaleOverride ?? VisioScaleSetting.FromUnit(ScaleMeasurementUnit);
            return scale.Normalized();
        }

        internal void ApplyLoadedPageScale(VisioScaleSetting scale) {
            ScaleMeasurementUnit = scale.Unit;
            if (scale.IsDefault) {
                _pageScaleOverride = null;
            } else {
                _pageScaleOverride = scale.Normalized();
            }
        }

        internal void ApplyLoadedDrawingScale(VisioScaleSetting scale) {
            if (scale.IsDefault && scale.Unit == ScaleMeasurementUnit) {
                _drawingScaleOverride = null;
            } else {
                _drawingScaleOverride = scale.Normalized();
            }
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

        /// <summary>
        /// Adds a rectangle shape.
        /// </summary>
        /// <param name="x">X coordinate of the shape origin.</param>
        /// <param name="y">Y coordinate of the shape origin.</param>
        /// <param name="width">Width of the rectangle.</param>
        /// <param name="height">Height of the rectangle.</param>
        /// <param name="text">Optional text placed on the shape.</param>
        /// <param name="unit">Measurement unit for the provided values.</param>
        /// <returns>The created rectangle shape.</returns>
        public VisioShape AddRectangle(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Rectangle" };
            _shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a rectangle shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the rectangle.</param>
        /// <param name="height">Height of the rectangle.</param>
        /// <param name="text">Optional text.</param>
        /// <returns>The created rectangle shape.</returns>
        public VisioShape AddRectangle(double x, double y, double width, double height, string? text = null) =>
            AddRectangle(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a square shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="size">Width and height of the square.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created square shape.</returns>
        public VisioShape AddSquare(double x, double y, double size, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            var s = AddRectangle(x, y, size, size, text, unit);
            s.NameU = "Square";
            return s;
        }

        /// <summary>
        /// Adds a square using the page <see cref="DefaultUnit"/>.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="size">Width and height of the square.</param>
        /// <param name="text">Optional text.</param>
        /// <returns>The created square shape.</returns>
        public VisioShape AddSquare(double x, double y, double size, string? text = null) =>
            AddSquare(x, y, size, text, DefaultUnit);

        /// <summary>
        /// Adds a circle shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="diameter">Diameter of the circle.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created circle shape.</returns>
        public VisioShape AddCircle(double x, double y, double diameter, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            // Avoid double-converting when passing the same variable for width/height
            double w = diameter, h = diameter;
            ApplyUnits(ref x, ref y, ref w, ref h, unit);
            var s = new VisioShape(NextId(), x, y, w, h, text ?? string.Empty) { NameU = "Circle" };
            _shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a circle using the page <see cref="DefaultUnit"/>.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="diameter">Diameter of the circle.</param>
        /// <param name="text">Optional text.</param>
        /// <returns>The created circle shape.</returns>
        public VisioShape AddCircle(double x, double y, double diameter, string? text = null) =>
            AddCircle(x, y, diameter, text, DefaultUnit);

        /// <summary>
        /// Adds an ellipse shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the ellipse.</param>
        /// <param name="height">Height of the ellipse.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created ellipse shape.</returns>
        public VisioShape AddEllipse(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Ellipse" };
            _shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds an ellipse using the page <see cref="DefaultUnit"/>.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the ellipse.</param>
        /// <param name="height">Height of the ellipse.</param>
        /// <param name="text">Optional text.</param>
        /// <returns>The created ellipse shape.</returns>
        public VisioShape AddEllipse(double x, double y, double width, double height, string? text = null) =>
            AddEllipse(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a diamond shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the diamond.</param>
        /// <param name="height">Height of the diamond.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created diamond shape.</returns>
        public VisioShape AddDiamond(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Diamond" };
            _shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a diamond using the page <see cref="DefaultUnit"/>.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the diamond.</param>
        /// <param name="height">Height of the diamond.</param>
        /// <param name="text">Optional text.</param>
        /// <returns>The created diamond shape.</returns>
        public VisioShape AddDiamond(double x, double y, double width, double height, string? text = null) =>
            AddDiamond(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a triangle shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the triangle's bounding box.</param>
        /// <param name="height">Height of the triangle's bounding box.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created triangle shape.</returns>
        public VisioShape AddTriangle(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Triangle" };
            _shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a triangle using the page <see cref="DefaultUnit"/>.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the triangle's bounding box.</param>
        /// <param name="height">Height of the triangle's bounding box.</param>
        /// <param name="text">Optional text.</param>
        /// <returns>The created triangle shape.</returns>
        public VisioShape AddTriangle(double x, double y, double width, double height, string? text = null) =>
            AddTriangle(x, y, width, height, text, DefaultUnit);

        private static int SideToIndex(VisioSide side) => side switch {
            VisioSide.Left => 0,
            VisioSide.Right => 1,
            VisioSide.Bottom => 2,
            VisioSide.Top => 3,
            _ => -1
        };

        /// <summary>
        /// Adds a connector between two shapes, optionally specifying side connection points.
        /// </summary>
        /// <param name="from">Source shape.</param>
        /// <param name="to">Target shape.</param>
        /// <param name="kind">Connector kind (straight, curved, etc.).</param>
        /// <param name="fromSide">Preferred side on the source shape.</param>
        /// <param name="toSide">Preferred side on the target shape.</param>
        /// <returns>The created connector.</returns>
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

