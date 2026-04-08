using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a single page within a Visio document.
    /// </summary>
    public class VisioPage {
        internal sealed class PreservedConnectChildEntry {
            public PreservedConnectChildEntry(XElement rawElement) {
                RawElement = new XElement(rawElement);
            }

            public PreservedConnectChildEntry(VisioConnector connector, VisioConnectorEndpointScope endpointScope) {
                Connector = connector;
                EndpointScope = endpointScope;
            }

            public XElement? RawElement { get; }

            public VisioConnector? Connector { get; }

            public VisioConnectorEndpointScope? EndpointScope { get; }
        }

        internal sealed class PreservedConnectRowEntry {
            public PreservedConnectRowEntry(XElement rawElement) {
                RawElement = new XElement(rawElement);
            }

            public PreservedConnectRowEntry(VisioConnector connector, VisioConnectorEndpointScope endpointScope) {
                Connector = connector;
                EndpointScope = endpointScope;
            }

            public XElement? RawElement { get; }

            public VisioConnector? Connector { get; }

            public VisioConnectorEndpointScope? EndpointScope { get; }
        }

        private readonly List<VisioShape> _shapes = new();
        private readonly List<VisioConnector> _connectors = new();
        private readonly IList<VisioShape> _shapeCollection;
        private readonly IList<VisioConnector> _connectorCollection;
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
            _shapeCollection = new ShapeCollection(this);
            _connectorCollection = new ConnectorCollection(this);
        }

        /// <summary>
        /// Gets the identifier of the page within the document.
        /// </summary>
        public int Id { get; internal set; }

        internal VisioDocument? OwnerDocument { get; set; }

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
            set {
                if (!Enum.IsDefined(typeof(VisioMeasurementUnit), value)) {
                    throw new ArgumentOutOfRangeException(nameof(value));
                }

                if (_scaleMeasurementUnit == value) {
                    return;
                }

                VisioMeasurementUnit previous = _scaleMeasurementUnit;
                _scaleMeasurementUnit = value;
                NormalizeScaleOverrides(previous, value);
            }
        }

        /// <summary>
        /// Gets or sets the page scale (the ratio between page units and real-world units).
        /// </summary>
        public VisioScaleSetting PageScale {
            get {
                return _pageScaleOverride ?? VisioScaleSetting.FromUnit(ScaleMeasurementUnit);
            }
            set => _pageScaleOverride = value.Normalized();
        }

        /// <summary>
        /// Gets or sets the drawing scale (the ratio between drawing units and real-world units).
        /// </summary>
        public VisioScaleSetting DrawingScale {
            get {
                return _drawingScaleOverride ?? VisioScaleSetting.FromUnit(ScaleMeasurementUnit);
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

        internal VisioScaleSetting GetEffectivePageScale() => _pageScaleOverride ?? VisioScaleSetting.FromUnit(ScaleMeasurementUnit);

        internal VisioScaleSetting GetEffectiveDrawingScale() => _drawingScaleOverride ?? VisioScaleSetting.FromUnit(ScaleMeasurementUnit);

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

        private void NormalizeScaleOverrides(VisioMeasurementUnit previousUnit, VisioMeasurementUnit newUnit) {
            if (_pageScaleOverride.HasValue && _pageScaleOverride.Value.Unit == previousUnit) {
                _pageScaleOverride = _pageScaleOverride.Value.ConvertTo(newUnit);
            }

            if (_drawingScaleOverride.HasValue && _drawingScaleOverride.Value.Unit == previousUnit) {
                _drawingScaleOverride = _drawingScaleOverride.Value.ConvertTo(newUnit);
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

        internal IList<XElement> PreservedPageSheetCells { get; } = new List<XElement>();

        internal IList<XElement> PreservedPageSheetSections { get; } = new List<XElement>();

        internal IList<XAttribute> PreservedPageAttributes { get; } = new List<XAttribute>();

        internal IList<XAttribute> PreservedPageContentAttributes { get; } = new List<XAttribute>();

        internal IList<XElement> PreservedPageContentElements { get; } = new List<XElement>();

        internal IList<XAttribute> PreservedShapesContainerAttributes { get; } = new List<XAttribute>();

        internal IList<XElement> PreservedShapesContainerElements { get; } = new List<XElement>();

        internal IList<XAttribute> PreservedConnectsAttributes { get; } = new List<XAttribute>();

        internal IList<XElement> PreservedConnectsElements { get; } = new List<XElement>();

        internal IList<PreservedConnectChildEntry> PreservedConnectChildren { get; } = new List<PreservedConnectChildEntry>();

        internal IList<PreservedConnectRowEntry> PreservedConnectRows { get; } = new List<PreservedConnectRowEntry>();

        /// <summary>
        /// Shapes placed on the page.
        /// </summary>
        public IList<VisioShape> Shapes => _shapeCollection;

        /// <summary>
        /// Connectors placed on the page.
        /// </summary>
        public IList<VisioConnector> Connectors => _connectorCollection;

        private string NextId(VisioConnector? ignoredConnector = null) {
            HashSet<int> usedIds = new();

            void Reserve(string? id) {
                if (int.TryParse(id, out int numericId) && numericId > 0) {
                    usedIds.Add(numericId);
                }
            }

            void VisitShape(VisioShape shape) {
                Reserve(shape.Id);
                foreach (VisioShape child in shape.Children) {
                    VisitShape(child);
                }
            }

            foreach (VisioShape shape in _shapes) {
                VisitShape(shape);
            }

            foreach (VisioConnector connector in _connectors) {
                if (ReferenceEquals(connector, ignoredConnector)) {
                    continue;
                }

                Reserve(connector.Id);
            }

            int nextId = 1;
            while (usedIds.Contains(nextId)) {
                nextId++;
            }

            return nextId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private void PrepareConnectorForPage(VisioConnector connector, VisioConnector? ignoredConnector = null) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            if (connector.HasAutomaticId) {
                connector.Id = NextId(ignoredConnector);
            }
        }

        private void PrepareShapeForPage(VisioShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

             if (_shapes.Contains(shape)) {
                throw new InvalidOperationException("The shape is already part of this page.");
            }

            if (shape.Parent != null) {
                throw new InvalidOperationException("A child shape must be removed from its parent before being added to a page.");
            }

            shape.Parent = null;
            shape.NormalizeDescendantParentLinks();
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
            Shapes.Add(s);
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
        /// Adds a flowchart process shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the process box.</param>
        /// <param name="height">Height of the process box.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created process shape.</returns>
        public VisioShape AddProcess(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            var s = AddRectangle(x, y, width, height, text, unit);
            s.NameU = "Process";
            return s;
        }

        /// <summary>
        /// Adds a flowchart process shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddProcess(double x, double y, double width, double height, string? text = null) =>
            AddProcess(x, y, width, height, text, DefaultUnit);

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
            Shapes.Add(s);
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
            Shapes.Add(s);
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
            Shapes.Add(s);
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
        /// Adds a flowchart decision shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the decision shape.</param>
        /// <param name="height">Height of the decision shape.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created decision shape.</returns>
        public VisioShape AddDecision(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            var s = AddDiamond(x, y, width, height, text, unit);
            s.NameU = "Decision";
            return s;
        }

        /// <summary>
        /// Adds a flowchart decision shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddDecision(double x, double y, double width, double height, string? text = null) =>
            AddDecision(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a flowchart data shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the data shape.</param>
        /// <param name="height">Height of the data shape.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created data shape.</returns>
        public VisioShape AddData(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Data" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a flowchart data shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddData(double x, double y, double width, double height, string? text = null) =>
            AddData(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a flowchart preparation shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the preparation shape.</param>
        /// <param name="height">Height of the preparation shape.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created preparation shape.</returns>
        public VisioShape AddPreparation(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Preparation" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a flowchart preparation shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddPreparation(double x, double y, double width, double height, string? text = null) =>
            AddPreparation(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a parallelogram shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the parallelogram.</param>
        /// <param name="height">Height of the parallelogram.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created parallelogram shape.</returns>
        public VisioShape AddParallelogram(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Parallelogram" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a parallelogram shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddParallelogram(double x, double y, double width, double height, string? text = null) =>
            AddParallelogram(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a hexagon shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the hexagon.</param>
        /// <param name="height">Height of the hexagon.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created hexagon shape.</returns>
        public VisioShape AddHexagon(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Hexagon" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a hexagon shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddHexagon(double x, double y, double width, double height, string? text = null) =>
            AddHexagon(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a trapezoid shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the trapezoid.</param>
        /// <param name="height">Height of the trapezoid.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created trapezoid shape.</returns>
        public VisioShape AddTrapezoid(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Trapezoid" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a trapezoid shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddTrapezoid(double x, double y, double width, double height, string? text = null) =>
            AddTrapezoid(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a pentagon shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the pentagon.</param>
        /// <param name="height">Height of the pentagon.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created pentagon shape.</returns>
        public VisioShape AddPentagon(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Pentagon" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a pentagon shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddPentagon(double x, double y, double width, double height, string? text = null) =>
            AddPentagon(x, y, width, height, text, DefaultUnit);

        /// <summary>
         /// Adds a flowchart manual operation shape.
         /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the manual operation shape.</param>
        /// <param name="height">Height of the manual operation shape.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created manual operation shape.</returns>
        public VisioShape AddManualOperation(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Manual operation" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a flowchart manual operation shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddManualOperation(double x, double y, double width, double height, string? text = null) =>
            AddManualOperation(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a flowchart off-page reference shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the off-page reference shape.</param>
        /// <param name="height">Height of the off-page reference shape.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created off-page reference shape.</returns>
        public VisioShape AddOffPageReference(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Off-page reference" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a flowchart off-page reference shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddOffPageReference(double x, double y, double width, double height, string? text = null) =>
            AddOffPageReference(x, y, width, height, text, DefaultUnit);

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
            Shapes.Add(s);
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

        /// <summary>
        /// Adds a connector between two shapes, optionally specifying side connection points.
        /// </summary>
        /// <param name="from">Source shape.</param>
        /// <param name="to">Target shape.</param>
        /// <param name="kind">Connector kind (straight, curved, etc.).</param>
        /// <param name="fromSide">Preferred side on the source shape.</param>
        /// <param name="toSide">Preferred side on the target shape.</param>
        /// <returns>The created connector.</returns>
        public VisioConnector AddConnector(VisioShape from, VisioShape to, ConnectorKind kind = ConnectorKind.Dynamic, VisioSide fromSide = VisioSide.Auto, VisioSide toSide = VisioSide.Auto) {
            var conn = new VisioConnector(NextId(), from, to) { Kind = kind };
            if (fromSide != VisioSide.Auto) conn.FromConnectionPoint = from.EnsureSideConnectionPoint(fromSide);
            if (toSide != VisioSide.Auto) conn.ToConnectionPoint = to.EnsureSideConnectionPoint(toSide);
            Connectors.Add(conn);
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
        /// <param name="unit">
        /// Optional measurement unit. When omitted, values are interpreted using
        /// the page <see cref="DefaultUnit"/>.
        /// </param>
        /// <returns>The created shape.</returns>
        public VisioShape AddShape(string id, VisioMaster master, double x, double y, double w, double h, string? text = null, VisioMeasurementUnit? unit = null) {
            VisioMeasurementUnit effectiveUnit = unit ?? DefaultUnit;
            x = x.ToInches(effectiveUnit);
            y = y.ToInches(effectiveUnit);
            w = w.ToInches(effectiveUnit);
            h = h.ToInches(effectiveUnit);

            VisioShape shape = new VisioShape(id, x, y, w, h, text ?? string.Empty) {
                Master = master,
                NameU = master.NameU
            };
            Shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// Adds a shape using a document-registered master by its NameU.
        /// </summary>
        /// <param name="id">Identifier of the shape.</param>
        /// <param name="masterNameU">Registered master universal name.</param>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="w">Width.</param>
        /// <param name="h">Height.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit for the provided values.</param>
        /// <returns>The created shape.</returns>
        public VisioShape AddShape(string id, string masterNameU, double x, double y, double w, double h, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            if (OwnerDocument == null) {
                throw new InvalidOperationException("This page is not attached to a VisioDocument, so master lookup by name is unavailable.");
            }

            VisioMaster master = OwnerDocument.GetMaster(masterNameU);
            x = x.ToInches(unit);
            y = y.ToInches(unit);
            w = w.ToInches(unit);
            h = h.ToInches(unit);

            VisioShape shape = new VisioShape(id, x, y, w, h, text ?? string.Empty) {
                Master = master,
                NameU = master.NameU
            };
            Shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// Adds a shape using the page <see cref="DefaultUnit"/> and a document-registered master.
        /// </summary>
        public VisioShape AddShape(string id, string masterNameU, double x, double y, double w, double h, string? text = null) =>
            AddShape(id, masterNameU, x, y, w, h, text, DefaultUnit);

        /// <summary>
        /// Moves a shape from its current location in the page hierarchy into the provided group shape.
        /// </summary>
        /// <param name="shape">The shape to move.</param>
        /// <param name="newParent">The group that should own the shape after the move.</param>
        /// <param name="childIndex">
        /// Optional insertion index within the target group's children.
        /// Use <c>-1</c> to append.
        /// </param>
        public void ReparentShape(VisioShape shape, VisioShape newParent, int childIndex = -1) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            if (newParent == null) {
                throw new ArgumentNullException(nameof(newParent));
            }

            if (childIndex < -1) {
                throw new ArgumentOutOfRangeException(nameof(childIndex), "Child index must be -1 or greater.");
            }

            if (ReferenceEquals(shape, newParent)) {
                throw new InvalidOperationException("A shape cannot be reparented into itself.");
            }

            if (!TryFindShapeCollection(shape, out IList<VisioShape>? currentCollection, out int currentIndex)) {
                throw new InvalidOperationException("The shape is not part of this page.");
            }

            if (!TryFindShapeCollection(newParent, out _, out _)) {
                throw new InvalidOperationException("The target parent shape is not part of this page.");
            }

            if (ReferenceEquals(currentCollection, newParent.Children)) {
                if (childIndex < 0 || childIndex == currentIndex) {
                    return;
                }

                if (childIndex > currentCollection.Count) {
                    throw new ArgumentOutOfRangeException(nameof(childIndex), "Child index cannot exceed the number of children in the target group.");
                }

                currentCollection.RemoveAt(currentIndex);
                if (childIndex > currentIndex) {
                    childIndex--;
                }

                currentCollection.Insert(childIndex, shape);
                return;
            }

            if (childIndex > newParent.Children.Count) {
                throw new ArgumentOutOfRangeException(nameof(childIndex), "Child index cannot exceed the number of children in the target group.");
            }

            IList<VisioShape> currentOwnerCollection = currentCollection!;
            currentOwnerCollection.RemoveAt(currentIndex);
            try {
                if (childIndex < 0) {
                    newParent.Children.Add(shape);
                } else {
                    newParent.Children.Insert(childIndex, shape);
                }
            } catch {
                currentOwnerCollection.Insert(currentIndex, shape);
                throw;
            }
        }

        /// <summary>
        /// Removes a group shape and promotes its children into the group's former position.
        /// </summary>
        /// <param name="group">The group to ungroup.</param>
        /// <returns>The children that were promoted.</returns>
        public IReadOnlyList<VisioShape> UngroupShape(VisioShape group) {
            if (group == null) {
                throw new ArgumentNullException(nameof(group));
            }

            if (!TryFindShapeCollection(group, out IList<VisioShape>? ownerCollection, out int index)) {
                throw new InvalidOperationException("The group shape is not part of this page.");
            }

            IList<VisioShape> resolvedOwnerCollection = ownerCollection!;

            if (group.Children.Count == 0) {
                resolvedOwnerCollection.RemoveAt(index);
                return Array.Empty<VisioShape>();
            }

            List<VisioShape> promotedChildren = new(group.Children);
            resolvedOwnerCollection.RemoveAt(index);
            group.Children.Clear();

            for (int i = 0; i < promotedChildren.Count; i++) {
                resolvedOwnerCollection.Insert(index + i, promotedChildren[i]);
            }

            return promotedChildren;
        }

        /// <summary>
        /// Reconnects the start of an existing connector to a different shape.
        /// </summary>
        public void ReconnectConnectorStart(VisioConnector connector, VisioShape newFrom, VisioSide side = VisioSide.Auto) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            if (newFrom == null) {
                throw new ArgumentNullException(nameof(newFrom));
            }

            EnsureConnectorBelongsToPage(connector);
            EnsureShapeBelongsToPage(newFrom, "The source shape is not part of this page.");

            connector.From = newFrom;
            connector.FromConnectionPoint = ResolveConnectionPoint(newFrom, side);
            connector.PreservedFromConnectionCell = null;
            connector.PreservedBeginConnectAttributes.Clear();
            connector.PreservedBeginConnectAttributeOrder.Clear();
        }

        /// <summary>
        /// Reconnects the end of an existing connector to a different shape.
        /// </summary>
        public void ReconnectConnectorEnd(VisioConnector connector, VisioShape newTo, VisioSide side = VisioSide.Auto) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            if (newTo == null) {
                throw new ArgumentNullException(nameof(newTo));
            }

            EnsureConnectorBelongsToPage(connector);
            EnsureShapeBelongsToPage(newTo, "The target shape is not part of this page.");

            connector.To = newTo;
            connector.ToConnectionPoint = ResolveConnectionPoint(newTo, side);
            connector.PreservedToConnectionCell = null;
            connector.PreservedEndConnectAttributes.Clear();
            connector.PreservedEndConnectAttributeOrder.Clear();
        }

        /// <summary>
        /// Reconnects both ends of an existing connector.
        /// </summary>
        public void ReconnectConnector(VisioConnector connector, VisioShape newFrom, VisioShape newTo, VisioSide fromSide = VisioSide.Auto, VisioSide toSide = VisioSide.Auto) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            if (newFrom == null) {
                throw new ArgumentNullException(nameof(newFrom));
            }

            if (newTo == null) {
                throw new ArgumentNullException(nameof(newTo));
            }

            EnsureConnectorBelongsToPage(connector);
            EnsureShapeBelongsToPage(newFrom, "The source shape is not part of this page.");
            EnsureShapeBelongsToPage(newTo, "The target shape is not part of this page.");

            connector.From = newFrom;
            connector.To = newTo;
            connector.FromConnectionPoint = ResolveConnectionPoint(newFrom, fromSide);
            connector.ToConnectionPoint = ResolveConnectionPoint(newTo, toSide);
            connector.PreservedFromConnectionCell = null;
            connector.PreservedToConnectionCell = null;
            connector.PreservedBeginConnectAttributes.Clear();
            connector.PreservedEndConnectAttributes.Clear();
            connector.PreservedBeginConnectAttributeOrder.Clear();
            connector.PreservedEndConnectAttributeOrder.Clear();
        }

        /// <summary>
        /// Retargets all connector endpoints on this page that currently reference one shape to another.
        /// </summary>
        /// <param name="oldShape">The existing shape referenced by matching connectors.</param>
        /// <param name="newShape">The replacement shape that matching connectors should reference.</param>
        /// <param name="endpointScope">Controls whether start points, end points, or both are updated.</param>
        /// <param name="fromSide">The side to glue to when a start point is updated.</param>
        /// <param name="toSide">The side to glue to when an end point is updated.</param>
        /// <returns>The connectors that were updated.</returns>
        public IReadOnlyList<VisioConnector> RetargetConnectors(VisioShape oldShape, VisioShape newShape, VisioConnectorEndpointScope endpointScope = VisioConnectorEndpointScope.Both, VisioSide fromSide = VisioSide.Auto, VisioSide toSide = VisioSide.Auto) {
            if (oldShape == null) {
                throw new ArgumentNullException(nameof(oldShape));
            }

            if (newShape == null) {
                throw new ArgumentNullException(nameof(newShape));
            }

            EnsureShapeBelongsToPage(newShape, "The replacement shape is not part of this page.");

            if (ReferenceEquals(oldShape, newShape)) {
                return Array.Empty<VisioConnector>();
            }

            List<VisioConnector> updatedConnectors = new();
            for (int i = 0; i < _connectors.Count; i++) {
                VisioConnector connector = _connectors[i];
                bool updateStart = endpointScope != VisioConnectorEndpointScope.End && ReferenceEquals(connector.From, oldShape);
                bool updateEnd = endpointScope != VisioConnectorEndpointScope.Start && ReferenceEquals(connector.To, oldShape);

                if (!updateStart && !updateEnd) {
                    continue;
                }

                if (updateStart && updateEnd) {
                    ReconnectConnector(connector, newShape, newShape, fromSide, toSide);
                } else if (updateStart) {
                    ReconnectConnectorStart(connector, newShape, fromSide);
                } else {
                    ReconnectConnectorEnd(connector, newShape, toSide);
                }

                updatedConnectors.Add(connector);
            }

            if (updatedConnectors.Count == 0 && !TryFindShapeCollection(oldShape, out _, out _)) {
                throw new InvalidOperationException("The original shape is not part of this page or referenced by any connector on this page.");
            }

            return updatedConnectors;
        }

        private void EnsureConnectorBelongsToPage(VisioConnector connector) {
            if (!_connectors.Contains(connector)) {
                throw new InvalidOperationException("The connector is not part of this page.");
            }
        }

        private void EnsureShapeBelongsToPage(VisioShape shape, string message) {
            if (!TryFindShapeCollection(shape, out _, out _)) {
                throw new InvalidOperationException(message);
            }
        }

        private static VisioConnectionPoint? ResolveConnectionPoint(VisioShape shape, VisioSide side) {
            return side == VisioSide.Auto ? null : shape.EnsureSideConnectionPoint(side);
        }

        private bool TryFindShapeCollection(VisioShape target, out IList<VisioShape>? ownerCollection, out int index) {
            return TryFindShapeCollection(_shapeCollection, target, out ownerCollection, out index);
        }

        private static bool TryFindShapeCollection(IList<VisioShape> collection, VisioShape target, out IList<VisioShape>? ownerCollection, out int index) {
            for (int i = 0; i < collection.Count; i++) {
                VisioShape shape = collection[i];
                if (ReferenceEquals(shape, target)) {
                    ownerCollection = collection;
                    index = i;
                    return true;
                }

                if (TryFindShapeCollection(shape.Children, target, out ownerCollection, out index)) {
                    return true;
                }
            }

            ownerCollection = null;
            index = -1;
            return false;
        }

        private sealed class ShapeCollection : IList<VisioShape> {
            private readonly VisioPage _page;

            public ShapeCollection(VisioPage page) {
                _page = page;
            }

            public VisioShape this[int index] {
                get => _page._shapes[index];
                set {
                    if (ReferenceEquals(_page._shapes[index], value)) {
                        return;
                    }

                    _page.PrepareShapeForPage(value);
                    _page._shapes[index] = value;
                }
            }

            public int Count => _page._shapes.Count;

            public bool IsReadOnly => false;

            public void Add(VisioShape item) {
                _page.PrepareShapeForPage(item);
                _page._shapes.Add(item);
            }

            public void Clear() => _page._shapes.Clear();

            public bool Contains(VisioShape item) => _page._shapes.Contains(item);

            public void CopyTo(VisioShape[] array, int arrayIndex) => _page._shapes.CopyTo(array, arrayIndex);

            public IEnumerator<VisioShape> GetEnumerator() => _page._shapes.GetEnumerator();

            public int IndexOf(VisioShape item) => _page._shapes.IndexOf(item);

            public void Insert(int index, VisioShape item) {
                _page.PrepareShapeForPage(item);
                _page._shapes.Insert(index, item);
            }

            public bool Remove(VisioShape item) => _page._shapes.Remove(item);

            public void RemoveAt(int index) => _page._shapes.RemoveAt(index);

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
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
            Connectors.Add(connector);
            return connector;
        }

        private sealed class ConnectorCollection : IList<VisioConnector> {
            private readonly VisioPage _page;

            public ConnectorCollection(VisioPage page) {
                _page = page;
            }

            public VisioConnector this[int index] {
                get => _page._connectors[index];
                set {
                    _page.PrepareConnectorForPage(value, _page._connectors[index]);
                    _page._connectors[index] = value;
                }
            }

            public int Count => _page._connectors.Count;

            public bool IsReadOnly => false;

            public void Add(VisioConnector item) {
                _page.PrepareConnectorForPage(item);
                _page._connectors.Add(item);
            }

            public void Clear() => _page._connectors.Clear();

            public bool Contains(VisioConnector item) => _page._connectors.Contains(item);

            public void CopyTo(VisioConnector[] array, int arrayIndex) => _page._connectors.CopyTo(array, arrayIndex);

            public IEnumerator<VisioConnector> GetEnumerator() => _page._connectors.GetEnumerator();

            public int IndexOf(VisioConnector item) => _page._connectors.IndexOf(item);

            public void Insert(int index, VisioConnector item) {
                _page.PrepareConnectorForPage(item);
                _page._connectors.Insert(index, item);
            }

            public bool Remove(VisioConnector item) => _page._connectors.Remove(item);

            public void RemoveAt(int index) => _page._connectors.RemoveAt(index);

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
        }
    }
}

