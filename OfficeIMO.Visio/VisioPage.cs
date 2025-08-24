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

