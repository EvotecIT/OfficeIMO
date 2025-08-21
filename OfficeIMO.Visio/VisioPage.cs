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

        public VisioPage(string name) : this(name, 8.26771653543307, 11.69291338582677) {
        }

        public VisioPage(string name, double widthInches, double heightInches) {
            Name = name;
            NameU = name;
            _width = widthInches;
            _height = heightInches;
            ViewScale = -1;
            ViewCenterX = widthInches / 2;
            ViewCenterY = heightInches / 2;
        }

        public int Id { get; internal set; }

        /// <summary>
        /// Gets the page name.
        /// </summary>
        public string Name { get; }

        public string? NameU { get; set; }

        public double ViewScale { get; set; }

        public double ViewCenterX { get; set; }

        public double ViewCenterY { get; set; }

        public double Width {
            get => _width;
            set {
                _width = value;
                ViewCenterX = value / 2;
            }
        }

        public double WidthCentimeters {
            get => _width.FromInches(VisioMeasurementUnit.Centimeters);
            set => Width = value.ToInches(VisioMeasurementUnit.Centimeters);
        }

        public double Height {
            get => _height;
            set {
                _height = value;
                ViewCenterY = value / 2;
            }
        }

        public double HeightCentimeters {
            get => _height.FromInches(VisioMeasurementUnit.Centimeters);
            set => Height = value.ToInches(VisioMeasurementUnit.Centimeters);
        }

        [System.Obsolete("Use Width instead")]
        public double PageWidth {
            get => Width;
            set => Width = value;
        }

        [System.Obsolete("Use Height instead")]
        public double PageHeight {
            get => Height;
            set => Height = value;
        }

        public bool GridVisible {
            get => _gridVisible;
            set => _gridVisible = value;
        }

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

        public VisioPage Size(double w, double h, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            Width = w.ToInches(unit);
            Height = h.ToInches(unit);
            return this;
        }

        public VisioPage Grid(bool visible, bool snap) {
            GridVisible = visible;
            Snap = snap;
            return this;
        }

        public VisioShape AddShape(string id, VisioMaster master, double x, double y, double w, double h, string? text = null) {
            VisioShape shape = new VisioShape(id, x, y, w, h, text ?? string.Empty) { Master = master };
            _shapes.Add(shape);
            return shape;
        }

        public VisioConnector AddConnector(string id, VisioShape from, VisioShape to, ConnectorKind kind) {
            VisioConnector connector = new VisioConnector(id, from, to) { Kind = kind };
            _connectors.Add(connector);
            return connector;
        }
    }
}

