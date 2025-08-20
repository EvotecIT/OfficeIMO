using System.Collections.Generic;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a single page within a Visio document.
    /// </summary>
    public class VisioPage {
        private readonly List<VisioShape> _shapes = new();
        private readonly List<VisioConnector> _connectors = new();
        private double _pageWidth = 8.26771653543307; // A4 width in inches
        private double _pageHeight = 11.69291338582677; // A4 height in inches
        private bool _gridVisible;
        private bool _snap = true;

        public VisioPage(string name) {
            Name = name;
            NameU = name;
            ViewScale = -1;
            ViewCenterX = _pageWidth / 2;
            ViewCenterY = _pageHeight / 2;
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

        public double PageWidth {
            get => _pageWidth;
            set {
                _pageWidth = value;
                ViewCenterX = value / 2;
            }
        }

        public double PageHeight {
            get => _pageHeight;
            set {
                _pageHeight = value;
                ViewCenterY = value / 2;
            }
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

        public VisioPage Size(double w, double h) {
            PageWidth = w;
            PageHeight = h;
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

