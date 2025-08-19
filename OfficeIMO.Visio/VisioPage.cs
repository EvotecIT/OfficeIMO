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

        /// <summary>
        /// Shapes placed on the page.
        /// </summary>
        public IList<VisioShape> Shapes => _shapes;

        /// <summary>
        /// Connectors placed on the page.
        /// </summary>
        public IList<VisioConnector> Connectors => _connectors;
    }
}

