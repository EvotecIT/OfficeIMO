using System.Collections.Generic;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a single page within a Visio document.
    /// </summary>
    public class VisioPage {
        private readonly List<VisioShape> _shapes = new();
        private readonly List<VisioConnector> _connectors = new();

        public VisioPage(string name) {
            Name = name;
            NameU = name;
            ViewScale = 1;
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

