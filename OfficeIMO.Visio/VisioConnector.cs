using System;
using System.Globalization;
using SixLabors.ImageSharp;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Connects two shapes together.
    /// </summary>
    public class VisioConnector {
        private static int _idCounter;

        /// <summary>
        /// Initializes a new instance of the <see cref="VisioConnector"/> class connecting two shapes.
        /// </summary>
        /// <param name="from">Shape from which the connector starts.</param>
        /// <param name="to">Shape at which the connector ends.</param>
        public VisioConnector(VisioShape from, VisioShape to) : this(GetNextId(from, to), from, to) {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="VisioConnector"/> class with an explicit identifier.
        /// </summary>
        /// <param name="id">Identifier of the connector.</param>
        /// <param name="from">Shape from which the connector starts.</param>
        /// <param name="to">Shape at which the connector ends.</param>
        public VisioConnector(string id, VisioShape from, VisioShape to) {
            Id = id;
            From = from;
            To = to;
            LineColor = Color.Black;
            LineWeight = 0.0138889;
            LinePattern = 1; // Solid line
        }

        /// <summary>
        /// Identifier of the connector.
        /// </summary>
        public string Id { get; }

        /// <summary>
        /// Shape from which the connector starts.
        /// </summary>
        public VisioShape From { get; }

        /// <summary>
        /// Shape at which the connector ends.
        /// </summary>
        public VisioShape To { get; }

        /// <summary>
        /// Connection point on the starting shape.
        /// </summary>
        public VisioConnectionPoint? FromConnectionPoint { get; set; }

        /// <summary>
        /// Connection point on the ending shape.
        /// </summary>
        public VisioConnectionPoint? ToConnectionPoint { get; set; }

        /// <summary>
        /// Gets or sets the kind of connector.
        /// </summary>
        public ConnectorKind Kind { get; set; } = ConnectorKind.Straight;

        /// <summary>
        /// Gets or sets the end arrow style.
        /// </summary>
        public int? EndArrow { get; set; }
        
        /// <summary>
        /// Line color of the connector.
        /// </summary>
        public Color LineColor { get; set; }
        
        /// <summary>
        /// Line weight (thickness) of the connector.
        /// </summary>
        public double LineWeight { get; set; }
        
        /// <summary>
        /// Line pattern (0=None, 1=Solid, 2=Dashed, etc.).
        /// </summary>
        public int LinePattern { get; set; }

        private static string GetNextId(VisioShape from, VisioShape to) {
            int fromId = int.TryParse(from.Id, out int fi) ? fi : 0;
            int toId = int.TryParse(to.Id, out int ti) ? ti : 0;
            int newId = Math.Max(Math.Max(fromId, toId) + 1, _idCounter + 1);
            _idCounter = newId;
            return newId.ToString(CultureInfo.InvariantCulture);
        }
    }
}

