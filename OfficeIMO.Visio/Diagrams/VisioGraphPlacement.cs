using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>Explicit graph bounds in the builder's page unit, with a center pin and bottom-up Y axis.</summary>
    public sealed class VisioGraphPlacement {
        /// <summary>Creates validated fixed bounds. Width and height must be positive.</summary>
        public VisioGraphPlacement(double pinX, double pinY, double width, double height) {
            RequireFinite(pinX, nameof(pinX)); RequireFinite(pinY, nameof(pinY));
            RequireFinite(width, nameof(width)); RequireFinite(height, nameof(height));
            if (width <= 0 || height <= 0) throw new ArgumentOutOfRangeException(nameof(width), "Bounds must have positive dimensions.");
            PinX = pinX; PinY = pinY; Width = width; Height = height;
        }
        /// <summary>Gets the horizontal center.</summary>
        public double PinX { get; }
        /// <summary>Gets the vertical center.</summary>
        public double PinY { get; }
        /// <summary>Gets the width.</summary>
        public double Width { get; }
        /// <summary>Gets the height.</summary>
        public double Height { get; }
        internal static void RequireFinite(double value, string name) {
            if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(name);
        }
    }

    /// <summary>A complete fixed polyline in page units, including both glued endpoints.</summary>
    public sealed class VisioGraphRoute {
        private readonly IReadOnlyList<(double X, double Y)> _points;
        /// <summary>Snapshots a finite route containing at least two points.</summary>
        public VisioGraphRoute(IEnumerable<VisioConnectorWaypoint> points) {
            if (points == null) throw new ArgumentNullException(nameof(points));
            _points = Array.AsReadOnly(points.Select(point => {
                if (point == null) throw new ArgumentException("Route points cannot be null.", nameof(points));
                VisioGraphPlacement.RequireFinite(point.X, nameof(points));
                VisioGraphPlacement.RequireFinite(point.Y, nameof(points));
                return (point.X, point.Y);
            }).ToArray());
            if (_points.Count < 2) throw new ArgumentException("A route requires both endpoints.", nameof(points));
        }
        /// <summary>Gets the immutable route snapshot, including both endpoints, in page units.</summary>
        public IReadOnlyList<(double X, double Y)> Points => _points;
    }
}
