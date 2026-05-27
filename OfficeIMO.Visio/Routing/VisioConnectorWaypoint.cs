using System;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Absolute page coordinate used as a connector routing waypoint.
    /// </summary>
    public sealed class VisioConnectorWaypoint {
        /// <summary>
        /// Initializes a new connector waypoint.
        /// </summary>
        /// <param name="x">Page X coordinate.</param>
        /// <param name="y">Page Y coordinate.</param>
        public VisioConnectorWaypoint(double x, double y) {
            X = x;
            Y = y;
        }

        /// <summary>Page X coordinate.</summary>
        public double X { get; set; }

        /// <summary>Page Y coordinate.</summary>
        public double Y { get; set; }

        /// <summary>
        /// Creates a waypoint at the specified page coordinate.
        /// </summary>
        /// <param name="x">Page X coordinate.</param>
        /// <param name="y">Page Y coordinate.</param>
        public static VisioConnectorWaypoint At(double x, double y) {
            return new VisioConnectorWaypoint(x, y);
        }

        /// <inheritdoc />
        public override string ToString() {
            return FormattableString.Invariant($"{X},{Y}");
        }
    }
}
