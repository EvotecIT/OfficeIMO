using System;

namespace OfficeIMO.Visio.Fluent {
    /// <summary>
    /// Fluent helper for configuring a connector between two shapes.
    /// </summary>
    public class VisioFluentConnector {
        private readonly VisioConnector _c;

        /// <summary>Initializes a new connector wrapper.</summary>
        /// <param name="connector">Underlying connector model.</param>
        internal VisioFluentConnector(VisioConnector connector) { _c = connector; }

        /// <summary>Sets connector kind to a straight line.</summary>
        public VisioFluentConnector Straight() { _c.Kind = ConnectorKind.Straight; return this; }

        /// <summary>Sets connector kind to right-angle (orthogonal) routing.</summary>
        public VisioFluentConnector RightAngle() { _c.Kind = ConnectorKind.RightAngle; return this; }

        /// <summary>Sets connector kind to curved routing.</summary>
        public VisioFluentConnector Curved() { _c.Kind = ConnectorKind.Curved; return this; }

        /// <summary>Sets a begin arrowhead style.</summary>
        /// <param name="arrow">Arrowhead enum value.</param>
        public VisioFluentConnector ArrowStart(EndArrow arrow) { _c.BeginArrow = arrow; return this; }

        /// <summary>Sets an end arrowhead style.</summary>
        /// <param name="arrow">Arrowhead enum value.</param>
        public VisioFluentConnector ArrowEnd(EndArrow arrow) { _c.EndArrow = arrow; return this; }

        /// <summary>Sets a connector label.</summary>
        /// <param name="text">Label text.</param>
        public VisioFluentConnector Label(string text) { _c.Label = text; return this; }

        /// <summary>Sets connector line weight (thickness) in inches.</summary>
        /// <param name="weight">Line weight in inches.</param>
        public VisioFluentConnector LineWeight(double weight) { _c.LineWeight = weight; return this; }

        /// <summary>Sets connector line pattern (Visio pattern index).</summary>
        /// <param name="pattern">Pattern index (0=None, 1=Solid, ...).</param>
        public VisioFluentConnector LinePattern(int pattern) { _c.LinePattern = pattern; return this; }
    }
}

