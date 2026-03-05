using System;
using SixLabors.ImageSharp;

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

        /// <summary>Sets connector kind to dynamic routing.</summary>
        public VisioFluentConnector Dynamic() { _c.Kind = ConnectorKind.Dynamic; return this; }

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

        /// <summary>Sets connector line color.</summary>
        /// <param name="color">Line color.</param>
        public VisioFluentConnector LineColor(Color color) { _c.LineColor = color; return this; }

        /// <summary>Connects both ends to explicit shape sides.</summary>
        /// <param name="fromSide">Preferred source side.</param>
        /// <param name="toSide">Preferred target side.</param>
        public VisioFluentConnector Sides(VisioSide fromSide, VisioSide toSide) {
            ApplySide(_c.From, fromSide, point => _c.FromConnectionPoint = point);
            ApplySide(_c.To, toSide, point => _c.ToConnectionPoint = point);
            return this;
        }

        /// <summary>Connects the start of the connector to an explicit side.</summary>
        /// <param name="side">Preferred source side.</param>
        public VisioFluentConnector FromSide(VisioSide side) {
            ApplySide(_c.From, side, point => _c.FromConnectionPoint = point);
            return this;
        }

        /// <summary>Connects the end of the connector to an explicit side.</summary>
        /// <param name="side">Preferred target side.</param>
        public VisioFluentConnector ToSide(VisioSide side) {
            ApplySide(_c.To, side, point => _c.ToConnectionPoint = point);
            return this;
        }

        private static void ApplySide(VisioShape shape, VisioSide side, Action<VisioConnectionPoint?> assign) {
            if (side == VisioSide.Auto) {
                assign(null);
                return;
            }

            assign(shape.EnsureSideConnectionPoint(side));
        }
    }
}

