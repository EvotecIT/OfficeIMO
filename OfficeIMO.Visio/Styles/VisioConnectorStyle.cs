using System;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Reusable visual style for Visio connectors.
    /// </summary>
    public sealed class VisioConnectorStyle {
        /// <summary>
        /// Initializes a new connector style.
        /// </summary>
        /// <param name="lineColor">Line color.</param>
        /// <param name="lineWeight">Line weight in inches.</param>
        /// <param name="linePattern">Line pattern index.</param>
        /// <param name="endArrow">Optional end arrow style.</param>
        public VisioConnectorStyle(Color lineColor, double lineWeight = VisioShape.DefaultLineWeight, int linePattern = 1, EndArrow? endArrow = OfficeIMO.Visio.EndArrow.Triangle) {
            LineColor = lineColor;
            LineWeight = lineWeight;
            LinePattern = linePattern;
            EndArrow = endArrow;
        }

        /// <summary>Line color.</summary>
        public Color LineColor { get; set; }

        /// <summary>Line weight in inches.</summary>
        public double LineWeight { get; set; }

        /// <summary>Line pattern index.</summary>
        public int LinePattern { get; set; }

        /// <summary>Optional begin arrow style. Null leaves the current value unchanged.</summary>
        public EndArrow? BeginArrow { get; set; }

        /// <summary>Optional end arrow style. Null leaves the current value unchanged.</summary>
        public EndArrow? EndArrow { get; set; }

        /// <summary>Optional connector kind. Null leaves the current value unchanged.</summary>
        public ConnectorKind? Kind { get; set; }

        /// <summary>Optional connector label text style to apply with this connector style.</summary>
        public VisioTextStyle? TextStyle { get; set; }

        /// <summary>Creates a detached copy of this style.</summary>
        public VisioConnectorStyle Clone() {
            return new VisioConnectorStyle(LineColor, LineWeight, LinePattern, EndArrow) {
                BeginArrow = BeginArrow,
                Kind = Kind,
                TextStyle = TextStyle?.Clone()
            };
        }

        /// <summary>Applies this style to a connector.</summary>
        /// <param name="connector">Connector to update.</param>
        public void ApplyTo(VisioConnector connector) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            connector.LineColor = LineColor;
            connector.LineWeight = LineWeight;
            connector.LinePattern = LinePattern;
            if (BeginArrow.HasValue) {
                connector.BeginArrow = BeginArrow.Value;
            }

            if (EndArrow.HasValue) {
                connector.EndArrow = EndArrow.Value;
            }

            if (Kind.HasValue) {
                connector.Kind = Kind.Value;
            }

            if (TextStyle != null) {
                connector.TextStyle = TextStyle.Clone();
            }
        }
    }
}
