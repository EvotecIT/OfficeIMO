using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Editable set of Visio connectors returned by query helpers.
    /// </summary>
    public sealed class VisioConnectorSelection : IReadOnlyList<VisioConnector> {
        private readonly IReadOnlyList<VisioConnector> _connectors;

        /// <summary>
        /// Initializes a new connector selection.
        /// </summary>
        /// <param name="connectors">Connectors included in the selection.</param>
        public VisioConnectorSelection(IEnumerable<VisioConnector> connectors) {
            if (connectors == null) {
                throw new ArgumentNullException(nameof(connectors));
            }

            _connectors = connectors.ToList();
        }

        /// <inheritdoc />
        public int Count => _connectors.Count;

        /// <inheritdoc />
        public VisioConnector this[int index] => _connectors[index];

        /// <inheritdoc />
        public IEnumerator<VisioConnector> GetEnumerator() {
            return _connectors.GetEnumerator();
        }

        /// <inheritdoc />
        IEnumerator IEnumerable.GetEnumerator() {
            return GetEnumerator();
        }

        /// <summary>
        /// Sets connector kind on every selected connector.
        /// </summary>
        /// <param name="kind">Connector kind.</param>
        public VisioConnectorSelection Kind(ConnectorKind kind) {
            foreach (VisioConnector connector in _connectors) {
                connector.Kind = kind;
            }

            return this;
        }

        /// <summary>
        /// Sets connector label on every selected connector.
        /// </summary>
        /// <param name="label">Connector label.</param>
        public VisioConnectorSelection Label(string? label) {
            foreach (VisioConnector connector in _connectors) {
                connector.Label = label;
            }

            return this;
        }

        /// <summary>
        /// Places connector labels along every selected connector.
        /// </summary>
        /// <param name="position">Position along each connector path, from 0.0 to 1.0.</param>
        /// <param name="offsetX">Horizontal page-coordinate offset.</param>
        /// <param name="offsetY">Vertical page-coordinate offset.</param>
        /// <param name="width">Label text box width in page units.</param>
        /// <param name="height">Label text box height in page units.</param>
        public VisioConnectorSelection LabelPosition(double position = 0.5D, double offsetX = 0D, double offsetY = 0D, double width = 1.25D, double height = 0.3D) {
            foreach (VisioConnector connector in _connectors) {
                connector.PlaceLabel(position, offsetX, offsetY, width, height);
            }

            return this;
        }

        /// <summary>
        /// Sets line color, weight, and pattern on every selected connector.
        /// </summary>
        /// <param name="color">Line color.</param>
        /// <param name="weight">Line weight in inches.</param>
        /// <param name="pattern">Line pattern index.</param>
        public VisioConnectorSelection Stroke(Color color, double weight = VisioShape.DefaultLineWeight, int pattern = 1) {
            foreach (VisioConnector connector in _connectors) {
                connector.LineColor = color;
                connector.LineWeight = weight;
                connector.LinePattern = pattern;
            }

            return this;
        }

        /// <summary>
        /// Applies a reusable connector style on every selected connector.
        /// </summary>
        /// <param name="style">Connector style to apply.</param>
        public VisioConnectorSelection Style(VisioConnectorStyle style) {
            return this.ApplyStyle(style);
        }

        /// <summary>
        /// Sets line color on every selected connector.
        /// </summary>
        /// <param name="color">Line color.</param>
        public VisioConnectorSelection LineColor(Color color) {
            foreach (VisioConnector connector in _connectors) {
                connector.LineColor = color;
            }

            return this;
        }

        /// <summary>
        /// Sets line weight on every selected connector.
        /// </summary>
        /// <param name="weight">Line weight in inches.</param>
        public VisioConnectorSelection LineWeight(double weight) {
            foreach (VisioConnector connector in _connectors) {
                connector.LineWeight = weight;
            }

            return this;
        }

        /// <summary>
        /// Sets line pattern on every selected connector.
        /// </summary>
        /// <param name="pattern">Line pattern index.</param>
        public VisioConnectorSelection LinePattern(int pattern) {
            foreach (VisioConnector connector in _connectors) {
                connector.LinePattern = pattern;
            }

            return this;
        }

        /// <summary>
        /// Sets begin arrow on every selected connector.
        /// </summary>
        /// <param name="arrow">Begin arrow style.</param>
        public VisioConnectorSelection BeginArrow(EndArrow? arrow) {
            foreach (VisioConnector connector in _connectors) {
                connector.BeginArrow = arrow;
            }

            return this;
        }

        /// <summary>
        /// Sets end arrow on every selected connector.
        /// </summary>
        /// <param name="arrow">End arrow style.</param>
        public VisioConnectorSelection EndArrow(EndArrow? arrow) {
            foreach (VisioConnector connector in _connectors) {
                connector.EndArrow = arrow;
            }

            return this;
        }

        /// <summary>
        /// Adds a hyperlink to every selected connector.
        /// </summary>
        /// <param name="address">External hyperlink address.</param>
        /// <param name="description">Optional display description.</param>
        /// <param name="subAddress">Optional internal sub-address.</param>
        public VisioConnectorSelection Hyperlink(string address, string? description = null, string? subAddress = null) {
            foreach (VisioConnector connector in _connectors) {
                connector.AddHyperlink(address, description, subAddress);
            }

            return this;
        }

        /// <summary>
        /// Applies a reusable Shape Data schema on every selected connector.
        /// </summary>
        /// <param name="schema">Shape Data schema to apply.</param>
        /// <param name="overwriteValues">Whether schema defaults should replace existing values.</param>
        public VisioConnectorSelection ShapeData(VisioShapeDataSchema schema, bool overwriteValues = false) {
            if (schema == null) {
                throw new ArgumentNullException(nameof(schema));
            }

            foreach (VisioConnector connector in _connectors) {
                schema.ApplyTo(connector, overwriteValues);
            }

            return this;
        }

        /// <summary>
        /// Configures protection for every selected connector.
        /// </summary>
        /// <param name="configure">Protection configuration delegate.</param>
        public VisioConnectorSelection Protect(Action<VisioProtection> configure) {
            if (configure == null) {
                throw new ArgumentNullException(nameof(configure));
            }

            foreach (VisioConnector connector in _connectors) {
                configure(connector.Protection);
            }

            return this;
        }

        /// <summary>
        /// Locks or unlocks endpoints for every selected connector.
        /// </summary>
        public VisioConnectorSelection LockEndpoints(bool locked = true) {
            foreach (VisioConnector connector in _connectors) {
                connector.Protection.Endpoints(locked);
            }

            return this;
        }

        /// <summary>
        /// Clears explicit protection settings for every selected connector.
        /// </summary>
        public VisioConnectorSelection ClearProtection() {
            foreach (VisioConnector connector in _connectors) {
                connector.Protection.Clear();
            }

            return this;
        }
    }
}
