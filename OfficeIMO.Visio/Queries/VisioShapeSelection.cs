using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Editable set of Visio shapes returned by query helpers.
    /// </summary>
    public sealed class VisioShapeSelection : IReadOnlyList<VisioShape> {
        private readonly IReadOnlyList<VisioShape> _shapes;

        /// <summary>
        /// Initializes a new shape selection.
        /// </summary>
        /// <param name="shapes">Shapes included in the selection.</param>
        /// <param name="ownerPage">Optional page that owns the selection.</param>
        public VisioShapeSelection(IEnumerable<VisioShape> shapes, VisioPage? ownerPage = null) {
            if (shapes == null) {
                throw new ArgumentNullException(nameof(shapes));
            }

            _shapes = shapes.ToList();
            OwnerPage = ownerPage;
        }

        /// <summary>
        /// Gets the page that produced this selection, when known.
        /// </summary>
        public VisioPage? OwnerPage { get; }

        /// <inheritdoc />
        public int Count => _shapes.Count;

        /// <inheritdoc />
        public VisioShape this[int index] => _shapes[index];

        /// <inheritdoc />
        public IEnumerator<VisioShape> GetEnumerator() {
            return _shapes.GetEnumerator();
        }

        /// <inheritdoc />
        IEnumerator IEnumerable.GetEnumerator() {
            return GetEnumerator();
        }

        /// <summary>
        /// Sets text on every selected shape.
        /// </summary>
        /// <param name="text">Text to assign.</param>
        public VisioShapeSelection Text(string? text) {
            foreach (VisioShape shape in _shapes) {
                shape.Text = text;
            }

            return this;
        }

        /// <summary>
        /// Updates text on every selected shape from the current shape.
        /// </summary>
        /// <param name="textFactory">Factory returning text for each shape.</param>
        public VisioShapeSelection Text(Func<VisioShape, string?> textFactory) {
            if (textFactory == null) {
                throw new ArgumentNullException(nameof(textFactory));
            }

            foreach (VisioShape shape in _shapes) {
                shape.Text = textFactory(shape);
            }

            return this;
        }

        /// <summary>
        /// Sets fill color and optional fill pattern on every selected shape.
        /// </summary>
        /// <param name="color">Fill color.</param>
        /// <param name="pattern">Fill pattern index.</param>
        public VisioShapeSelection Fill(Color color, int pattern = 1) {
            foreach (VisioShape shape in _shapes) {
                shape.FillColor = color;
                shape.FillPattern = pattern;
            }

            return this;
        }

        /// <summary>
        /// Applies a reusable shape style on every selected shape.
        /// </summary>
        /// <param name="style">Shape style to apply.</param>
        public VisioShapeSelection Style(VisioShapeStyle style) {
            return this.ApplyStyle(style);
        }

        /// <summary>
        /// Sets stroke color, weight, and pattern on every selected shape.
        /// </summary>
        /// <param name="color">Line color.</param>
        /// <param name="weight">Line weight in inches.</param>
        /// <param name="pattern">Line pattern index.</param>
        public VisioShapeSelection Stroke(Color color, double weight = VisioShape.DefaultLineWeight, int pattern = 1) {
            foreach (VisioShape shape in _shapes) {
                shape.LineColor = color;
                shape.LineWeight = weight;
                shape.LinePattern = pattern;
            }

            return this;
        }

        /// <summary>
        /// Sets line color on every selected shape.
        /// </summary>
        /// <param name="color">Line color.</param>
        public VisioShapeSelection LineColor(Color color) {
            foreach (VisioShape shape in _shapes) {
                shape.LineColor = color;
            }

            return this;
        }

        /// <summary>
        /// Sets line weight on every selected shape.
        /// </summary>
        /// <param name="weight">Line weight in inches.</param>
        public VisioShapeSelection LineWeight(double weight) {
            foreach (VisioShape shape in _shapes) {
                shape.LineWeight = weight;
            }

            return this;
        }

        /// <summary>
        /// Sets line pattern on every selected shape.
        /// </summary>
        /// <param name="pattern">Line pattern index.</param>
        public VisioShapeSelection LinePattern(int pattern) {
            foreach (VisioShape shape in _shapes) {
                shape.LinePattern = pattern;
            }

            return this;
        }

        /// <summary>
        /// Sets or replaces a data value on every selected shape.
        /// </summary>
        /// <param name="key">Data key.</param>
        /// <param name="value">Data value.</param>
        public VisioShapeSelection Data(string key, string value) {
            if (string.IsNullOrWhiteSpace(key)) {
                throw new ArgumentException("Data key cannot be empty.", nameof(key));
            }

            foreach (VisioShape shape in _shapes) {
                shape.SetShapeData(key, value);
            }

            return this;
        }

        /// <summary>
        /// Sets or replaces a Visio Shape Data row on every selected shape.
        /// </summary>
        /// <param name="name">Shape Data row name.</param>
        /// <param name="value">Shape Data value.</param>
        /// <param name="label">Optional label shown in Visio's Shape Data window.</param>
        /// <param name="type">Optional Shape Data type.</param>
        /// <param name="prompt">Optional help prompt.</param>
        /// <param name="format">Optional format picture or list values.</param>
        public VisioShapeSelection ShapeData(string name, string? value, string? label = null, VisioShapeDataType? type = null, string? prompt = null, string? format = null) {
            foreach (VisioShape shape in _shapes) {
                shape.SetShapeData(name, value, label, type, prompt, format);
            }

            return this;
        }

        /// <summary>
        /// Applies a reusable Shape Data schema on every selected shape.
        /// </summary>
        /// <param name="schema">Shape Data schema to apply.</param>
        /// <param name="overwriteValues">Whether schema defaults should replace existing values.</param>
        public VisioShapeSelection ShapeData(VisioShapeDataSchema schema, bool overwriteValues = false) {
            if (schema == null) {
                throw new ArgumentNullException(nameof(schema));
            }

            foreach (VisioShape shape in _shapes) {
                schema.ApplyTo(shape, overwriteValues);
            }

            return this;
        }

        /// <summary>
        /// Sets or replaces a Visio User cell on every selected shape.
        /// </summary>
        /// <param name="name">User cell row name.</param>
        /// <param name="value">User cell value.</param>
        /// <param name="unit">Optional Visio unit code.</param>
        /// <param name="formula">Optional ShapeSheet formula.</param>
        /// <param name="prompt">Optional prompt value.</param>
        public VisioShapeSelection UserCell(string name, string? value, string? unit = null, string? formula = null, string? prompt = null) {
            foreach (VisioShape shape in _shapes) {
                shape.SetUserCell(name, value, unit, formula, prompt);
            }

            return this;
        }

        /// <summary>
        /// Adds every selected shape to a page layer.
        /// </summary>
        /// <param name="layerName">Layer name.</param>
        public VisioShapeSelection Layer(string layerName) {
            if (string.IsNullOrWhiteSpace(layerName)) {
                throw new ArgumentException("Layer name cannot be empty.", nameof(layerName));
            }

            foreach (VisioShape shape in _shapes) {
                shape.LayerNames.Add(layerName);
            }

            return this;
        }

        /// <summary>
        /// Adds a hyperlink to every selected shape.
        /// </summary>
        /// <param name="address">External hyperlink address.</param>
        /// <param name="description">Optional display description.</param>
        /// <param name="subAddress">Optional internal sub-address.</param>
        public VisioShapeSelection Hyperlink(string address, string? description = null, string? subAddress = null) {
            foreach (VisioShape shape in _shapes) {
                shape.AddHyperlink(address, description, subAddress);
            }

            return this;
        }

        /// <summary>
        /// Configures protection for every selected shape.
        /// </summary>
        /// <param name="configure">Protection configuration delegate.</param>
        public VisioShapeSelection Protect(Action<VisioShapeProtection> configure) {
            if (configure == null) {
                throw new ArgumentNullException(nameof(configure));
            }

            foreach (VisioShape shape in _shapes) {
                configure(shape.Protection);
            }

            return this;
        }

        /// <summary>
        /// Locks or unlocks size for every selected shape.
        /// </summary>
        public VisioShapeSelection LockSize(bool locked = true) {
            foreach (VisioShape shape in _shapes) {
                shape.Protection.Size(locked);
            }

            return this;
        }

        /// <summary>
        /// Locks or unlocks position for every selected shape.
        /// </summary>
        public VisioShapeSelection LockPosition(bool locked = true) {
            foreach (VisioShape shape in _shapes) {
                shape.Protection.Position(locked);
            }

            return this;
        }

        /// <summary>
        /// Clears explicit protection settings for every selected shape.
        /// </summary>
        public VisioShapeSelection ClearProtection() {
            foreach (VisioShape shape in _shapes) {
                shape.Protection.Clear();
            }

            return this;
        }

        /// <summary>
        /// Moves every selected shape by the provided offset.
        /// </summary>
        /// <param name="deltaX">Horizontal offset in inches.</param>
        /// <param name="deltaY">Vertical offset in inches.</param>
        public VisioShapeSelection MoveBy(double deltaX, double deltaY) {
            foreach (VisioShape shape in _shapes) {
                shape.PinX += deltaX;
                shape.PinY += deltaY;
            }

            return this;
        }

        /// <summary>
        /// Resizes every selected shape.
        /// </summary>
        /// <param name="width">New width in inches.</param>
        /// <param name="height">New height in inches.</param>
        public VisioShapeSelection Size(double width, double height) {
            foreach (VisioShape shape in _shapes) {
                shape.Width = width;
                shape.Height = height;
                shape.LocPinX = width / 2;
                shape.LocPinY = height / 2;
            }

            return this;
        }
    }
}
