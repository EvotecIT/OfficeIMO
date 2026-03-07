using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio.Fluent {
    /// <summary>
    /// Fluent builder for a single Visio page. Provides direct verbs like
    /// Rect/Ellipse/Diamond/Triangle/Connect, consistent with other OfficeIMO fluent APIs.
    /// </summary>
    public class VisioFluentPage {
        private readonly VisioFluentDocument _fluent;
        private readonly VisioPage _page;
        private readonly Dictionary<string, VisioShape> _byId;

        /// <summary>Initializes a new fluent page wrapper.</summary>
        /// <param name="fluent">Parent fluent document.</param>
        /// <param name="page">Underlying page model.</param>
        internal VisioFluentPage(VisioFluentDocument fluent, VisioPage page) {
            _fluent = fluent;
            _page = page;
            // Build once up-front; shapes added through this fluent API keep it in sync.
            _byId = new Dictionary<string, VisioShape>(Math.Max(4, page.Shapes.Count), StringComparer.Ordinal);
            foreach (var s in page.Shapes) RegisterShape(s);
        }

        /// <summary>Sets page size.</summary>
        /// <param name="width">Width value in the specified unit.</param>
        /// <param name="height">Height value in the specified unit.</param>
        /// <param name="unit">Measurement unit (defaults to inches).</param>
        public VisioFluentPage Size(double width, double height, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            _page.Size(width, height, unit);
            return this;
        }

        /// <summary>Adds a rectangle shape with inline geometry.</summary>
        public VisioFluentPage Rect(string id, double x, double y, double width, double height, string? text = null) {
            var shape = CreateShape(id, "Rectangle", x, y, width, height, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a square shape (width = height = size).</summary>
        public VisioFluentPage Square(string id, double x, double y, double size, string? text = null) {
            var shape = CreateShape(id, "Square", x, y, size, size, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds an ellipse shape with explicit width/height.</summary>
        public VisioFluentPage Ellipse(string id, double x, double y, double width, double height, string? text = null) {
            var shape = CreateShape(id, "Ellipse", x, y, width, height, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a diamond (rhombus) shape.</summary>
        public VisioFluentPage Diamond(string id, double x, double y, double width, double height, string? text = null) {
            var shape = CreateShape(id, "Diamond", x, y, width, height, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a flowchart process shape.</summary>
        public VisioFluentPage Process(string id, double x, double y, double width, double height, string? text = null) {
            var shape = CreateShape(id, "Process", x, y, width, height, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a flowchart decision shape.</summary>
        public VisioFluentPage Decision(string id, double x, double y, double width, double height, string? text = null) {
            var shape = CreateShape(id, "Decision", x, y, width, height, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a flowchart data shape.</summary>
        public VisioFluentPage Data(string id, double x, double y, double width, double height, string? text = null) {
            var shape = CreateShape(id, "Data", x, y, width, height, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a parallelogram shape.</summary>
        public VisioFluentPage Parallelogram(string id, double x, double y, double width, double height, string? text = null) {
            var shape = CreateShape(id, "Parallelogram", x, y, width, height, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a flowchart preparation shape.</summary>
        public VisioFluentPage Preparation(string id, double x, double y, double width, double height, string? text = null) {
            var shape = CreateShape(id, "Preparation", x, y, width, height, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a hexagon shape.</summary>
        public VisioFluentPage Hexagon(string id, double x, double y, double width, double height, string? text = null) {
            var shape = CreateShape(id, "Hexagon", x, y, width, height, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a flowchart manual operation shape.</summary>
        public VisioFluentPage ManualOperation(string id, double x, double y, double width, double height, string? text = null) {
            var shape = CreateShape(id, "Manual operation", x, y, width, height, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a trapezoid shape.</summary>
        public VisioFluentPage Trapezoid(string id, double x, double y, double width, double height, string? text = null) {
            var shape = CreateShape(id, "Trapezoid", x, y, width, height, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a flowchart off-page reference shape.</summary>
        public VisioFluentPage OffPageReference(string id, double x, double y, double width, double height, string? text = null) {
            var shape = CreateShape(id, "Off-page reference", x, y, width, height, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a pentagon shape.</summary>
        public VisioFluentPage Pentagon(string id, double x, double y, double width, double height, string? text = null) {
            var shape = CreateShape(id, "Pentagon", x, y, width, height, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a circle by diameter.</summary>
        public VisioFluentPage Circle(string id, double x, double y, double diameter, string? text = null) {
            var shape = CreateShape(id, "Circle", x, y, diameter, diameter, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds an isosceles triangle with explicit width and height.</summary>
        public VisioFluentPage Triangle(string id, double x, double y, double width, double height, string? text = null) {
            var shape = CreateShape(id, "Triangle", x, y, width, height, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a shape using a document-registered master.</summary>
        public VisioFluentPage Master(string id, string masterNameU, double x, double y, double width, double height, string? text = null) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Shape id cannot be null or whitespace.", nameof(id));
            }

            if (_byId.ContainsKey(id)) {
                throw new ArgumentException($"A shape with id '{id}' already exists on page '{_page.Name}'.", nameof(id));
            }

            var shape = _page.AddShape(id, masterNameU, x, y, width, height, text, _page.DefaultUnit);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Configures an existing shape (text, stroke, fill, etc.).</summary>
        public VisioFluentPage Shape(string id, Action<VisioFluentShape> configure) {
            if (!_byId.TryGetValue(id, out var shape)) throw new ArgumentException($"Unknown shape id '{id}'", nameof(id));
            configure?.Invoke(new VisioFluentShape(shape));
            return this;
        }

        /// <summary>Connects two shapes by id and optionally configures the connector.</summary>
        public VisioFluentPage Connect(string fromId, string toId, Action<VisioFluentConnector>? configure = null) {
            if (!_byId.TryGetValue(fromId, out var from)) throw new ArgumentException($"Unknown shape id '{fromId}'", nameof(fromId));
            if (!_byId.TryGetValue(toId, out var to)) throw new ArgumentException($"Unknown shape id '{toId}'", nameof(toId));
            var conn = new VisioConnector(from, to);
            _page.Connectors.Add(conn);
            configure?.Invoke(new VisioFluentConnector(conn));
            return this;
        }

        /// <summary>Connects two shapes by id and preselects connector sides.</summary>
        public VisioFluentPage Connect(string fromId, string toId, VisioSide fromSide, VisioSide toSide, Action<VisioFluentConnector>? configure = null) {
            if (!_byId.TryGetValue(fromId, out var from)) throw new ArgumentException($"Unknown shape id '{fromId}'", nameof(fromId));
            if (!_byId.TryGetValue(toId, out var to)) throw new ArgumentException($"Unknown shape id '{toId}'", nameof(toId));
            var conn = new VisioConnector(from, to);
            _page.Connectors.Add(conn);
            var builder = new VisioFluentConnector(conn).Sides(fromSide, toSide);
            configure?.Invoke(builder);
            return this;
        }

        /// <summary>Returns to the document-level fluent builder for chaining.</summary>
        public VisioFluentDocument EndPage() => _fluent;

        private VisioShape CreateShape(string id, string nameU, double x, double y, double width, double height, string? text) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Shape id cannot be null or whitespace.", nameof(id));
            }

            if (_byId.ContainsKey(id)) {
                throw new ArgumentException($"A shape with id '{id}' already exists on page '{_page.Name}'.", nameof(id));
            }

            VisioMeasurementUnit unit = _page.DefaultUnit;
            x = x.ToInches(unit);
            y = y.ToInches(unit);
            width = width.ToInches(unit);
            height = height.ToInches(unit);

            var shape = new VisioShape(id, x, y, width, height, text ?? string.Empty) { NameU = nameU };
            _page.Shapes.Add(shape);
            return shape;
        }

        private void RegisterShape(VisioShape shape) {
            if (string.IsNullOrWhiteSpace(shape.Id)) {
                throw new ArgumentException("Shape id cannot be null or whitespace.", nameof(shape));
            }

            if (_byId.ContainsKey(shape.Id)) {
                throw new ArgumentException($"A shape with id '{shape.Id}' already exists on page '{_page.Name}'.", nameof(shape));
            }

            _byId[shape.Id] = shape;
            foreach (VisioShape child in shape.Children) {
                RegisterShape(child);
            }
        }
    }
}
