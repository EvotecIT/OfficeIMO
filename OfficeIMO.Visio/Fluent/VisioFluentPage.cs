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
        private readonly Dictionary<string, VisioShape> _byId = new(StringComparer.Ordinal);

        /// <summary>Initializes a new fluent page wrapper.</summary>
        /// <param name="fluent">Parent fluent document.</param>
        /// <param name="page">Underlying page model.</param>
        internal VisioFluentPage(VisioFluentDocument fluent, VisioPage page) {
            _fluent = fluent;
            _page = page;
            foreach (var s in page.Shapes) _byId[s.Id] = s;
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
            var shape = new VisioShape(id, x, y, width, height, text ?? string.Empty) { NameU = "Rectangle" };
            _page.Shapes.Add(shape);
            _byId[id] = shape;
            return this;
        }

        /// <summary>Adds a square shape (width = height = size).</summary>
        public VisioFluentPage Square(string id, double x, double y, double size, string? text = null) {
            var shape = new VisioShape(id, x, y, size, size, text ?? string.Empty) { NameU = "Square" };
            _page.Shapes.Add(shape);
            _byId[id] = shape;
            return this;
        }

        /// <summary>Adds an ellipse shape with explicit width/height.</summary>
        public VisioFluentPage Ellipse(string id, double x, double y, double width, double height, string? text = null) {
            var shape = new VisioShape(id, x, y, width, height, text ?? string.Empty) { NameU = "Ellipse" };
            _page.Shapes.Add(shape);
            _byId[id] = shape;
            return this;
        }

        /// <summary>Adds a diamond (rhombus) shape.</summary>
        public VisioFluentPage Diamond(string id, double x, double y, double width, double height, string? text = null) {
            var shape = new VisioShape(id, x, y, width, height, text ?? string.Empty) { NameU = "Diamond" };
            _page.Shapes.Add(shape);
            _byId[id] = shape;
            return this;
        }

        /// <summary>Adds a circle by diameter.</summary>
        public VisioFluentPage Circle(string id, double x, double y, double diameter, string? text = null) {
            var shape = new VisioShape(id, x, y, diameter, diameter, text ?? string.Empty) { NameU = "Circle" };
            _page.Shapes.Add(shape);
            _byId[id] = shape;
            return this;
        }

        /// <summary>Adds an isosceles triangle with explicit width and height.</summary>
        public VisioFluentPage Triangle(string id, double x, double y, double width, double height, string? text = null) {
            var shape = new VisioShape(id, x, y, width, height, text ?? string.Empty) { NameU = "Triangle" };
            _page.Shapes.Add(shape);
            _byId[id] = shape;
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

        /// <summary>Returns to the document-level fluent builder for chaining.</summary>
        public VisioFluentDocument EndPage() => _fluent;
    }
}
