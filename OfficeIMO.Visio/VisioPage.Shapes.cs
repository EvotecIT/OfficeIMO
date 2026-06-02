using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VisioPage {

        private string NextId(VisioConnector? ignoredConnector = null) {
            HashSet<int> usedIds = new();

            void Reserve(string? id) {
                if (int.TryParse(id, out int numericId) && numericId > 0) {
                    usedIds.Add(numericId);
                }
            }

            void VisitShape(VisioShape shape) {
                Reserve(shape.Id);
                foreach (VisioShape child in shape.Children) {
                    VisitShape(child);
                }
            }

            foreach (VisioShape shape in _shapes) {
                VisitShape(shape);
            }

            foreach (VisioConnector connector in _connectors) {
                if (ReferenceEquals(connector, ignoredConnector)) {
                    continue;
                }

                Reserve(connector.Id);
            }

            int nextId = 1;
            while (usedIds.Contains(nextId)) {
                nextId++;
            }

            return nextId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private void PrepareConnectorForPage(VisioConnector connector, VisioConnector? ignoredConnector = null) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            if (_connectors.Contains(connector) && !ReferenceEquals(connector, ignoredConnector)) {
                throw new InvalidOperationException("The connector is already part of this page.");
            }

            if (connector.HasAutomaticId) {
                connector.Id = NextId(ignoredConnector);
            }
        }

        private void PrepareShapeForPage(VisioShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

             if (_shapes.Contains(shape)) {
                throw new InvalidOperationException("The shape is already part of this page.");
            }

            if (shape.Parent != null) {
                throw new InvalidOperationException("A child shape must be removed from its parent before being added to a page.");
            }

            shape.Parent = null;
            shape.NormalizeDescendantParentLinks();
        }

        private static void ApplyUnits(ref double x, ref double y, ref double w, ref double h, VisioMeasurementUnit unit) {
            x = x.ToInches(unit); y = y.ToInches(unit); w = w.ToInches(unit); h = h.ToInches(unit);
        }

        /// <summary>
        /// Adds a rectangle shape.
        /// </summary>
        /// <param name="x">X coordinate of the shape origin.</param>
        /// <param name="y">Y coordinate of the shape origin.</param>
        /// <param name="width">Width of the rectangle.</param>
        /// <param name="height">Height of the rectangle.</param>
        /// <param name="text">Optional text placed on the shape.</param>
        /// <param name="unit">Measurement unit for the provided values.</param>
        /// <returns>The created rectangle shape.</returns>
        public VisioShape AddRectangle(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Rectangle" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a rectangle shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the rectangle.</param>
        /// <param name="height">Height of the rectangle.</param>
        /// <param name="text">Optional text.</param>
        /// <returns>The created rectangle shape.</returns>
        public VisioShape AddRectangle(double x, double y, double width, double height, string? text = null) =>
            AddRectangle(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds an editable text box without a visible border or fill.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the text box.</param>
        /// <param name="height">Height of the text box.</param>
        /// <param name="text">Text to place in the box.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created text box shape.</returns>
        public VisioShape AddTextBox(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            VisioShape shape = AddTextBoxCore(NextId(), x, y, width, height, text, unit);
            Shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// Adds an editable text box with a caller-provided shape id and without a visible border or fill.
        /// </summary>
        /// <param name="id">Shape identifier.</param>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the text box.</param>
        /// <param name="height">Height of the text box.</param>
        /// <param name="text">Text to place in the box.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created text box shape.</returns>
        public VisioShape AddTextBox(string id, double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Shape id cannot be null or whitespace.", nameof(id));
            }

            VisioShape shape = AddTextBoxCore(id, x, y, width, height, text, unit);
            Shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// Adds an editable text box with a caller-provided shape id and without a visible border or fill using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddTextBox(string id, double x, double y, double width, double height, string? text = null) =>
            AddTextBox(id, x, y, width, height, text, DefaultUnit);

        private static VisioShape CreateTextBoxShape(string id, double x, double y, double width, double height, string? text) {
            VisioShape shape = new VisioShape(id, x, y, width, height, text ?? string.Empty);
            shape.NameU = "Text Box";
            shape.LinePattern = 0;
            shape.FillPattern = 0;
            shape.LineColor = OfficeIMO.Drawing.OfficeColor.Transparent;
            shape.FillColor = OfficeIMO.Drawing.OfficeColor.Transparent;
            shape.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.DiagramAdornmentKind, "STR", prompt: "OfficeIMO semantic kind");
            shape.SetUserCell(VisioSemanticUserCells.DiagramAdornmentRole, VisioSemanticUserCells.UserAdornmentRole, "STR", prompt: "OfficeIMO diagram adornment role");
            return shape;
        }

        private static void ValidateTextBoxDimensions(double width, double height) {
            if (double.IsNaN(width) || double.IsInfinity(width) || width <= 0D) {
                throw new ArgumentOutOfRangeException(nameof(width), "Width must be a finite positive number.");
            }

            if (double.IsNaN(height) || double.IsInfinity(height) || height <= 0D) {
                throw new ArgumentOutOfRangeException(nameof(height), "Height must be a finite positive number.");
            }
        }

        private VisioShape AddTextBoxCore(string id, double x, double y, double width, double height, string? text, VisioMeasurementUnit unit) {
            ValidateTextBoxDimensions(width, height);
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            return CreateTextBoxShape(id, x, y, width, height, text);
        }

        /// <summary>
        /// Adds an editable text box without a visible border or fill using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddTextBox(double x, double y, double width, double height, string? text = null) =>
            AddTextBox(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a flowchart process shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the process box.</param>
        /// <param name="height">Height of the process box.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created process shape.</returns>
        public VisioShape AddProcess(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            var s = AddRectangle(x, y, width, height, text, unit);
            s.NameU = "Process";
            return s;
        }

        /// <summary>
        /// Adds a flowchart process shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddProcess(double x, double y, double width, double height, string? text = null) =>
            AddProcess(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a square shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="size">Width and height of the square.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created square shape.</returns>
        public VisioShape AddSquare(double x, double y, double size, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            var s = AddRectangle(x, y, size, size, text, unit);
            s.NameU = "Square";
            return s;
        }

        /// <summary>
        /// Adds a square using the page <see cref="DefaultUnit"/>.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="size">Width and height of the square.</param>
        /// <param name="text">Optional text.</param>
        /// <returns>The created square shape.</returns>
        public VisioShape AddSquare(double x, double y, double size, string? text = null) =>
            AddSquare(x, y, size, text, DefaultUnit);

        /// <summary>
        /// Adds a circle shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="diameter">Diameter of the circle.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created circle shape.</returns>
        public VisioShape AddCircle(double x, double y, double diameter, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            // Avoid double-converting when passing the same variable for width/height
            double w = diameter, h = diameter;
            ApplyUnits(ref x, ref y, ref w, ref h, unit);
            var s = new VisioShape(NextId(), x, y, w, h, text ?? string.Empty) { NameU = "Circle" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a circle using the page <see cref="DefaultUnit"/>.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="diameter">Diameter of the circle.</param>
        /// <param name="text">Optional text.</param>
        /// <returns>The created circle shape.</returns>
        public VisioShape AddCircle(double x, double y, double diameter, string? text = null) =>
            AddCircle(x, y, diameter, text, DefaultUnit);

        /// <summary>
        /// Adds an ellipse shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the ellipse.</param>
        /// <param name="height">Height of the ellipse.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created ellipse shape.</returns>
        public VisioShape AddEllipse(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Ellipse" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds an ellipse using the page <see cref="DefaultUnit"/>.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the ellipse.</param>
        /// <param name="height">Height of the ellipse.</param>
        /// <param name="text">Optional text.</param>
        /// <returns>The created ellipse shape.</returns>
        public VisioShape AddEllipse(double x, double y, double width, double height, string? text = null) =>
            AddEllipse(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a diamond shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the diamond.</param>
        /// <param name="height">Height of the diamond.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created diamond shape.</returns>
        public VisioShape AddDiamond(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Diamond" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a diamond using the page <see cref="DefaultUnit"/>.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the diamond.</param>
        /// <param name="height">Height of the diamond.</param>
        /// <param name="text">Optional text.</param>
        /// <returns>The created diamond shape.</returns>
        public VisioShape AddDiamond(double x, double y, double width, double height, string? text = null) =>
            AddDiamond(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a flowchart decision shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the decision shape.</param>
        /// <param name="height">Height of the decision shape.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created decision shape.</returns>
        public VisioShape AddDecision(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            var s = AddDiamond(x, y, width, height, text, unit);
            s.NameU = "Decision";
            return s;
        }

        /// <summary>
        /// Adds a flowchart decision shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddDecision(double x, double y, double width, double height, string? text = null) =>
            AddDecision(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a flowchart data shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the data shape.</param>
        /// <param name="height">Height of the data shape.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created data shape.</returns>
        public VisioShape AddData(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Data" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a flowchart data shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddData(double x, double y, double width, double height, string? text = null) =>
            AddData(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a flowchart preparation shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the preparation shape.</param>
        /// <param name="height">Height of the preparation shape.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created preparation shape.</returns>
        public VisioShape AddPreparation(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Preparation" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a flowchart preparation shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddPreparation(double x, double y, double width, double height, string? text = null) =>
            AddPreparation(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a parallelogram shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the parallelogram.</param>
        /// <param name="height">Height of the parallelogram.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created parallelogram shape.</returns>
        public VisioShape AddParallelogram(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Parallelogram" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a parallelogram shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddParallelogram(double x, double y, double width, double height, string? text = null) =>
            AddParallelogram(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a hexagon shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the hexagon.</param>
        /// <param name="height">Height of the hexagon.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created hexagon shape.</returns>
        public VisioShape AddHexagon(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Hexagon" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a hexagon shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddHexagon(double x, double y, double width, double height, string? text = null) =>
            AddHexagon(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a trapezoid shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the trapezoid.</param>
        /// <param name="height">Height of the trapezoid.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created trapezoid shape.</returns>
        public VisioShape AddTrapezoid(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Trapezoid" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a trapezoid shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddTrapezoid(double x, double y, double width, double height, string? text = null) =>
            AddTrapezoid(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a pentagon shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the pentagon.</param>
        /// <param name="height">Height of the pentagon.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created pentagon shape.</returns>
        public VisioShape AddPentagon(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Pentagon" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a pentagon shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddPentagon(double x, double y, double width, double height, string? text = null) =>
            AddPentagon(x, y, width, height, text, DefaultUnit);

        /// <summary>
         /// Adds a flowchart manual operation shape.
         /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the manual operation shape.</param>
        /// <param name="height">Height of the manual operation shape.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created manual operation shape.</returns>
        public VisioShape AddManualOperation(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Manual operation" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a flowchart manual operation shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddManualOperation(double x, double y, double width, double height, string? text = null) =>
            AddManualOperation(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a flowchart off-page reference shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the off-page reference shape.</param>
        /// <param name="height">Height of the off-page reference shape.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created off-page reference shape.</returns>
        public VisioShape AddOffPageReference(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Off-page reference" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a flowchart off-page reference shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        public VisioShape AddOffPageReference(double x, double y, double width, double height, string? text = null) =>
            AddOffPageReference(x, y, width, height, text, DefaultUnit);

        /// <summary>
         /// Adds a triangle shape.
         /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the triangle's bounding box.</param>
        /// <param name="height">Height of the triangle's bounding box.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created triangle shape.</returns>
        public VisioShape AddTriangle(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Triangle" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a triangle using the page <see cref="DefaultUnit"/>.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="width">Width of the triangle's bounding box.</param>
        /// <param name="height">Height of the triangle's bounding box.</param>
        /// <param name="text">Optional text.</param>
        /// <returns>The created triangle shape.</returns>
        public VisioShape AddTriangle(double x, double y, double width, double height, string? text = null) =>
            AddTriangle(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a connector between two shapes, optionally specifying side connection points.
        /// </summary>
        /// <param name="from">Source shape.</param>
        /// <param name="to">Target shape.</param>
        /// <param name="kind">Connector kind (straight, curved, etc.).</param>
        /// <param name="fromSide">Preferred side on the source shape.</param>
        /// <param name="toSide">Preferred side on the target shape.</param>
        /// <returns>The created connector.</returns>
        public VisioConnector AddConnector(VisioShape from, VisioShape to, ConnectorKind kind = ConnectorKind.Dynamic, VisioSide fromSide = VisioSide.Auto, VisioSide toSide = VisioSide.Auto) {
            return AddConnectorCore(NextId(), from, to, kind, fromSide, toSide);
        }

        private VisioConnector AddConnectorCore(string id, VisioShape from, VisioShape to, ConnectorKind kind, VisioSide fromSide, VisioSide toSide) {
            var conn = new VisioConnector(id, from, to) { Kind = kind };
            if (fromSide != VisioSide.Auto) conn.FromConnectionPoint = from.EnsureSideConnectionPoint(fromSide);
            if (toSide != VisioSide.Auto) conn.ToConnectionPoint = to.EnsureSideConnectionPoint(toSide);
            Connectors.Add(conn);
            return conn;
        }

        /// <summary>
        /// Sets the page size.
        /// </summary>
        /// <param name="w">Width of the page.</param>
        /// <param name="h">Height of the page.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The current page.</returns>
        public VisioPage Size(double w, double h, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            Width = w.ToInches(unit);
            Height = h.ToInches(unit);
            return this;
        }

        /// <summary>
        /// Configures grid visibility and snapping.
        /// </summary>
        /// <param name="visible">Whether the grid is visible.</param>
        /// <param name="snap">Whether snapping is enabled.</param>
        /// <returns>The current page.</returns>
        public VisioPage Grid(bool visible, bool snap) {
            GridVisible = visible;
            Snap = snap;
            return this;
        }

        /// <summary>
        /// Adds a shape to the page.
        /// </summary>
        /// <param name="id">Identifier of the shape.</param>
        /// <param name="master">Master associated with the shape.</param>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="w">Width of the shape.</param>
        /// <param name="h">Height of the shape.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">
        /// Optional measurement unit. When omitted, values are interpreted using
        /// the page <see cref="DefaultUnit"/>.
        /// </param>
        /// <returns>The created shape.</returns>
        public VisioShape AddShape(string id, VisioMaster master, double x, double y, double w, double h, string? text = null, VisioMeasurementUnit? unit = null) {
            VisioMeasurementUnit effectiveUnit = unit ?? DefaultUnit;
            x = x.ToInches(effectiveUnit);
            y = y.ToInches(effectiveUnit);
            w = w.ToInches(effectiveUnit);
            h = h.ToInches(effectiveUnit);

            VisioShape shape = new VisioShape(id, x, y, w, h, text ?? string.Empty) {
                Master = master,
                NameU = master.NameU
            };
            Shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// Adds a shape using a document-registered master by its NameU.
        /// </summary>
        /// <param name="id">Identifier of the shape.</param>
        /// <param name="masterNameU">Registered master universal name.</param>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="w">Width.</param>
        /// <param name="h">Height.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit for the provided values.</param>
        /// <returns>The created shape.</returns>
        public VisioShape AddShape(string id, string masterNameU, double x, double y, double w, double h, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            if (OwnerDocument == null) {
                throw new InvalidOperationException("This page is not attached to a VisioDocument, so master lookup by name is unavailable.");
            }

            VisioMaster master = OwnerDocument.GetMaster(masterNameU);
            x = x.ToInches(unit);
            y = y.ToInches(unit);
            w = w.ToInches(unit);
            h = h.ToInches(unit);

            VisioShape shape = new VisioShape(id, x, y, w, h, text ?? string.Empty) {
                Master = master,
                NameU = master.NameU
            };
            Shapes.Add(shape);
            return shape;
        }

        /// <summary>
        /// Adds a shape using the page <see cref="DefaultUnit"/> and a document-registered master.
        /// </summary>
        public VisioShape AddShape(string id, string masterNameU, double x, double y, double w, double h, string? text = null) =>
            AddShape(id, masterNameU, x, y, w, h, text, DefaultUnit);
    }
}
