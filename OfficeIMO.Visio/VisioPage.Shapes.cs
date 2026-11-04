using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VisioPage {

        private string NextId(VisioConnector? ignoredConnector = null) {
            if (ignoredConnector == null) {
                EnsurePageObjectIdIndex();
                return _nextAutomaticObjectId.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            InvalidatePageObjectIdIndex();
            HashSet<int> usedIds = _automaticObjectIdScratch;
            usedIds.Clear();

            foreach (VisioShape shape in _shapes) {
                ReserveShapeTreeIds(shape, usedIds);
            }

            foreach (VisioConnector connector in _connectors) {
                if (ReferenceEquals(connector, ignoredConnector)) {
                    continue;
                }

                ReserveNumericId(connector.Id, usedIds);
            }

            foreach (string reservedId in _reservedAutomaticObjectIds) {
                ReserveNumericId(reservedId, usedIds);
            }

            int nextId = 1;
            while (usedIds.Contains(nextId)) {
                nextId++;
            }

            return nextId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static void ReserveShapeTreeIds(VisioShape shape, HashSet<int> usedIds) {
            ReserveNumericId(shape.Id, usedIds);
            for (int index = 0; index < shape.Children.Count; index++) {
                ReserveShapeTreeIds(shape.Children[index], usedIds);
            }
        }

        private static void ReserveNumericId(string? id, HashSet<int> usedIds) {
            if (int.TryParse(id, out int numericId) && numericId > 0) {
                usedIds.Add(numericId);
            }
        }

        internal void ReserveAutomaticObjectId(string id) {
            if (!string.IsNullOrWhiteSpace(id)) {
                _reservedAutomaticObjectIds.Add(id);
                if (_pageObjectIdIndexCurrent) {
                    ReserveNumericId(id, _automaticObjectIdScratch);
                    AdvanceNextAutomaticObjectId();
                }
            }
        }

        private void EnsurePageObjectIdIndex() {
            if (_pageObjectIdIndexCurrent) {
                return;
            }

            _pageObjectIds.Clear();
            _automaticObjectIdScratch.Clear();
            _nextAutomaticObjectId = 1;
            _pageObjectIdIndexCurrent = true;
            for (int index = 0; index < _shapes.Count; index++) {
                RegisterShapeTreeIds(_shapes[index]);
            }

            for (int index = 0; index < _connectors.Count; index++) {
                RegisterPageObjectId(_connectors[index].Id);
            }

            foreach (string reservedId in _reservedAutomaticObjectIds) {
                ReserveNumericId(reservedId, _automaticObjectIdScratch);
            }

            AdvanceNextAutomaticObjectId();
        }

        private void AdvanceNextAutomaticObjectId() {
            while (_automaticObjectIdScratch.Contains(_nextAutomaticObjectId)) {
                _nextAutomaticObjectId++;
            }
        }

        private void RegisterPageObjectId(string id) {
            if (!_pageObjectIdIndexCurrent) {
                return;
            }

            _pageObjectIds.Add(id);
            ReserveNumericId(id, _automaticObjectIdScratch);
            AdvanceNextAutomaticObjectId();
        }

        private void RegisterShapeTreeIds(VisioShape shape) {
            RegisterPageObjectId(shape.Id);
            for (int index = 0; index < shape.Children.Count; index++) {
                RegisterShapeTreeIds(shape.Children[index]);
            }
        }

        private void InvalidatePageObjectIdIndex() {
            _pageObjectIdIndexCurrent = false;
        }

        private void AttachShapeTree(VisioShape shape) {
            shape.SetOwnerPage(this);
            RegisterShapeTreeIds(shape);
        }

        private void DetachShapeTree(VisioShape shape) {
            shape.SetOwnerPage(null);
            InvalidatePageObjectIdIndex();
        }

        internal void PrepareNestedShapeForPage(VisioShape shape, VisioShape? ignoredShape = null) {
            EnsureShapeNotOwnedByAnotherPage(shape);
            EnsureShapeTreeIdsAvailable(shape, ignoredShape);
        }

        internal void AttachNestedShapeTree(VisioShape shape) {
            AttachShapeTree(shape);
        }

        internal void DetachNestedShapeTree(VisioShape shape) {
            DetachShapeTree(shape);
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

            EnsurePageObjectIdAvailable(connector.Id, "Connector", ignoredConnector: ignoredConnector);
        }

        private void PrepareShapeForPage(VisioShape shape, VisioShape? ignoredShape = null) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            EnsureShapeNotOwnedByAnotherPage(shape);

            if (_shapes.Contains(shape) && !ReferenceEquals(shape, ignoredShape)) {
                throw new InvalidOperationException("The shape is already part of this page.");
            }

            if (shape.Parent != null) {
                throw new InvalidOperationException("A child shape must be removed from its parent before being added to a page.");
            }

            EnsureShapeTreeIdsAvailable(shape, ignoredShape);
            shape.Parent = null;
            shape.NormalizeDescendantParentLinks();
        }

        private void EnsureShapeNotOwnedByAnotherPage(VisioShape shape) {
            if (shape.OwnerPage != null && !ReferenceEquals(shape.OwnerPage, this)) {
                throw new InvalidOperationException("The shape already belongs to another page. Remove it from that page before reusing it.");
            }
        }

        private void EnsureShapeTreeIdsAvailable(VisioShape shape, VisioShape? ignoredShape) {
            if (shape.Children.Count == 0) {
                if (string.IsNullOrWhiteSpace(shape.Id)) {
                    throw new InvalidOperationException("Shape id cannot be null or whitespace.");
                }
                EnsurePageObjectIdAvailable(shape.Id, "Shape", ignoredShape: ignoredShape);
                return;
            }

            HashSet<string> newIds = new(StringComparer.Ordinal);

            void Visit(VisioShape current) {
                if (string.IsNullOrWhiteSpace(current.Id)) {
                    throw new InvalidOperationException("Shape id cannot be null or whitespace.");
                }

                if (!newIds.Add(current.Id)) {
                    throw new InvalidOperationException($"Duplicate shape id '{current.Id}' found in the shape tree being added to page '{Name}'.");
                }

                EnsurePageObjectIdAvailable(current.Id, "Shape", ignoredShape: ignoredShape);
                for (int index = 0; index < current.Children.Count; index++) {
                    Visit(current.Children[index]);
                }
            }

            Visit(shape);
        }

        private void EnsurePageObjectIdAvailable(string id, string kind, VisioShape? ignoredShape = null, VisioConnector? ignoredConnector = null) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new InvalidOperationException($"{kind} id cannot be null or whitespace on page '{Name}'.");
            }

            if (ignoredShape == null && ignoredConnector == null) {
                EnsurePageObjectIdIndex();
                if (_pageObjectIds.Contains(id)) {
                    throw new InvalidOperationException($"{kind} id '{id}' is already used by another page object on page '{Name}'.");
                }
                return;
            }

            foreach (VisioShape shape in _shapes) {
                if (ShapeTreeContainsId(shape, id, ignoredShape)) {
                    throw new InvalidOperationException($"{kind} id '{id}' is already used by another page object on page '{Name}'.");
                }
            }

            foreach (VisioConnector connector in _connectors) {
                if (!ReferenceEquals(connector, ignoredConnector) &&
                    string.Equals(connector.Id, id, StringComparison.Ordinal)) {
                    throw new InvalidOperationException($"{kind} id '{id}' is already used by another page object on page '{Name}'.");
                }
            }
        }

        private static bool ShapeTreeContainsId(VisioShape shape, string id, VisioShape? ignoredShape) {
            if (ReferenceEquals(shape, ignoredShape)) {
                return false;
            }

            if (string.Equals(shape.Id, id, StringComparison.Ordinal)) {
                return true;
            }

            for (int index = 0; index < shape.Children.Count; index++) {
                if (ShapeTreeContainsId(shape.Children[index], id, ignoredShape)) {
                    return true;
                }
            }

            return false;
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
        /// <param name="id">Explicit identifier, or null to generate the next available identifier.</param>
        internal VisioShape AddRectangleInches(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, string? id = null) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(id ?? NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Rectangle" };
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
        internal VisioShape AddRectangleInches(double x, double y, double width, double height, string? text = null) =>
            AddRectangleInches(x, y, width, height, text, DefaultUnit);

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
        internal VisioShape AddTextBoxInches(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
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
        internal VisioShape AddTextBoxInches(string id, double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
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
        internal VisioShape AddTextBoxInches(string id, double x, double y, double width, double height, string? text = null) =>
            AddTextBoxInches(id, x, y, width, height, text, DefaultUnit);

        private static VisioShape CreateTextBoxShape(string id, double x, double y, double width, double height, string? text) {
            VisioShape shape = new VisioShape(id, x, y, width, height, text ?? string.Empty);
            shape.NameU = "Text Box";
            shape.LinePattern = 0;
            shape.FillPattern = 0;
            shape.LineColor = OfficeIMO.OfficeColor.Transparent;
            shape.FillColor = OfficeIMO.OfficeColor.Transparent;
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
        internal VisioShape AddTextBoxInches(double x, double y, double width, double height, string? text = null) =>
            AddTextBoxInches(x, y, width, height, text, DefaultUnit);

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
        /// <param name="id">Explicit identifier, or null to generate the next available identifier.</param>
        internal VisioShape AddProcessInches(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, string? id = null) {
            var s = AddRectangleInches(x, y, width, height, text, unit, id);
            s.NameU = "Process";
            return s;
        }

        /// <summary>
        /// Adds a flowchart process shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        internal VisioShape AddProcessInches(double x, double y, double width, double height, string? text = null) =>
            AddProcessInches(x, y, width, height, text, DefaultUnit);

        /// <summary>
        /// Adds a square shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="size">Width and height of the square.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created square shape.</returns>
        /// <param name="id">Explicit identifier, or null to generate the next available identifier.</param>
        internal VisioShape AddSquareInches(double x, double y, double size, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, string? id = null) {
            var s = AddRectangleInches(x, y, size, size, text, unit, id);
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
        internal VisioShape AddSquareInches(double x, double y, double size, string? text = null) =>
            AddSquareInches(x, y, size, text, DefaultUnit);

        /// <summary>
        /// Adds a circle shape.
        /// </summary>
        /// <param name="x">X coordinate.</param>
        /// <param name="y">Y coordinate.</param>
        /// <param name="diameter">Diameter of the circle.</param>
        /// <param name="text">Optional text.</param>
        /// <param name="unit">Measurement unit.</param>
        /// <returns>The created circle shape.</returns>
        /// <param name="id">Explicit identifier, or null to generate the next available identifier.</param>
        internal VisioShape AddCircleInches(double x, double y, double diameter, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, string? id = null) {
            // Avoid double-converting when passing the same variable for width/height
            double w = diameter, h = diameter;
            ApplyUnits(ref x, ref y, ref w, ref h, unit);
            var s = new VisioShape(id ?? NextId(), x, y, w, h, text ?? string.Empty) { NameU = "Circle" };
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
        internal VisioShape AddCircleInches(double x, double y, double diameter, string? text = null) =>
            AddCircleInches(x, y, diameter, text, DefaultUnit);

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
        /// <param name="id">Explicit identifier, or null to generate the next available identifier.</param>
        internal VisioShape AddEllipseInches(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, string? id = null) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(id ?? NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Ellipse" };
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
        internal VisioShape AddEllipseInches(double x, double y, double width, double height, string? text = null) =>
            AddEllipseInches(x, y, width, height, text, DefaultUnit);

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
        /// <param name="id">Explicit identifier, or null to generate the next available identifier.</param>
        internal VisioShape AddDiamondInches(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, string? id = null) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(id ?? NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Diamond" };
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
        internal VisioShape AddDiamondInches(double x, double y, double width, double height, string? text = null) =>
            AddDiamondInches(x, y, width, height, text, DefaultUnit);

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
        /// <param name="id">Explicit identifier, or null to generate the next available identifier.</param>
        internal VisioShape AddDecisionInches(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, string? id = null) {
            var s = AddDiamondInches(x, y, width, height, text, unit, id);
            s.NameU = "Decision";
            return s;
        }

        /// <summary>
        /// Adds a flowchart decision shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        internal VisioShape AddDecisionInches(double x, double y, double width, double height, string? text = null) =>
            AddDecisionInches(x, y, width, height, text, DefaultUnit);

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
        /// <param name="id">Explicit identifier, or null to generate the next available identifier.</param>
        internal VisioShape AddDataInches(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, string? id = null) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(id ?? NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Data" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a flowchart data shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        internal VisioShape AddDataInches(double x, double y, double width, double height, string? text = null) =>
            AddDataInches(x, y, width, height, text, DefaultUnit);

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
        /// <param name="id">Explicit identifier, or null to generate the next available identifier.</param>
        internal VisioShape AddPreparationInches(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, string? id = null) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(id ?? NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Preparation" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a flowchart preparation shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        internal VisioShape AddPreparationInches(double x, double y, double width, double height, string? text = null) =>
            AddPreparationInches(x, y, width, height, text, DefaultUnit);

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
        /// <param name="id">Explicit identifier, or null to generate the next available identifier.</param>
        internal VisioShape AddParallelogramInches(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, string? id = null) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(id ?? NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Parallelogram" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a parallelogram shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        internal VisioShape AddParallelogramInches(double x, double y, double width, double height, string? text = null) =>
            AddParallelogramInches(x, y, width, height, text, DefaultUnit);

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
        /// <param name="id">Explicit identifier, or null to generate the next available identifier.</param>
        internal VisioShape AddHexagonInches(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, string? id = null) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(id ?? NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Hexagon" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a hexagon shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        internal VisioShape AddHexagonInches(double x, double y, double width, double height, string? text = null) =>
            AddHexagonInches(x, y, width, height, text, DefaultUnit);

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
        /// <param name="id">Explicit identifier, or null to generate the next available identifier.</param>
        internal VisioShape AddTrapezoidInches(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, string? id = null) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(id ?? NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Trapezoid" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a trapezoid shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        internal VisioShape AddTrapezoidInches(double x, double y, double width, double height, string? text = null) =>
            AddTrapezoidInches(x, y, width, height, text, DefaultUnit);

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
        /// <param name="id">Explicit identifier, or null to generate the next available identifier.</param>
        internal VisioShape AddPentagonInches(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, string? id = null) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(id ?? NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Pentagon" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a pentagon shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        internal VisioShape AddPentagonInches(double x, double y, double width, double height, string? text = null) =>
            AddPentagonInches(x, y, width, height, text, DefaultUnit);

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
        /// <param name="id">Explicit identifier, or null to generate the next available identifier.</param>
        internal VisioShape AddManualOperationInches(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, string? id = null) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(id ?? NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Manual operation" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a flowchart manual operation shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        internal VisioShape AddManualOperationInches(double x, double y, double width, double height, string? text = null) =>
            AddManualOperationInches(x, y, width, height, text, DefaultUnit);

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
        /// <param name="id">Explicit identifier, or null to generate the next available identifier.</param>
        internal VisioShape AddOffPageReferenceInches(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, string? id = null) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(id ?? NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Off-page reference" };
            Shapes.Add(s);
            return s;
        }

        /// <summary>
        /// Adds a flowchart off-page reference shape using the page <see cref="DefaultUnit"/>.
        /// </summary>
        internal VisioShape AddOffPageReferenceInches(double x, double y, double width, double height, string? text = null) =>
            AddOffPageReferenceInches(x, y, width, height, text, DefaultUnit);

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
        /// <param name="id">Explicit identifier, or null to generate the next available identifier.</param>
        internal VisioShape AddTriangleInches(double x, double y, double width, double height, string? text = null, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches, string? id = null) {
            ApplyUnits(ref x, ref y, ref width, ref height, unit);
            var s = new VisioShape(id ?? NextId(), x, y, width, height, text ?? string.Empty) { NameU = "Triangle" };
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
        internal VisioShape AddTriangleInches(double x, double y, double width, double height, string? text = null) =>
            AddTriangleInches(x, y, width, height, text, DefaultUnit);

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
        internal VisioPage SizeInches(double w, double h, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            WidthInches = w.ToInches(unit);
            HeightInches = h.ToInches(unit);
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

    }
}
