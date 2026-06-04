using System;
using System.Collections.Generic;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio.Fluent {
    /// <summary>
    /// Fluent builder for a single Visio page. Provides direct verbs like
    /// Rect/Ellipse/Diamond/Triangle/Connect, consistent with other OfficeIMO fluent APIs.
    /// </summary>
    public partial class VisioFluentPage {
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

        internal VisioPage Page => _page;

        internal void RebuildShapeIndex() {
            _byId.Clear();
            foreach (VisioShape shape in _page.Shapes) {
                RegisterShape(shape);
            }
        }

        /// <summary>Sets page size.</summary>
        /// <param name="width">Width value in the specified unit.</param>
        /// <param name="height">Height value in the specified unit.</param>
        /// <param name="unit">Measurement unit (defaults to inches).</param>
        public VisioFluentPage Size(double width, double height, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            _page.Size(width, height, unit);
            return this;
        }

        /// <summary>Sets all print margins.</summary>
        /// <param name="margin">Margin value.</param>
        /// <param name="unit">Measurement unit.</param>
        public VisioFluentPage Margins(double margin, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            _page.SetMargins(margin, unit);
            return this;
        }

        /// <summary>Sets horizontal and vertical print margins.</summary>
        /// <param name="horizontal">Left and right margin value.</param>
        /// <param name="vertical">Top and bottom margin value.</param>
        /// <param name="unit">Measurement unit.</param>
        public VisioFluentPage Margins(double horizontal, double vertical, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            _page.SetMargins(horizontal, vertical, unit);
            return this;
        }

        /// <summary>Sets individual print margins.</summary>
        public VisioFluentPage Margins(double left, double right, double top, double bottom, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            _page.SetMargins(left, right, top, bottom, unit);
            return this;
        }

        /// <summary>Sets print orientation.</summary>
        /// <param name="orientation">Print orientation.</param>
        public VisioFluentPage PrintOrientation(VisioPagePrintOrientation orientation) {
            _page.PrintOrientation = orientation;
            return this;
        }

        /// <summary>Locks or unlocks page replacement.</summary>
        /// <param name="locked">Whether replacement is locked.</param>
        public VisioFluentPage LockReplace(bool locked = true) {
            _page.PageLockReplace = locked;
            return this;
        }

        /// <summary>Locks or unlocks page duplication.</summary>
        /// <param name="locked">Whether duplication is locked.</param>
        public VisioFluentPage LockDuplicate(bool locked = true) {
            _page.PageLockDuplicate = locked;
            return this;
        }

        /// <summary>Sets how Visio determines the drawing page size.</summary>
        /// <param name="sizeType">Drawing size behavior.</param>
        public VisioFluentPage DrawingSize(VisioDrawingSizeType sizeType) {
            _page.DrawingSizeType = sizeType;
            return this;
        }

        /// <summary>Enables or disables automatic page resizing to fit the diagram.</summary>
        /// <param name="enabled">Whether automatic drawing resize is enabled.</param>
        public VisioFluentPage AutoResizeDrawing(bool enabled = true) {
            _page.AutoResizeDrawing = enabled;
            return this;
        }

        /// <summary>Enables or disables automatic shape splitting on this page.</summary>
        /// <param name="enabled">Whether automatic shape splitting is enabled.</param>
        public VisioFluentPage ShapeSplitting(bool enabled = true) {
            _page.AllowShapeSplitting = enabled;
            return this;
        }

        /// <summary>Shows or hides the page name in Visio UI page lists.</summary>
        /// <param name="visibility">Page UI visibility.</param>
        public VisioFluentPage UiVisibility(VisioPageUiVisibility visibility) {
            _page.UiVisibility = visibility;
            return this;
        }

        /// <summary>Sets Visio's page-level placement style.</summary>
        /// <param name="style">Placement style.</param>
        public VisioFluentPage PlacementStyle(VisioPlacementStyle style) {
            _page.PlacementStyle = style;
            return this;
        }

        /// <summary>Sets Visio's page-level placement analysis depth.</summary>
        /// <param name="depth">Placement depth.</param>
        public VisioFluentPage PlacementDepth(VisioPlacementDepth depth) {
            _page.PlacementDepth = depth;
            return this;
        }

        /// <summary>Sets Visio's page-level placement flip behavior.</summary>
        /// <param name="flip">Placement flip flags.</param>
        public VisioFluentPage PlacementFlip(VisioPlacementFlip flip) {
            _page.PlacementFlip = flip;
            return this;
        }

        /// <summary>Enables or disables moving nearby shapes away on drop.</summary>
        /// <param name="enabled">Whether nearby placeable shapes move away on drop.</param>
        public VisioFluentPage MoveShapesAwayOnDrop(bool enabled = true) {
            _page.MoveShapesAwayOnDrop = enabled;
            return this;
        }

        /// <summary>Enables or disables enlarging the page after Visio lays out shapes.</summary>
        /// <param name="enabled">Whether Visio should resize the page after layout.</param>
        public VisioFluentPage ResizePageToFitLayout(bool enabled = true) {
            _page.ResizePageToFitLayout = enabled;
            return this;
        }

        /// <summary>Enables or disables Visio's internal layout grid for page layout.</summary>
        /// <param name="enabled">Whether Visio should use the internal layout grid.</param>
        public VisioFluentPage EnableLayoutGrid(bool enabled = true) {
            _page.EnableLayoutGrid = enabled;
            return this;
        }

        /// <summary>Sets Visio layout grid block size and spacing.</summary>
        /// <param name="blockSize">Horizontal and vertical average shape block size.</param>
        /// <param name="avenueSize">Horizontal and vertical spacing between shapes.</param>
        /// <param name="unit">Measurement unit.</param>
        public VisioFluentPage LayoutGridSizing(double blockSize, double avenueSize, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            _page.SetLayoutGridSizing(blockSize, avenueSize, unit);
            return this;
        }

        /// <summary>Sets individual Visio layout grid block sizes and spacing values.</summary>
        public VisioFluentPage LayoutGridSizing(
            double blockSizeX,
            double blockSizeY,
            double avenueSizeX,
            double avenueSizeY,
            VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            _page.SetLayoutGridSizing(blockSizeX, blockSizeY, avenueSizeX, avenueSizeY, unit);
            return this;
        }

        /// <summary>Clears Visio layout grid sizing cells.</summary>
        public VisioFluentPage ClearLayoutGridSizing() {
            _page.ClearLayoutGridSizing();
            return this;
        }

        /// <summary>Clears Visio layout grid enablement and sizing cells.</summary>
        public VisioFluentPage ClearLayoutGridPolicy() {
            _page.ClearLayoutGridPolicy();
            return this;
        }

        /// <summary>Clears page-level placement policy cells.</summary>
        public VisioFluentPage ClearPlacementPolicy() {
            _page.ClearPlacementPolicy();
            return this;
        }

        /// <summary>Sets Visio's page-level connector routing style.</summary>
        /// <param name="style">Routing style to use for connectors without local routing.</param>
        public VisioFluentPage ConnectorRouteStyle(VisioPageRouteStyle style) {
            _page.ConnectorRouteStyle = style;
            return this;
        }

        /// <summary>Sets the default appearance for routed connectors on this page.</summary>
        /// <param name="appearance">Default routed connector appearance.</param>
        public VisioFluentPage ConnectorRouteAppearance(VisioLineRouteExtension appearance) {
            _page.ConnectorRouteAppearance = appearance;
            return this;
        }

        /// <summary>Sets the default line jump behavior for this page.</summary>
        /// <param name="style">Line jump style.</param>
        /// <param name="code">Which connectors receive line jumps.</param>
        /// <param name="horizontalDirection">Default jump direction for horizontal dynamic connectors.</param>
        /// <param name="verticalDirection">Default jump direction for vertical dynamic connectors.</param>
        public VisioFluentPage LineJumps(
            VisioLineJumpStyle style,
            VisioLineJumpCode code,
            VisioHorizontalLineJumpDirection horizontalDirection = VisioHorizontalLineJumpDirection.Default,
            VisioVerticalLineJumpDirection verticalDirection = VisioVerticalLineJumpDirection.Default) {
            _page.LineJumpStyle = style;
            _page.LineJumpCode = code;
            _page.HorizontalLineJumpDirection = horizontalDirection;
            _page.VerticalLineJumpDirection = verticalDirection;
            return this;
        }

        /// <summary>Sets connector-to-connector and connector-to-shape routing clearances.</summary>
        /// <param name="lineToLine">Horizontal and vertical connector-to-connector clearance.</param>
        /// <param name="lineToNode">Horizontal and vertical connector-to-shape clearance.</param>
        /// <param name="unit">Measurement unit.</param>
        public VisioFluentPage ConnectorSpacing(double lineToLine, double lineToNode, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            _page.SetConnectorSpacing(lineToLine, lineToNode, unit);
            return this;
        }

        /// <summary>Sets individual connector routing clearances.</summary>
        public VisioFluentPage ConnectorSpacing(
            double lineToLineX,
            double lineToLineY,
            double lineToNodeX,
            double lineToNodeY,
            VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            _page.SetConnectorSpacing(lineToLineX, lineToLineY, lineToNodeX, lineToNodeY, unit);
            return this;
        }

        /// <summary>Clears page-level connector spacing cells.</summary>
        public VisioFluentPage ClearConnectorSpacing() {
            _page.ClearConnectorSpacing();
            return this;
        }

        /// <summary>Clears page-level connector routing and line-jump policy cells.</summary>
        public VisioFluentPage ClearConnectorRoutingPolicy() {
            _page.ClearConnectorRoutingPolicy();
            return this;
        }

        /// <summary>Adds a page layer.</summary>
        /// <param name="name">Layer display name.</param>
        /// <param name="nameU">Optional universal name.</param>
        public VisioFluentPage Layer(string name, string? nameU = null) {
            _page.AddLayer(name, nameU);
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

        /// <summary>Adds an editable text box without a visible border or fill.</summary>
        public VisioFluentPage TextBox(string id, double x, double y, double width, double height, string? text = null, Action<VisioFluentShape>? configure = null) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Shape id cannot be null or whitespace.", nameof(id));
            }

            if (_byId.ContainsKey(id)) {
                throw new ArgumentException($"A shape with id '{id}' already exists on page '{_page.Name}'.", nameof(id));
            }

            VisioShape shape = _page.AddTextBox(id, x, y, width, height, text, _page.DefaultUnit);
            RegisterShape(shape);
            configure?.Invoke(new VisioFluentShape(shape));
            return this;
        }

        /// <summary>Adds a centered page title near the top of the page.</summary>
        public VisioFluentPage Title(string text, string id = "title", double height = 0.5D, double topMargin = 0.35D, Action<VisioFluentShape>? configure = null) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Shape id cannot be null or whitespace.", nameof(id));
            }

            if (_byId.ContainsKey(id)) {
                throw new ArgumentException($"A shape with id '{id}' already exists on page '{_page.Name}'.", nameof(id));
            }

            if (double.IsNaN(height) || double.IsInfinity(height) || height <= 0D) {
                throw new ArgumentOutOfRangeException(nameof(height), "Height must be a finite positive number.");
            }

            if (double.IsNaN(topMargin) || double.IsInfinity(topMargin) || topMargin < 0D) {
                throw new ArgumentOutOfRangeException(nameof(topMargin), "Top margin must be a finite non-negative number.");
            }

            double y = _page.Height - topMargin - (height / 2D);
            double width = Math.Max(1D, _page.Width - (topMargin * 2D));
            VisioShape shape = _page.AddTextBox(id, _page.Width / 2D, y, width, height, text, VisioMeasurementUnit.Inches);
            shape.TextStyle = new VisioTextStyle {
                FontFamily = "Aptos Display",
                Color = Color.FromRgb(0, 73, 108),
                Size = 22,
                Bold = true,
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };
            RegisterShape(shape);
            configure?.Invoke(new VisioFluentShape(shape));
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

        /// <summary>Adds a shape from an OfficeIMO-native stencil definition using its default size.</summary>
        public VisioFluentPage Stencil(string id, VisioStencilShape stencil, double x, double y, string? text = null) {
            if (_byId.ContainsKey(id)) {
                throw new ArgumentException($"A shape with id '{id}' already exists on page '{_page.Name}'.", nameof(id));
            }

            VisioShape shape = _page.AddStencilShape(stencil, id, x, y, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a shape from an OfficeIMO-native stencil definition using an explicit size.</summary>
        public VisioFluentPage Stencil(string id, VisioStencilShape stencil, double x, double y, double width, double height, string? text = null) {
            if (_byId.ContainsKey(id)) {
                throw new ArgumentException($"A shape with id '{id}' already exists on page '{_page.Name}'.", nameof(id));
            }

            VisioShape shape = _page.AddStencilShape(stencil, id, x, y, width, height, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a shape from an OfficeIMO-native stencil catalog using its default size.</summary>
        public VisioFluentPage Stencil(string id, VisioStencilCatalog catalog, string stencilIdOrName, double x, double y, string? text = null) {
            if (_byId.ContainsKey(id)) {
                throw new ArgumentException($"A shape with id '{id}' already exists on page '{_page.Name}'.", nameof(id));
            }

            VisioShape shape = _page.AddStencilShape(catalog, stencilIdOrName, id, x, y, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a shape from the combined built-in stencil catalog using its default size.</summary>
        public VisioFluentPage Stencil(string id, string stencilIdOrName, double x, double y, string? text = null) {
            if (_byId.ContainsKey(id)) {
                throw new ArgumentException($"A shape with id '{id}' already exists on page '{_page.Name}'.", nameof(id));
            }

            VisioShape shape = _page.AddStencilShape(stencilIdOrName, id, x, y, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a shape from the combined built-in stencil catalog using an explicit size.</summary>
        public VisioFluentPage Stencil(string id, string stencilIdOrName, double x, double y, double width, double height, string? text = null) {
            if (_byId.ContainsKey(id)) {
                throw new ArgumentException($"A shape with id '{id}' already exists on page '{_page.Name}'.", nameof(id));
            }

            VisioShape shape = _page.AddStencilShape(stencilIdOrName, id, x, y, width, height, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Configures an existing shape (text, stroke, fill, etc.).</summary>
        public VisioFluentPage Shape(string id, Action<VisioFluentShape> configure) {
            if (!_byId.TryGetValue(id, out var shape)) throw new ArgumentException($"Unknown shape id '{id}'", nameof(id));
            configure?.Invoke(new VisioFluentShape(shape));
            return this;
        }

        /// <summary>Adds a Visio-native container around existing shapes.</summary>
        /// <param name="id">Container shape id.</param>
        /// <param name="text">Container heading text.</param>
        /// <param name="memberIds">Existing shape ids to include in the container.</param>
        /// <param name="options">Optional container layout and style settings.</param>
        public VisioFluentPage Container(string id, string text, IEnumerable<string> memberIds, VisioContainerOptions? options = null) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Container id cannot be null or whitespace.", nameof(id));
            }

            if (_byId.ContainsKey(id)) {
                throw new ArgumentException($"A shape with id '{id}' already exists on page '{_page.Name}'.", nameof(id));
            }

            if (memberIds == null) {
                throw new ArgumentNullException(nameof(memberIds));
            }

            List<VisioShape> members = new();
            foreach (string memberId in memberIds) {
                if (!_byId.TryGetValue(memberId, out VisioShape? member)) {
                    throw new ArgumentException($"Unknown shape id '{memberId}'.", nameof(memberIds));
                }

                members.Add(member);
            }

            VisioShape container = _page.AddContainer(id, text, members, options);
            RegisterShape(container);
            return this;
        }

        /// <summary>Connects two shapes by id and optionally configures the connector.</summary>
        public VisioFluentPage Connect(string fromId, string toId, Action<VisioFluentConnector>? configure = null) {
            if (!_byId.TryGetValue(fromId, out var from)) throw new ArgumentException($"Unknown shape id '{fromId}'", nameof(fromId));
            if (!_byId.TryGetValue(toId, out var to)) throw new ArgumentException($"Unknown shape id '{toId}'", nameof(toId));
            var conn = _page.AddConnector(from, to);
            configure?.Invoke(new VisioFluentConnector(conn));
            return this;
        }

        /// <summary>Connects two shapes by id and preselects connector sides.</summary>
        public VisioFluentPage Connect(string fromId, string toId, VisioSide fromSide, VisioSide toSide, Action<VisioFluentConnector>? configure = null) {
            if (!_byId.TryGetValue(fromId, out var from)) throw new ArgumentException($"Unknown shape id '{fromId}'", nameof(fromId));
            if (!_byId.TryGetValue(toId, out var to)) throw new ArgumentException($"Unknown shape id '{toId}'", nameof(toId));
            var conn = _page.AddConnector(from, to, ConnectorKind.Dynamic, fromSide, toSide);
            var builder = new VisioFluentConnector(conn);
            configure?.Invoke(builder);
            return this;
        }

        /// <summary>Applies high-level deterministic cleanup to this page.</summary>
        public VisioFluentPage Polish(VisioDiagramPolishOptions? options = null) {
            _page.PolishDiagram(options);
            return this;
        }

        /// <summary>Moves overlapping top-level shapes apart using deterministic page cleanup.</summary>
        public VisioFluentPage ResolveShapeOverlaps(double step = 0.25D, int maxAttempts = 24, bool includeContainers = false) {
            _page.ResolveShapeOverlaps(step, maxAttempts, includeContainers);
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
