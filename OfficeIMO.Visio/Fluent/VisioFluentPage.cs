using System;
using System.Collections.Generic;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Visio.Fluent {
    /// <summary>
    /// Fluent wrapper for a single Visio page. Authoring operations forward to the object model.
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

        /// <summary>Sets the physical page size.</summary>
        public VisioFluentPage Size(OfficeLength w, OfficeLength h) { _page.Size(w, h); return this; }

        /// <summary>Sets all print margins.</summary>
        public VisioFluentPage SetMargins(OfficeLength margin) { _page.SetMargins(margin); return this; }

        /// <summary>Sets horizontal and vertical print margins.</summary>
        public VisioFluentPage SetMargins(OfficeLength horizontal, OfficeLength vertical) { _page.SetMargins(horizontal, vertical); return this; }

        /// <summary>Sets individual print margins.</summary>
        public VisioFluentPage SetMargins(OfficeLength left, OfficeLength right, OfficeLength top, OfficeLength bottom) {
            _page.SetMargins(left, right, top, bottom); return this;
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
        public VisioFluentPage AddLayer(string name, string? nameU = null) {
            _page.AddLayer(name, nameU);
            return this;
        }

        /// <summary>Adds a rectangle through the page model.</summary>
        public VisioFluentPage AddRectangle(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddRectangle(x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a rectangle through the page model.</summary>
        public VisioFluentPage AddRectangle(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddRectangle(id, x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a process through the page model.</summary>
        public VisioFluentPage AddProcess(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddProcess(x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a process through the page model.</summary>
        public VisioFluentPage AddProcess(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddProcess(id, x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a square through the page model.</summary>
        public VisioFluentPage AddSquare(OfficeLength x, OfficeLength y, OfficeLength size, string? text = null) {
            RegisterShape(_page.AddSquare(x, y, size, text));
            return this;
        }

        /// <summary>Adds a square through the page model.</summary>
        public VisioFluentPage AddSquare(string id, OfficeLength x, OfficeLength y, OfficeLength size, string? text = null) {
            RegisterShape(_page.AddSquare(id, x, y, size, text));
            return this;
        }

        /// <summary>Adds a circle through the page model.</summary>
        public VisioFluentPage AddCircle(OfficeLength x, OfficeLength y, OfficeLength diameter, string? text = null) {
            RegisterShape(_page.AddCircle(x, y, diameter, text));
            return this;
        }

        /// <summary>Adds a circle through the page model.</summary>
        public VisioFluentPage AddCircle(string id, OfficeLength x, OfficeLength y, OfficeLength diameter, string? text = null) {
            RegisterShape(_page.AddCircle(id, x, y, diameter, text));
            return this;
        }

        /// <summary>Adds a ellipse through the page model.</summary>
        public VisioFluentPage AddEllipse(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddEllipse(x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a ellipse through the page model.</summary>
        public VisioFluentPage AddEllipse(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddEllipse(id, x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a diamond through the page model.</summary>
        public VisioFluentPage AddDiamond(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddDiamond(x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a diamond through the page model.</summary>
        public VisioFluentPage AddDiamond(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddDiamond(id, x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a decision through the page model.</summary>
        public VisioFluentPage AddDecision(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddDecision(x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a decision through the page model.</summary>
        public VisioFluentPage AddDecision(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddDecision(id, x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a data through the page model.</summary>
        public VisioFluentPage AddData(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddData(x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a data through the page model.</summary>
        public VisioFluentPage AddData(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddData(id, x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a preparation through the page model.</summary>
        public VisioFluentPage AddPreparation(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddPreparation(x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a preparation through the page model.</summary>
        public VisioFluentPage AddPreparation(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddPreparation(id, x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a parallelogram through the page model.</summary>
        public VisioFluentPage AddParallelogram(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddParallelogram(x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a parallelogram through the page model.</summary>
        public VisioFluentPage AddParallelogram(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddParallelogram(id, x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a hexagon through the page model.</summary>
        public VisioFluentPage AddHexagon(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddHexagon(x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a hexagon through the page model.</summary>
        public VisioFluentPage AddHexagon(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddHexagon(id, x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a trapezoid through the page model.</summary>
        public VisioFluentPage AddTrapezoid(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddTrapezoid(x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a trapezoid through the page model.</summary>
        public VisioFluentPage AddTrapezoid(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddTrapezoid(id, x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a pentagon through the page model.</summary>
        public VisioFluentPage AddPentagon(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddPentagon(x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a pentagon through the page model.</summary>
        public VisioFluentPage AddPentagon(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddPentagon(id, x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a manualoperation through the page model.</summary>
        public VisioFluentPage AddManualOperation(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddManualOperation(x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a manualoperation through the page model.</summary>
        public VisioFluentPage AddManualOperation(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddManualOperation(id, x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a offpagereference through the page model.</summary>
        public VisioFluentPage AddOffPageReference(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddOffPageReference(x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a offpagereference through the page model.</summary>
        public VisioFluentPage AddOffPageReference(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddOffPageReference(id, x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a triangle through the page model.</summary>
        public VisioFluentPage AddTriangle(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddTriangle(x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a triangle through the page model.</summary>
        public VisioFluentPage AddTriangle(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddTriangle(id, x, y, width, height, text));
            return this;
        }

        /// <summary>Adds an editable text box without a visible border or fill.</summary>
        public VisioFluentPage AddTextBox(string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null, Action<VisioFluentShape>? configure = null) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Shape id cannot be null or whitespace.", nameof(id));
            }

            if (_byId.ContainsKey(id)) {
                throw new ArgumentException($"A shape with id '{id}' already exists on page '{_page.Name}'.", nameof(id));
            }

            VisioShape shape = _page.AddTextBox(id, x, y, width, height, text);
            RegisterShape(shape);
            configure?.Invoke(new VisioFluentShape(shape));
            return this;
        }

        /// <summary>Adds an automatically identified text box through the page model.</summary>
        public VisioFluentPage AddTextBox(OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height,
            string? text = null, Action<VisioFluentShape>? configure = null) {
            VisioShape shape = _page.AddTextBox(x, y, width, height, text);
            RegisterShape(shape);
            configure?.Invoke(new VisioFluentShape(shape));
            return this;
        }

        /// <summary>Adds an instance of an explicitly supplied master.</summary>
        public VisioFluentPage AddShape(string id, VisioMaster master, OfficeLength x, OfficeLength y,
            OfficeLength w, OfficeLength h, string? text = null) {
            RegisterShape(_page.AddShape(id, master, x, y, w, h, text));
            return this;
        }

        /// <summary>Adds a centered title through the page model.</summary>
        public VisioFluentPage AddTitle(string text, string id = "title", OfficeLength? height = null,
            OfficeLength? topMargin = null, Action<VisioFluentShape>? configure = null) {
            VisioShape shape = _page.AddTitle(text, id, height, topMargin);
            RegisterShape(shape);
            configure?.Invoke(new VisioFluentShape(shape));
            return this;
        }

        /// <summary>Adds a shape using a document-registered master.</summary>
        public VisioFluentPage AddShape(string id, string masterNameU, OfficeLength x, OfficeLength y, OfficeLength w, OfficeLength h, string? text = null) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Shape id cannot be null or whitespace.", nameof(id));
            }

            if (_byId.ContainsKey(id)) {
                throw new ArgumentException($"A shape with id '{id}' already exists on page '{_page.Name}'.", nameof(id));
            }

            var shape = _page.AddShape(id, masterNameU, x, y, w, h, text);
            RegisterShape(shape);
            return this;
        }

        /// <summary>Adds a stencil through the public page model.</summary>
        public VisioFluentPage AddStencilShape(VisioStencilShape stencil, string id, OfficeLength x, OfficeLength y, string? text = null) {
            RegisterShape(_page.AddStencilShape(stencil, id, x, y, text));
            return this;
        }

        /// <summary>Adds a stencil through the public page model.</summary>
        public VisioFluentPage AddStencilShape(VisioStencilShape stencil, string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddStencilShape(stencil, id, x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a stencil through the public page model.</summary>
        public VisioFluentPage AddStencilShape(VisioStencilCatalog catalog, string stencilIdOrName, string id, OfficeLength x, OfficeLength y, string? text = null) {
            RegisterShape(_page.AddStencilShape(catalog, stencilIdOrName, id, x, y, text));
            return this;
        }

        /// <summary>Adds a stencil through the public page model.</summary>
        public VisioFluentPage AddStencilShape(VisioStencilCatalog catalog, string stencilIdOrName, string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddStencilShape(catalog, stencilIdOrName, id, x, y, width, height, text));
            return this;
        }

        /// <summary>Adds a stencil through the public page model.</summary>
        public VisioFluentPage AddStencilShape(string stencilIdOrName, string id, OfficeLength x, OfficeLength y, string? text = null) {
            RegisterShape(_page.AddStencilShape(stencilIdOrName, id, x, y, text));
            return this;
        }

        /// <summary>Adds a stencil through the public page model.</summary>
        public VisioFluentPage AddStencilShape(string stencilIdOrName, string id, OfficeLength x, OfficeLength y, OfficeLength width, OfficeLength height, string? text = null) {
            RegisterShape(_page.AddStencilShape(stencilIdOrName, id, x, y, width, height, text));
            return this;
        }

        /// <summary>Configures an existing shape selected by identifier.</summary>
        public VisioFluentPage GetShape(string id, Action<VisioFluentShape> configure) {
            ArgumentNullException.ThrowIfNull(configure);
            configure(new VisioFluentShape(_page.GetShape(id)));
            return this;
        }

        /// <summary>Adds a container around existing shapes selected by identifier.</summary>
        public VisioFluentPage AddContainer(string id, string text, IEnumerable<string> memberIds, VisioContainerOptions? options = null) {
            RegisterShape(_page.AddContainer(id, text, memberIds, options));
            return this;
        }

        /// <summary>Connects existing shapes selected by identifier, with the model's routing defaults.</summary>
        public VisioFluentPage AddConnector(string fromId, string toId, ConnectorKind kind = ConnectorKind.Dynamic,
            VisioSide fromSide = VisioSide.Auto, VisioSide toSide = VisioSide.Auto, Action<VisioFluentConnector>? configure = null) {
            VisioConnector connector = _page.AddConnector(fromId, toId, kind, fromSide, toSide);
            configure?.Invoke(new VisioFluentConnector(connector));
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
