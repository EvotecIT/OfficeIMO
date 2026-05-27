using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level builder for cloud, infrastructure, and service architecture diagrams.
    /// </summary>
    public sealed class VisioArchitectureDiagramBuilder {
        private sealed class ComponentItem {
            public ComponentItem(string id, string text, int column, int row, VisioArchitectureShapeKind kind) {
                Id = id;
                Text = text;
                Column = column;
                Row = row;
                Kind = kind;
            }

            public string Id { get; }

            public string Text { get; }

            public int Column { get; }

            public int Row { get; }

            public VisioArchitectureShapeKind Kind { get; }

            public VisioShape? Shape { get; set; }
        }

        private sealed class RegionItem {
            public RegionItem(string id, string text, int column, int row, int columnSpan, int rowSpan) {
                Id = id;
                Text = text;
                Column = column;
                Row = row;
                ColumnSpan = columnSpan;
                RowSpan = rowSpan;
            }

            public string Id { get; }

            public string Text { get; }

            public int Column { get; }

            public int Row { get; }

            public int ColumnSpan { get; }

            public int RowSpan { get; }
        }

        private sealed class LinkItem {
            public LinkItem(string fromId, string toId, VisioArchitectureConnectorKind kind, string? label) {
                FromId = fromId;
                ToId = toId;
                Kind = kind;
                Label = label;
            }

            public string FromId { get; }

            public string ToId { get; }

            public VisioArchitectureConnectorKind Kind { get; }

            public string? Label { get; }
        }

        private sealed class CalloutItem {
            public CalloutItem(string targetId, string id, string text, double pinX, double pinY, VisioCalloutOptions options) {
                TargetId = targetId;
                Id = id;
                Text = text;
                PinX = pinX;
                PinY = pinY;
                Options = options;
            }

            public CalloutItem(string targetId, string id, string text, VisioSide placement, double gap, VisioCalloutOptions options) {
                TargetId = targetId;
                Id = id;
                Text = text;
                Placement = placement;
                Gap = gap;
                Options = options;
                UsePlacement = true;
            }

            public string TargetId { get; }

            public string Id { get; }

            public string Text { get; }

            public double PinX { get; }

            public double PinY { get; }

            public VisioSide Placement { get; }

            public double Gap { get; }

            public bool UsePlacement { get; }

            public VisioCalloutOptions Options { get; }
        }

        private readonly VisioDocument _document;
        private readonly string _pageName;
        private readonly List<ComponentItem> _components = new();
        private readonly Dictionary<string, ComponentItem> _componentsById = new(StringComparer.Ordinal);
        private readonly List<RegionItem> _regions = new();
        private readonly List<LinkItem> _links = new();
        private readonly List<CalloutItem> _callouts = new();
        private VisioStyleTheme _theme = VisioStyleTheme.Technical();
        private VisioMeasurementUnit _unit = VisioMeasurementUnit.Inches;
        private double _pageWidth = 14;
        private double _pageHeight = 8.5;
        private double _leftMargin = 0.85;
        private double _topMargin = 0.85;
        private double _columnGap = 0.65;
        private double _rowGap = 0.75;
        private double _componentWidth = 1.75;
        private double _componentHeight = 0.95;
        private string? _titleText;
        private string _titleId = "title";
        private double _titleHeight = 0.45;
        private double _titleGap = 0.35;
        private bool _showLegend;
        private string _dataFlowLegendLabel = "Data Flow";
        private string _controlFlowLegendLabel = "Control Flow";
        private double _legendHeight = 0.28;
        private double _legendGap = 0.35;
        private bool _built;

        internal VisioArchitectureDiagramBuilder(VisioDocument document, string pageName) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _pageName = string.IsNullOrWhiteSpace(pageName) ? "Architecture Diagram" : pageName;
        }

        /// <summary>Sets the page size used by the generated architecture page.</summary>
        public VisioArchitectureDiagramBuilder PageSize(double width, double height, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _pageWidth = width;
            _pageHeight = height;
            _unit = unit;
            return this;
        }

        /// <summary>Sets the visual theme.</summary>
        public VisioArchitectureDiagramBuilder Theme(VisioStyleTheme theme) {
            _theme = (theme ?? throw new ArgumentNullException(nameof(theme))).Clone();
            return this;
        }

        /// <summary>Sets outer page margins used by the grid layout.</summary>
        public VisioArchitectureDiagramBuilder Margins(double left, double top) {
            ValidateNonNegative(left, nameof(left));
            ValidateNonNegative(top, nameof(top));
            _leftMargin = left;
            _topMargin = top;
            return this;
        }

        /// <summary>Sets grid spacing between components.</summary>
        public VisioArchitectureDiagramBuilder Spacing(double columnGap, double rowGap) {
            ValidateNonNegative(columnGap, nameof(columnGap));
            ValidateNonNegative(rowGap, nameof(rowGap));
            _columnGap = columnGap;
            _rowGap = rowGap;
            return this;
        }

        /// <summary>Sets the default component size.</summary>
        public VisioArchitectureDiagramBuilder ComponentSize(double width, double height) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _componentWidth = width;
            _componentHeight = height;
            return this;
        }

        /// <summary>Adds a centered editable title above the generated architecture diagram.</summary>
        public VisioArchitectureDiagramBuilder Title(string? text = null, string id = "title", double height = 0.45, double gap = 0.35) {
            string normalizedId = RequireId(id, nameof(id), "Title id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            ValidatePositive(height, nameof(height));
            ValidateNonNegative(gap, nameof(gap));
            _titleText = string.IsNullOrWhiteSpace(text) ? _pageName : text;
            _titleId = normalizedId;
            _titleHeight = height;
            _titleGap = gap;
            return this;
        }

        /// <summary>Adds a compact data/control flow legend above the generated architecture grid.</summary>
        public VisioArchitectureDiagramBuilder Legend(bool enabled = true, string dataFlowLabel = "Data Flow", string controlFlowLabel = "Control Flow") {
            _showLegend = enabled;
            _dataFlowLegendLabel = string.IsNullOrWhiteSpace(dataFlowLabel) ? "Data Flow" : dataFlowLabel;
            _controlFlowLegendLabel = string.IsNullOrWhiteSpace(controlFlowLabel) ? "Control Flow" : controlFlowLabel;
            return this;
        }

        /// <summary>Adds a light background region around a grid area.</summary>
        public VisioArchitectureDiagramBuilder Region(string id, string text, int column, int row, int columnSpan, int rowSpan) {
            string normalizedId = RequireId(id, nameof(id), "Region id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            ValidateGridPosition(column, row);
            if (columnSpan <= 0) throw new ArgumentOutOfRangeException(nameof(columnSpan), "Column span must be positive.");
            if (rowSpan <= 0) throw new ArgumentOutOfRangeException(nameof(rowSpan), "Row span must be positive.");
            _regions.Add(new RegionItem(normalizedId, text ?? string.Empty, column, row, columnSpan, rowSpan));
            return this;
        }

        /// <summary>Adds a component at a deterministic grid position.</summary>
        public VisioArchitectureDiagramBuilder Component(string id, string text, int column, int row, VisioArchitectureShapeKind kind = VisioArchitectureShapeKind.Service) {
            string normalizedId = RequireId(id, nameof(id), "Component id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            ValidateGridPosition(column, row);
            if (!Enum.IsDefined(typeof(VisioArchitectureShapeKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            ComponentItem item = new(normalizedId, text ?? string.Empty, column, row, kind);
            _components.Add(item);
            _componentsById.Add(normalizedId, item);
            return this;
        }

        /// <summary>Adds an actor component.</summary>
        public VisioArchitectureDiagramBuilder Actor(string id, string text, int column, int row) =>
            Component(id, text, column, row, VisioArchitectureShapeKind.Actor);

        /// <summary>Adds a service component.</summary>
        public VisioArchitectureDiagramBuilder Service(string id, string text, int column, int row) =>
            Component(id, text, column, row, VisioArchitectureShapeKind.Service);

        /// <summary>Adds a compute component.</summary>
        public VisioArchitectureDiagramBuilder Compute(string id, string text, int column, int row) =>
            Component(id, text, column, row, VisioArchitectureShapeKind.Compute);

        /// <summary>Adds a gateway component.</summary>
        public VisioArchitectureDiagramBuilder Gateway(string id, string text, int column, int row) =>
            Component(id, text, column, row, VisioArchitectureShapeKind.Gateway);

        /// <summary>Adds a database component.</summary>
        public VisioArchitectureDiagramBuilder Database(string id, string text, int column, int row) =>
            Component(id, text, column, row, VisioArchitectureShapeKind.Database);

        /// <summary>Adds a storage component.</summary>
        public VisioArchitectureDiagramBuilder Storage(string id, string text, int column, int row) =>
            Component(id, text, column, row, VisioArchitectureShapeKind.Storage);

        /// <summary>Adds a queue or broker component.</summary>
        public VisioArchitectureDiagramBuilder Queue(string id, string text, int column, int row) =>
            Component(id, text, column, row, VisioArchitectureShapeKind.Queue);

        /// <summary>Adds a security or identity component.</summary>
        public VisioArchitectureDiagramBuilder Security(string id, string text, int column, int row) =>
            Component(id, text, column, row, VisioArchitectureShapeKind.Security);

        /// <summary>Adds a network component.</summary>
        public VisioArchitectureDiagramBuilder Network(string id, string text, int column, int row) =>
            Component(id, text, column, row, VisioArchitectureShapeKind.Network);

        /// <summary>Adds a primary data/request flow.</summary>
        public VisioArchitectureDiagramBuilder DataFlow(string fromId, string toId, string? label = null) =>
            Link(fromId, toId, VisioArchitectureConnectorKind.Data, label);

        /// <summary>Adds a management or orchestration flow.</summary>
        public VisioArchitectureDiagramBuilder ControlFlow(string fromId, string toId, string? label = null) =>
            Link(fromId, toId, VisioArchitectureConnectorKind.Control, label);

        /// <summary>Adds a dependency relationship.</summary>
        public VisioArchitectureDiagramBuilder Dependency(string fromId, string toId, string? label = null) =>
            Link(fromId, toId, VisioArchitectureConnectorKind.Dependency, label);

        /// <summary>Adds a connector between two known components.</summary>
        public VisioArchitectureDiagramBuilder Link(string fromId, string toId, VisioArchitectureConnectorKind kind, string? label = null) {
            EnsureKnownComponent(fromId, nameof(fromId));
            EnsureKnownComponent(toId, nameof(toId));
            if (!Enum.IsDefined(typeof(VisioArchitectureConnectorKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            _links.Add(new LinkItem(fromId, toId, kind, label));
            return this;
        }

        /// <summary>Adds a semantic callout connected to a known component using a generated callout id.</summary>
        public VisioArchitectureDiagramBuilder Callout(string targetId, string text, double pinX, double pinY, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            EnsureKnownComponent(normalizedTargetId, nameof(targetId));
            return Callout(normalizedTargetId, CreateCalloutId(normalizedTargetId), text, pinX, pinY, configure);
        }

        /// <summary>Adds a semantic callout connected to a known component.</summary>
        public VisioArchitectureDiagramBuilder Callout(string targetId, string id, string text, double pinX, double pinY, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            string normalizedId = RequireId(id, nameof(id), "Callout id");
            EnsureKnownComponent(normalizedTargetId, nameof(targetId));
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            ValidateFinite(pinX, nameof(pinX));
            ValidateFinite(pinY, nameof(pinY));
            VisioCalloutOptions options = CreateCalloutOptions();
            configure?.Invoke(options);
            ValidatePositive(options.Width, nameof(options.Width));
            ValidatePositive(options.Height, nameof(options.Height));
            _callouts.Add(new CalloutItem(normalizedTargetId, normalizedId, text ?? string.Empty, pinX, pinY, options));
            return this;
        }

        /// <summary>Adds a semantic callout placed beside a known component using a generated callout id.</summary>
        public VisioArchitectureDiagramBuilder Callout(string targetId, string text, VisioSide placement, double gap = 0.35D, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            EnsureKnownComponent(normalizedTargetId, nameof(targetId));
            return Callout(normalizedTargetId, CreateCalloutId(normalizedTargetId), text, placement, gap, configure);
        }

        /// <summary>Adds a semantic callout placed beside a known component.</summary>
        public VisioArchitectureDiagramBuilder Callout(string targetId, string id, string text, VisioSide placement, double gap = 0.35D, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            string normalizedId = RequireId(id, nameof(id), "Callout id");
            EnsureKnownComponent(normalizedTargetId, nameof(targetId));
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            ValidatePlacement(placement, nameof(placement));
            ValidateNonNegative(gap, nameof(gap));
            VisioCalloutOptions options = CreateCalloutOptions();
            configure?.Invoke(options);
            ValidatePositive(options.Width, nameof(options.Width));
            ValidatePositive(options.Height, nameof(options.Height));
            _callouts.Add(new CalloutItem(normalizedTargetId, normalizedId, text ?? string.Empty, placement, gap, options));
            return this;
        }

        internal VisioPage Build() {
            if (_built) {
                throw new InvalidOperationException("This architecture diagram builder has already produced a page.");
            }

            _built = true;
            if (_components.Count == 0) {
                throw new InvalidOperationException("An architecture diagram requires at least one component.");
            }

            VisioPage page = _document.AddPage(_pageName, _pageWidth, _pageHeight, _unit);
            page.Grid(visible: false, snap: true);
            AddRegions(page);
            AddComponents(page);
            AddLinks(page);
            AddCallouts(page);
            AddAdornments(page);
            _document.RequestRecalcOnOpen();
            return page;
        }

        private void AddAdornments(VisioPage page) {
            if (!string.IsNullOrWhiteSpace(_titleText)) {
                double titleY = _pageHeight - _topMargin - (_titleHeight / 2D);
                double width = Math.Max(1D, _pageWidth - 1.6D);
                VisioShape title = page.AddTextBox(_titleId, _pageWidth / 2D, titleY, width, _titleHeight, _titleText, _unit);
                title.TextStyle = CreateTitleTextStyle();
            }

            if (_showLegend) {
                double titleOffset = string.IsNullOrWhiteSpace(_titleText) ? 0D : _titleHeight + _titleGap;
                double legendY = _pageHeight - _topMargin - titleOffset - (_legendHeight / 2D);
                AddLegendItem(page, Math.Max(0.8D, _leftMargin), legendY, _dataFlowLegendLabel, _theme.DataConnector);
                AddLegendItem(page, Math.Max(0.8D, _pageWidth - 3.35D), legendY, _controlFlowLegendLabel, _theme.ControlConnector);
            }
        }

        private VisioTextStyle CreateTitleTextStyle() {
            VisioTextStyle style = _theme.Container.TextStyle?.Clone() ?? new VisioTextStyle();
            style.FontFamily = string.IsNullOrWhiteSpace(style.FontFamily) ? "Aptos Display" : style.FontFamily;
            style.Size = Math.Max(style.Size ?? 0D, 20D);
            style.Bold = true;
            style.HorizontalAlignment = VisioTextHorizontalAlignment.Center;
            style.VerticalAlignment = VisioTextVerticalAlignment.Middle;
            return style;
        }

        private void AddLegendItem(VisioPage page, double x, double y, string label, VisioConnectorStyle connectorStyle) {
            VisioShape sample = page.AddRectangle(x + 0.32D, y, 0.64D, 0.08D, string.Empty, _unit);
            sample.NameU = "Rectangle";
            sample.FillPattern = 0;
            sample.LineColor = connectorStyle.LineColor;
            sample.LinePattern = connectorStyle.LinePattern;
            sample.LineWeight = Math.Max(0.018D, connectorStyle.LineWeight);
            sample.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.DiagramAdornmentKind, "STR", prompt: "OfficeIMO semantic kind");

            VisioShape text = page.AddTextBox(x + 1.55D, y, 1.65D, _legendHeight, label, _unit);
            text.TextStyle = CreateLegendTextStyle(connectorStyle);
        }

        private VisioTextStyle CreateLegendTextStyle(VisioConnectorStyle connectorStyle) {
            VisioTextStyle style = connectorStyle.TextStyle?.Clone() ?? _theme.DataConnector.TextStyle?.Clone() ?? new VisioTextStyle();
            style.FontFamily = string.IsNullOrWhiteSpace(style.FontFamily) ? "Aptos" : style.FontFamily;
            style.Size = Math.Max(style.Size ?? 0D, 9D);
            style.HorizontalAlignment = VisioTextHorizontalAlignment.Left;
            style.VerticalAlignment = VisioTextVerticalAlignment.Middle;
            return style;
        }

        private void AddRegions(VisioPage page) {
            foreach (RegionItem region in _regions) {
                double width = (region.ColumnSpan * _componentWidth) + ((region.ColumnSpan - 1) * _columnGap) + 0.75;
                double height = (region.RowSpan * _componentHeight) + ((region.RowSpan - 1) * _rowGap) + 0.6;
                VisioShape shape = new(region.Id, GridX(region.Column, region.ColumnSpan), GridY(region.Row, region.RowSpan), width, height, region.Text) {
                    NameU = "Rectangle",
                    Master = _document.EnsureBuiltinMaster("Rectangle")
                };
                _theme.Container.ApplyTo(shape);
                shape.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.BackgroundSurfaceKind, "STR", prompt: "OfficeIMO semantic kind");
                page.Shapes.Add(shape);
            }
        }

        private void AddComponents(VisioPage page) {
            foreach (ComponentItem component in _components) {
                GetComponentShape(component.Kind, out string masterNameU, out double width, out double height);
                VisioShape shape = new(component.Id, GridX(component.Column, 1), GridY(component.Row, 1), width, height, component.Text) {
                    NameU = masterNameU,
                    Master = _document.EnsureBuiltinMaster(masterNameU)
                };
                GetComponentStyle(component.Kind).ApplyTo(shape);
                component.Shape = shape;
                page.Shapes.Add(shape);
            }
        }

        private void AddLinks(VisioPage page) {
            int routeIndex = 0;
            foreach (LinkItem link in _links) {
                ComponentItem from = _componentsById[link.FromId];
                ComponentItem to = _componentsById[link.ToId];
                if (from.Shape == null || to.Shape == null) {
                    throw new InvalidOperationException("Components must be placed before connectors are created.");
                }

                ResolveSides(from.Shape, to.Shape, out VisioSide fromSide, out VisioSide toSide);
                VisioConnector connector = page.AddConnector(from.Shape, to.Shape, ConnectorKind.RightAngle, fromSide, toSide);
                GetConnectorStyle(link.Kind).ApplyTo(connector);
                connector.Label = link.Label;
                connector.RouteOrthogonal(offset: (routeIndex % 3) * 0.08);
                if (!string.IsNullOrWhiteSpace(link.Label)) {
                    connector.PlaceLabel(0.5, offsetY: 0.16);
                }

                routeIndex++;
            }
        }

        private void AddCallouts(VisioPage page) {
            foreach (CalloutItem callout in _callouts) {
                ComponentItem target = _componentsById[callout.TargetId];
                if (target.Shape == null) {
                    throw new InvalidOperationException("Components must be placed before callouts are created.");
                }

                if (callout.UsePlacement) {
                    page.AddCallout(target.Shape, callout.Id, callout.Text, callout.Placement, callout.Gap, callout.Options);
                } else {
                    page.AddCallout(target.Shape, callout.Id, callout.Text, callout.PinX, callout.PinY, callout.Options);
                }
            }
        }

        private VisioCalloutOptions CreateCalloutOptions() {
            return new VisioCalloutOptions {
                ShapeStyle = _theme.Container.Clone(),
                LeaderStyle = new VisioConnectorStyle(_theme.Connector.LineColor, Math.Max(0.012D, _theme.Connector.LineWeight), 2, EndArrow.None) {
                    Kind = ConnectorKind.RightAngle,
                    TextStyle = _theme.Connector.TextStyle?.Clone()
                },
                RouteOffset = 0.08D
            };
        }

        private void GetComponentShape(VisioArchitectureShapeKind kind, out string masterNameU, out double width, out double height) {
            width = _componentWidth;
            height = _componentHeight;
            switch (kind) {
                case VisioArchitectureShapeKind.Actor:
                    masterNameU = "Circle";
                    width = 0.85;
                    height = 0.85;
                    break;
                case VisioArchitectureShapeKind.Database:
                case VisioArchitectureShapeKind.Storage:
                case VisioArchitectureShapeKind.Queue:
                    masterNameU = "Data";
                    break;
                case VisioArchitectureShapeKind.Gateway:
                case VisioArchitectureShapeKind.Security:
                    masterNameU = "Decision";
                    width = _componentWidth * 0.95;
                    height = _componentHeight * 1.15;
                    break;
                case VisioArchitectureShapeKind.Network:
                    masterNameU = "Rectangle";
                    width = _componentWidth * 1.1;
                    break;
                default:
                    masterNameU = "Process";
                    break;
            }
        }

        private VisioShapeStyle GetComponentStyle(VisioArchitectureShapeKind kind) {
            switch (kind) {
                case VisioArchitectureShapeKind.Actor:
                    return _theme.Marker;
                case VisioArchitectureShapeKind.Database:
                case VisioArchitectureShapeKind.Storage:
                    return _theme.Success;
                case VisioArchitectureShapeKind.Queue:
                    return _theme.Decision;
                case VisioArchitectureShapeKind.Gateway:
                case VisioArchitectureShapeKind.Security:
                    return _theme.Emphasis;
                case VisioArchitectureShapeKind.Network:
                case VisioArchitectureShapeKind.External:
                    return _theme.Container;
                default:
                    return _theme.Primary;
            }
        }

        private VisioConnectorStyle GetConnectorStyle(VisioArchitectureConnectorKind kind) {
            switch (kind) {
                case VisioArchitectureConnectorKind.Control:
                    return _theme.ControlConnector;
                case VisioArchitectureConnectorKind.Dependency:
                    return _theme.Connector;
                default:
                    return _theme.DataConnector;
            }
        }

        private double GridX(int column, int span) {
            double left = _leftMargin + column * (_componentWidth + _columnGap);
            double width = span * _componentWidth + (span - 1) * _columnGap;
            return left + width / 2D;
        }

        private double GridY(int row, int span) {
            double top = _pageHeight - _topMargin - HeaderHeight - row * (_componentHeight + _rowGap);
            double height = span * _componentHeight + (span - 1) * _rowGap;
            return top - height / 2D;
        }

        private double HeaderHeight {
            get {
                double height = 0D;
                if (!string.IsNullOrWhiteSpace(_titleText)) {
                    height += _titleHeight + _titleGap;
                }

                if (_showLegend) {
                    height += _legendHeight + _legendGap;
                }

                return height;
            }
        }

        private void EnsureKnownComponent(string id, string parameterName) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Component id cannot be null or whitespace.", parameterName);
            }

            if (!_componentsById.ContainsKey(id)) {
                throw new ArgumentException($"Unknown architecture component id '{id}'.", parameterName);
            }
        }

        private static string RequireId(string id, string parameterName, string label) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException(label + " cannot be null or whitespace.", parameterName);
            }

            return id;
        }

        private bool IsIdInUse(string id) {
            if (!string.IsNullOrWhiteSpace(_titleText) && string.Equals(_titleId, id, StringComparison.Ordinal)) {
                return true;
            }

            if (_componentsById.ContainsKey(id)) {
                return true;
            }

            foreach (RegionItem region in _regions) {
                if (string.Equals(region.Id, id, StringComparison.Ordinal)) {
                    return true;
                }
            }

            foreach (CalloutItem callout in _callouts) {
                if (string.Equals(callout.Id, id, StringComparison.Ordinal)) {
                    return true;
                }
            }

            return false;
        }

        private string CreateCalloutId(string targetId) {
            string id = targetId + "-callout";
            if (!IsIdInUse(id)) {
                return id;
            }

            int index = 2;
            while (IsIdInUse(id + "-" + index)) {
                index++;
            }

            return id + "-" + index;
        }

        private static void ValidateGridPosition(int column, int row) {
            if (column < 0) throw new ArgumentOutOfRangeException(nameof(column), "Column must be zero or greater.");
            if (row < 0) throw new ArgumentOutOfRangeException(nameof(row), "Row must be zero or greater.");
        }

        private static void ValidatePositive(double value, string parameterName) {
            if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0) {
                throw new ArgumentOutOfRangeException(parameterName, "Value must be a finite positive number.");
            }
        }

        private static void ValidateNonNegative(double value, string parameterName) {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 0) {
                throw new ArgumentOutOfRangeException(parameterName, "Value must be a finite non-negative number.");
            }
        }

        private static void ValidateFinite(double value, string parameterName) {
            if (double.IsNaN(value) || double.IsInfinity(value)) {
                throw new ArgumentOutOfRangeException(parameterName, "Value must be a finite number.");
            }
        }

        private static void ValidatePlacement(VisioSide placement, string parameterName) {
            if (placement == VisioSide.Auto || !Enum.IsDefined(typeof(VisioSide), placement)) {
                throw new ArgumentOutOfRangeException(parameterName, "Placement must be Left, Right, Bottom, or Top.");
            }
        }

        private static void ResolveSides(VisioShape from, VisioShape to, out VisioSide fromSide, out VisioSide toSide) {
            double dx = to.PinX - from.PinX;
            double dy = to.PinY - from.PinY;
            if (Math.Abs(dx) > Math.Abs(dy)) {
                if (dx >= 0) {
                    fromSide = VisioSide.Right;
                    toSide = VisioSide.Left;
                } else {
                    fromSide = VisioSide.Left;
                    toSide = VisioSide.Right;
                }
            } else if (dy >= 0) {
                fromSide = VisioSide.Top;
                toSide = VisioSide.Bottom;
            } else {
                fromSide = VisioSide.Bottom;
                toSide = VisioSide.Top;
            }
        }
    }
}
