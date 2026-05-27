using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level builder for generic graph diagrams where OfficeIMO lays out arbitrary nodes and edges.
    /// </summary>
    public sealed class VisioGraphDiagramBuilder {
        private sealed class NodeItem {
            public NodeItem(string id, string text, VisioGraphNodeKind kind, VisioStencilShape? stencil) {
                Id = id;
                Text = text;
                Kind = kind;
                Stencil = stencil;
            }

            public string Id { get; }

            public string Text { get; }

            public VisioGraphNodeKind Kind { get; }

            public VisioStencilShape? Stencil { get; }

            public int Layer { get; set; }

            public int Row { get; set; }

            public double PinX { get; set; }

            public double PinY { get; set; }

            public VisioShape? Shape { get; set; }
        }

        private sealed class EdgeItem {
            public EdgeItem(string fromId, string toId, VisioGraphConnectorKind kind, string? label, bool directed) {
                FromId = fromId;
                ToId = toId;
                Kind = kind;
                Label = label;
                Directed = directed;
            }

            public string FromId { get; }

            public string ToId { get; }

            public VisioGraphConnectorKind Kind { get; }

            public string? Label { get; }

            public bool Directed { get; }
        }

        private sealed class ZoneItem {
            public ZoneItem(string id, string text, IReadOnlyList<string> nodeIds) {
                Id = id;
                Text = text;
                NodeIds = nodeIds;
            }

            public string Id { get; }

            public string Text { get; }

            public IReadOnlyList<string> NodeIds { get; }
        }

        private readonly VisioDocument _document;
        private readonly string _pageName;
        private readonly List<NodeItem> _nodes = new();
        private readonly Dictionary<string, NodeItem> _nodesById = new(StringComparer.Ordinal);
        private readonly List<EdgeItem> _edges = new();
        private readonly List<string> _rootIds = new();
        private readonly HashSet<string> _rootIdSet = new(StringComparer.Ordinal);
        private readonly List<ZoneItem> _zones = new();
        private readonly HashSet<string> _zoneIds = new(StringComparer.Ordinal);
        private readonly HashSet<string> _generatedIds = new(StringComparer.Ordinal);
        private VisioStyleTheme _theme = VisioStyleTheme.Technical();
        private VisioMeasurementUnit _unit = VisioMeasurementUnit.Inches;
        private VisioGraphLayout _layout = VisioGraphLayout.Layered;
        private VisioGraphDirection _direction = VisioGraphDirection.LeftToRight;
        private double _pageWidth = 11;
        private double _pageHeight = 8.5;
        private double _leftMargin = 0.8;
        private double _topMargin = 0.8;
        private double _rightMargin = 0.8;
        private double _bottomMargin = 0.8;
        private double _nodeWidth = 1.65;
        private double _nodeHeight = 0.82;
        private double _columnGap = 1.1;
        private double _rowGap = 0.65;
        private string? _titleText;
        private string _titleId = "title";
        private double _titleHeight = 0.45;
        private double _titleGap = 0.35;
        private int _maximumRows = 1;
        private bool _fitPageToGraph = true;
        private bool _built;

        internal VisioGraphDiagramBuilder(VisioDocument document, string pageName) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _pageName = string.IsNullOrWhiteSpace(pageName) ? "Graph Diagram" : pageName;
        }

        /// <summary>Sets the page size used by the generated graph page.</summary>
        public VisioGraphDiagramBuilder PageSize(double width, double height, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _pageWidth = width;
            _pageHeight = height;
            _unit = unit;
            return this;
        }

        /// <summary>Sets whether the builder can grow the page to fit the graph. Enabled by default.</summary>
        public VisioGraphDiagramBuilder FitPageToGraph(bool enabled = true) {
            _fitPageToGraph = enabled;
            return this;
        }

        /// <summary>Sets the automatic graph layout strategy.</summary>
        public VisioGraphDiagramBuilder Layout(VisioGraphLayout layout) {
            if (!Enum.IsDefined(typeof(VisioGraphLayout), layout)) {
                throw new ArgumentOutOfRangeException(nameof(layout));
            }

            _layout = layout;
            return this;
        }

        /// <summary>Sets the primary flow direction for layered layouts.</summary>
        public VisioGraphDiagramBuilder Direction(VisioGraphDirection direction) {
            if (!Enum.IsDefined(typeof(VisioGraphDirection), direction)) {
                throw new ArgumentOutOfRangeException(nameof(direction));
            }

            _direction = direction;
            return this;
        }

        /// <summary>Sets the visual theme.</summary>
        public VisioGraphDiagramBuilder Theme(VisioStyleTheme theme) {
            _theme = (theme ?? throw new ArgumentNullException(nameof(theme))).Clone();
            return this;
        }

        /// <summary>Sets outer page margins used by automatic layout.</summary>
        public VisioGraphDiagramBuilder Margins(double left, double top, double right = 0.8D, double bottom = 0.8D) {
            ValidateNonNegative(left, nameof(left));
            ValidateNonNegative(top, nameof(top));
            ValidateNonNegative(right, nameof(right));
            ValidateNonNegative(bottom, nameof(bottom));
            _leftMargin = left;
            _topMargin = top;
            _rightMargin = right;
            _bottomMargin = bottom;
            return this;
        }

        /// <summary>Sets the default node size used for native graph nodes.</summary>
        public VisioGraphDiagramBuilder NodeSize(double width, double height) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _nodeWidth = width;
            _nodeHeight = height;
            return this;
        }

        /// <summary>Sets spacing between layers/columns and rows/rings.</summary>
        public VisioGraphDiagramBuilder Spacing(double columnGap, double rowGap) {
            ValidateNonNegative(columnGap, nameof(columnGap));
            ValidateNonNegative(rowGap, nameof(rowGap));
            _columnGap = columnGap;
            _rowGap = rowGap;
            return this;
        }

        /// <summary>Adds a centered editable title above the graph.</summary>
        public VisioGraphDiagramBuilder Title(string? text = null, string id = "title", double height = 0.45D, double gap = 0.35D) {
            string normalizedId = RequireId(id, nameof(id), "Title id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A graph diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            ValidatePositive(height, nameof(height));
            ValidateNonNegative(gap, nameof(gap));
            _titleText = string.IsNullOrWhiteSpace(text) ? _pageName : text;
            _titleId = normalizedId;
            _titleHeight = height;
            _titleGap = gap;
            return this;
        }

        /// <summary>Adds a background zone around graph nodes.</summary>
        public VisioGraphDiagramBuilder Zone(string id, string text, params string[] nodeIds) {
            string normalizedId = RequireId(id, nameof(id), "Zone id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A graph diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            IReadOnlyList<string> normalizedNodeIds = NormalizeZoneNodeIds(nodeIds);
            _zones.Add(new ZoneItem(normalizedId, text ?? string.Empty, normalizedNodeIds));
            _zoneIds.Add(normalizedId);
            return this;
        }

        /// <summary>Adds and marks a root node used by layered and radial layout.</summary>
        public VisioGraphDiagramBuilder Root(string id, string text, VisioGraphNodeKind kind = VisioGraphNodeKind.External) {
            Node(id, text, kind);
            AddRoot(RequireId(id, nameof(id), "Root node id"));
            return this;
        }

        /// <summary>Marks an existing node as a root used by layered and radial layout.</summary>
        public VisioGraphDiagramBuilder Root(string id) {
            string normalizedId = RequireId(id, nameof(id), "Root node id");
            EnsureKnownNode(normalizedId, nameof(id));
            AddRoot(normalizedId);
            return this;
        }

        /// <summary>Adds a native graph node.</summary>
        public VisioGraphDiagramBuilder Node(string id, string text, VisioGraphNodeKind kind = VisioGraphNodeKind.Process) {
            string normalizedId = RequireId(id, nameof(id), "Node id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A graph diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            if (!Enum.IsDefined(typeof(VisioGraphNodeKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            NodeItem node = new(normalizedId, text ?? string.Empty, kind, null);
            _nodes.Add(node);
            _nodesById.Add(normalizedId, node);
            return this;
        }

        /// <summary>Adds a graph node backed by an OfficeIMO, installed Visio, or external package stencil.</summary>
        public VisioGraphDiagramBuilder StencilNode(string id, string text, VisioStencilShape stencil) {
            if (stencil == null) throw new ArgumentNullException(nameof(stencil));
            string normalizedId = RequireId(id, nameof(id), "Node id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A graph diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            NodeItem node = new(normalizedId, text ?? string.Empty, VisioGraphNodeKind.Process, stencil);
            _nodes.Add(node);
            _nodesById.Add(normalizedId, node);
            return this;
        }

        /// <summary>Adds a standard directed graph edge.</summary>
        public VisioGraphDiagramBuilder Edge(string fromId, string toId, string? label = null) =>
            Edge(fromId, toId, VisioGraphConnectorKind.Standard, label, directed: true);

        /// <summary>Adds a data-flow graph edge.</summary>
        public VisioGraphDiagramBuilder DataEdge(string fromId, string toId, string? label = null) =>
            Edge(fromId, toId, VisioGraphConnectorKind.Data, label, directed: true);

        /// <summary>Adds a control-flow graph edge.</summary>
        public VisioGraphDiagramBuilder ControlEdge(string fromId, string toId, string? label = null) =>
            Edge(fromId, toId, VisioGraphConnectorKind.Control, label, directed: true);

        /// <summary>Adds an emphasized graph edge.</summary>
        public VisioGraphDiagramBuilder EmphasisEdge(string fromId, string toId, string? label = null) =>
            Edge(fromId, toId, VisioGraphConnectorKind.Emphasis, label, directed: true);

        /// <summary>Adds an undirected graph edge.</summary>
        public VisioGraphDiagramBuilder Relationship(string fromId, string toId, string? label = null) =>
            Edge(fromId, toId, VisioGraphConnectorKind.Standard, label, directed: false);

        /// <summary>Adds a graph edge between two known nodes.</summary>
        public VisioGraphDiagramBuilder Edge(string fromId, string toId, VisioGraphConnectorKind kind, string? label = null, bool directed = true) {
            string normalizedFromId = RequireId(fromId, nameof(fromId), "From node id");
            string normalizedToId = RequireId(toId, nameof(toId), "To node id");
            EnsureKnownNode(normalizedFromId, nameof(fromId));
            EnsureKnownNode(normalizedToId, nameof(toId));
            if (!Enum.IsDefined(typeof(VisioGraphConnectorKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            _edges.Add(new EdgeItem(normalizedFromId, normalizedToId, kind, label, directed));
            return this;
        }

        internal VisioPage Build() {
            if (_built) {
                throw new InvalidOperationException("This graph diagram builder has already produced a page.");
            }

            _built = true;
            if (_nodes.Count == 0) {
                throw new InvalidOperationException("A graph diagram requires at least one node.");
            }

            ValidateZones();
            AssignLayoutMetadata();
            SizePageForLayout();
            AssignCoordinates();

            VisioPage page = _document.AddPage(_pageName, _pageWidth, _pageHeight, _unit);
            page.Grid(visible: false, snap: true);
            AddZones(page);
            AddNodes(page);
            AddEdges(page);
            AddTitle(page);
            page.PolishDiagram(new VisioDiagramPolishOptions {
                FitToContent = false,
                ResizeShapesToText = false,
                ResizeConnectorLabelsToText = true,
                ResolveConnectorLabelOverlaps = true
            });
            _document.RequestRecalcOnOpen();
            return page;
        }

        private void ValidateZones() {
            foreach (ZoneItem zone in _zones) {
                foreach (string nodeId in zone.NodeIds) {
                    EnsureKnownNode(nodeId, nameof(zone.NodeIds));
                }
            }
        }

        private void AssignLayoutMetadata() {
            if (_layout == VisioGraphLayout.Grid) {
                AssignGridMetadata();
                return;
            }

            AssignBreadthFirstMetadata();
        }

        private void AssignGridMetadata() {
            int columns = Math.Max(1, (int)Math.Ceiling(Math.Sqrt(_nodes.Count)));
            for (int i = 0; i < _nodes.Count; i++) {
                _nodes[i].Layer = i % columns;
                _nodes[i].Row = i / columns;
            }

            _maximumRows = _nodes.GroupBy(node => node.Layer).Max(group => group.Count());
        }

        private void AssignBreadthFirstMetadata() {
            Dictionary<string, List<string>> outgoing = _nodes.ToDictionary(node => node.Id, _ => new List<string>(), StringComparer.Ordinal);
            Dictionary<string, List<string>> undirected = _nodes.ToDictionary(node => node.Id, _ => new List<string>(), StringComparer.Ordinal);
            Dictionary<string, int> indegree = _nodes.ToDictionary(node => node.Id, _ => 0, StringComparer.Ordinal);
            foreach (EdgeItem edge in _edges) {
                outgoing[edge.FromId].Add(edge.ToId);
                undirected[edge.FromId].Add(edge.ToId);
                undirected[edge.ToId].Add(edge.FromId);
                if (edge.Directed) {
                    indegree[edge.ToId]++;
                } else {
                    outgoing[edge.ToId].Add(edge.FromId);
                }
            }

            HashSet<string> assigned = new(StringComparer.Ordinal);
            Queue<NodeItem> ready = new();

            void Enqueue(NodeItem node, int layer) {
                if (assigned.Add(node.Id)) {
                    node.Layer = layer;
                    ready.Enqueue(node);
                }
            }

            if (_rootIds.Count > 0) {
                foreach (string rootId in _rootIds) {
                    Enqueue(_nodesById[rootId], 0);
                }
            } else {
                foreach (NodeItem root in _nodes.Where(node => indegree[node.Id] == 0)) {
                    Enqueue(root, 0);
                }
            }

            if (assigned.Count == 0 && _nodes.Count > 0) {
                Enqueue(_nodes[0], 0);
            }

            while (assigned.Count < _nodes.Count || ready.Count > 0) {
                while (ready.Count > 0) {
                    NodeItem node = ready.Dequeue();
                    IReadOnlyList<string> nextIds = outgoing[node.Id].Count > 0 ? outgoing[node.Id] : undirected[node.Id];
                    foreach (string nextId in nextIds) {
                        if (!assigned.Contains(nextId)) {
                            Enqueue(_nodesById[nextId], node.Layer + 1);
                        }
                    }
                }

                if (assigned.Count < _nodes.Count) {
                    NodeItem nextRoot = _nodes.First(node => !assigned.Contains(node.Id));
                    Enqueue(nextRoot, 0);
                }
            }

            foreach (IGrouping<int, NodeItem> layer in _nodes.GroupBy(node => node.Layer).OrderBy(group => group.Key)) {
                int row = 0;
                foreach (NodeItem node in layer.OrderBy(node => _nodes.IndexOf(node))) {
                    node.Row = row;
                    row++;
                }
            }

            _maximumRows = Math.Max(1, _nodes.GroupBy(node => node.Layer).Max(group => group.Count()));
        }

        private void SizePageForLayout() {
            if (!_fitPageToGraph) {
                return;
            }

            int layerCount = Math.Max(1, _nodes.Max(node => node.Layer) + 1);
            int rowCount = Math.Max(1, _nodes.GroupBy(node => node.Layer).Max(group => group.Count()));
            double requiredWidth;
            double requiredHeight;
            if (_layout == VisioGraphLayout.Radial) {
                double radius = Math.Max(1D, _nodes.Max(node => node.Layer)) * Math.Max(_nodeWidth + _columnGap, _nodeHeight + _rowGap);
                requiredWidth = _leftMargin + _rightMargin + (radius * 2D) + _nodeWidth * 2D;
                requiredHeight = _topMargin + _bottomMargin + HeaderHeight + (radius * 2D) + _nodeHeight * 2D;
            } else if (_direction == VisioGraphDirection.TopToBottom) {
                requiredWidth = _leftMargin + _rightMargin + (rowCount * _nodeWidth) + Math.Max(0, rowCount - 1) * _columnGap;
                requiredHeight = _topMargin + _bottomMargin + HeaderHeight + (layerCount * _nodeHeight) + Math.Max(0, layerCount - 1) * _rowGap;
            } else {
                requiredWidth = _leftMargin + _rightMargin + (layerCount * _nodeWidth) + Math.Max(0, layerCount - 1) * _columnGap;
                requiredHeight = _topMargin + _bottomMargin + HeaderHeight + (rowCount * _nodeHeight) + Math.Max(0, rowCount - 1) * _rowGap;
            }

            _pageWidth = Math.Max(_pageWidth, requiredWidth);
            _pageHeight = Math.Max(_pageHeight, requiredHeight);
        }

        private void AssignCoordinates() {
            if (_layout == VisioGraphLayout.Radial) {
                AssignRadialCoordinates();
                return;
            }

            foreach (NodeItem node in _nodes) {
                if (_direction == VisioGraphDirection.TopToBottom) {
                    node.PinX = XForRow(node.Row);
                    node.PinY = YForLayer(node.Layer);
                } else {
                    node.PinX = XForLayer(node.Layer);
                    node.PinY = YForRow(node.Row);
                }
            }
        }

        private void AssignRadialCoordinates() {
            double contentHeight = _pageHeight - _topMargin - _bottomMargin - HeaderHeight;
            double centerX = _leftMargin + ((_pageWidth - _leftMargin - _rightMargin) / 2D);
            double centerY = _bottomMargin + (contentHeight / 2D);
            double ringGap = Math.Max(_nodeWidth + _columnGap, _nodeHeight + _rowGap);
            foreach (IGrouping<int, NodeItem> layer in _nodes.GroupBy(node => node.Layer).OrderBy(group => group.Key)) {
                NodeItem[] layerNodes = layer.OrderBy(node => _nodes.IndexOf(node)).ToArray();
                double radius = layer.Key == 0 && layerNodes.Length == 1 ? 0D : Math.Max(0.9D, layer.Key) * ringGap;
                for (int i = 0; i < layerNodes.Length; i++) {
                    double angle = layerNodes.Length == 1 ? -Math.PI / 2D : (-Math.PI / 2D) + (2D * Math.PI * i / layerNodes.Length);
                    layerNodes[i].PinX = centerX + Math.Cos(angle) * radius;
                    layerNodes[i].PinY = centerY + Math.Sin(angle) * radius;
                }
            }
        }

        private void AddZones(VisioPage page) {
            foreach (ZoneItem zone in _zones) {
                GetZoneBounds(zone, out double left, out double bottom, out double right, out double top);
                double width = right - left;
                double height = top - bottom;
                VisioShape shape = VisioNetworkDiagramVisuals.CreateBackgroundZone(
                    _document,
                    zone.Id,
                    left + width / 2D,
                    bottom + height / 2D,
                    width,
                    height,
                    string.Empty,
                    _theme);
                page.Shapes.Add(shape);
                VisioShape label = page.AddTextBox(CreateGeneratedId(zone.Id + "-label"), left + width / 2D, top - 0.18D, Math.Max(0.8D, width - 0.25D), 0.28D, zone.Text);
                VisioTextStyle labelStyle = _theme.Container.TextStyle?.Clone() ?? new VisioTextStyle();
                labelStyle.FontFamily = string.IsNullOrWhiteSpace(labelStyle.FontFamily) ? "Aptos" : labelStyle.FontFamily;
                labelStyle.Size = Math.Max(labelStyle.Size ?? 0D, 9.5D);
                labelStyle.Bold = true;
                labelStyle.HorizontalAlignment = VisioTextHorizontalAlignment.Center;
                labelStyle.VerticalAlignment = VisioTextVerticalAlignment.Middle;
                label.TextStyle = labelStyle;
            }
        }

        private void AddNodes(VisioPage page) {
            foreach (NodeItem node in _nodes) {
                GetNodeShape(node, out string masterNameU, out double width, out double height);
                VisioShape shape;
                if (node.Stencil != null) {
                    shape = page.AddStencilShape(node.Stencil, node.Id, node.PinX, node.PinY, width, height, string.Empty);
                } else {
                    shape = new VisioShape(node.Id, node.PinX, node.PinY, width, height, node.Text) {
                        NameU = masterNameU,
                    };
                    GetNodeStyle(node.Kind).ApplyTo(shape);
                    page.Shapes.Add(shape);
                }

                node.Shape = shape;
                if (node.Stencil != null && !string.IsNullOrWhiteSpace(node.Text)) {
                    AddStencilNodeCaption(page, node, width, height);
                }
            }
        }

        private void AddStencilNodeCaption(VisioPage page, NodeItem node, double width, double height) {
            VisioShape label = page.AddTextBox(
                CreateGeneratedId(node.Id + "-label"),
                node.PinX,
                node.PinY - (height / 2D) - 0.26D,
                Math.Max(1.15D, width + 0.55D),
                0.34D,
                node.Text);
            label.TextStyle = new VisioTextStyle {
                FontFamily = "Aptos",
                Size = 9.5D,
                Color = Color.FromRgb(25, 35, 45),
                BackgroundColor = Color.White,
                BackgroundTransparency = 0,
                HorizontalAlignment = VisioTextHorizontalAlignment.Center,
                VerticalAlignment = VisioTextVerticalAlignment.Middle
            };
        }

        private void AddEdges(VisioPage page) {
            int routeIndex = 0;
            foreach (EdgeItem edge in _edges) {
                NodeItem from = _nodesById[edge.FromId];
                NodeItem to = _nodesById[edge.ToId];
                if (from.Shape == null || to.Shape == null) {
                    throw new InvalidOperationException("Nodes must be placed before graph edges are created.");
                }

                VisioNetworkDiagramVisuals.ResolveSides(from.Shape, to.Shape, out VisioSide fromSide, out VisioSide toSide);
                ConnectorKind connectorKind = _layout == VisioGraphLayout.Radial ? ConnectorKind.Straight : ConnectorKind.RightAngle;
                VisioConnector connector = page.AddConnector(from.Shape, to.Shape, connectorKind, fromSide, toSide);
                GetConnectorStyle(edge.Kind, edge.Directed).ApplyTo(connector);
                connector.Label = edge.Label;
                if (_layout != VisioGraphLayout.Radial) {
                    connector.RouteOrthogonal(offset: (routeIndex % 7) * 0.05D);
                }

                if (!string.IsNullOrWhiteSpace(edge.Label)) {
                    VisioTextStyle labelStyle = connector.TextStyle?.Clone() ?? new VisioTextStyle();
                    labelStyle.BackgroundColor = Color.White;
                    labelStyle.BackgroundTransparency = 0;
                    labelStyle.Size = Math.Max(labelStyle.Size ?? 0D, 8.5D);
                    connector.TextStyle = labelStyle;
                    connector.PlaceLabel(0.5D, offsetY: 0.26D, width: 1.35D, height: 0.32D);
                    connector.ResizeLabelToText(maximumWidth: 1.5D);
                }

                routeIndex++;
            }
        }

        private void AddTitle(VisioPage page) {
            if (string.IsNullOrWhiteSpace(_titleText)) {
                return;
            }

            double y = _pageHeight - _topMargin - (_titleHeight / 2D);
            VisioShape title = page.AddTextBox(_titleId, _pageWidth / 2D, y, Math.Max(1D, _pageWidth - _leftMargin - _rightMargin), _titleHeight, _titleText, _unit);
            VisioTextStyle style = _theme.Emphasis.TextStyle?.Clone() ?? new VisioTextStyle();
            style.FontFamily = string.IsNullOrWhiteSpace(style.FontFamily) ? "Aptos Display" : style.FontFamily;
            style.Size = Math.Max(style.Size ?? 0D, 20D);
            style.Bold = true;
            style.Color = Color.FromRgb(32, 55, 75);
            style.HorizontalAlignment = VisioTextHorizontalAlignment.Center;
            style.VerticalAlignment = VisioTextVerticalAlignment.Middle;
            title.TextStyle = style;
        }

        private void GetZoneBounds(ZoneItem zone, out double left, out double bottom, out double right, out double top) {
            const double horizontalPadding = 0.45D;
            const double verticalPadding = 0.35D;
            left = double.MaxValue;
            bottom = double.MaxValue;
            right = double.MinValue;
            top = double.MinValue;
            foreach (string nodeId in zone.NodeIds) {
                NodeItem node = _nodesById[nodeId];
                GetNodeShape(node, out _, out double width, out double height);
                left = Math.Min(left, node.PinX - width / 2D);
                bottom = Math.Min(bottom, node.PinY - height / 2D);
                right = Math.Max(right, node.PinX + width / 2D);
                top = Math.Max(top, node.PinY + height / 2D);
                if (node.Stencil != null && !string.IsNullOrWhiteSpace(node.Text)) {
                    bottom = Math.Min(bottom, node.PinY - height / 2D - 0.52D);
                }
            }

            left -= horizontalPadding;
            bottom -= verticalPadding;
            right += horizontalPadding;
            top += verticalPadding;
        }

        private double XForLayer(int layer) {
            return _leftMargin + (_nodeWidth / 2D) + layer * (_nodeWidth + _columnGap);
        }

        private double XForRow(int row) {
            double contentWidth = _maximumRows * _nodeWidth + Math.Max(0, _maximumRows - 1) * _columnGap;
            double availableWidth = _pageWidth - _leftMargin - _rightMargin;
            double start = _leftMargin + Math.Max(0D, (availableWidth - contentWidth) / 2D);
            return start + (_nodeWidth / 2D) + row * (_nodeWidth + _columnGap);
        }

        private double YForLayer(int layer) {
            double top = _pageHeight - _topMargin - HeaderHeight;
            return top - (_nodeHeight / 2D) - layer * (_nodeHeight + _rowGap);
        }

        private double YForRow(int row) {
            double contentHeight = _maximumRows * _nodeHeight + Math.Max(0, _maximumRows - 1) * _rowGap;
            double top = _pageHeight - _topMargin - HeaderHeight;
            double availableHeight = _pageHeight - _topMargin - _bottomMargin - HeaderHeight;
            double layerTop = top - Math.Max(0D, (availableHeight - contentHeight) / 2D);
            return layerTop - (_nodeHeight / 2D) - row * (_nodeHeight + _rowGap);
        }

        private double HeaderHeight => string.IsNullOrWhiteSpace(_titleText) ? 0D : _titleHeight + _titleGap;

        private void GetNodeShape(NodeItem node, out string masterNameU, out double width, out double height) {
            width = node.Stencil?.DefaultWidth ?? _nodeWidth;
            height = node.Stencil?.DefaultHeight ?? _nodeHeight;
            if (node.Stencil != null) {
                masterNameU = node.Stencil.MasterNameU;
                return;
            }

            switch (node.Kind) {
                case VisioGraphNodeKind.Data:
                    masterNameU = "Data";
                    break;
                case VisioGraphNodeKind.External:
                    masterNameU = "Circle";
                    width = Math.Min(_nodeWidth, _nodeHeight * 1.15D);
                    height = width;
                    break;
                case VisioGraphNodeKind.Decision:
                    masterNameU = "Decision";
                    height = _nodeHeight * 1.2D;
                    break;
                default:
                    masterNameU = "Process";
                    break;
            }
        }

        private VisioShapeStyle GetNodeStyle(VisioGraphNodeKind kind) {
            switch (kind) {
                case VisioGraphNodeKind.Data:
                    return _theme.Marker;
                case VisioGraphNodeKind.External:
                    return _theme.Success;
                case VisioGraphNodeKind.Decision:
                    return _theme.Decision;
                case VisioGraphNodeKind.Emphasis:
                    return _theme.Emphasis;
                default:
                    return _theme.Primary;
            }
        }

        private VisioConnectorStyle GetConnectorStyle(VisioGraphConnectorKind kind, bool directed) {
            VisioConnectorStyle style;
            switch (kind) {
                case VisioGraphConnectorKind.Data:
                    style = _theme.DataConnector.Clone();
                    break;
                case VisioGraphConnectorKind.Control:
                    style = _theme.ControlConnector.Clone();
                    break;
                case VisioGraphConnectorKind.Emphasis:
                    style = _theme.Connector.Clone();
                    style.LineWeight = Math.Max(style.LineWeight, 0.026D);
                    break;
                default:
                    style = _theme.Connector.Clone();
                    break;
            }

            if (!directed) {
                style.EndArrow = EndArrow.None;
            }

            return style;
        }

        private void AddRoot(string id) {
            if (_rootIdSet.Add(id)) {
                _rootIds.Add(id);
            }
        }

        private bool IsIdInUse(string id) {
            if (!string.IsNullOrWhiteSpace(_titleText) && string.Equals(_titleId, id, StringComparison.Ordinal)) {
                return true;
            }

            if (_nodesById.ContainsKey(id) || _zoneIds.Contains(id)) {
                return true;
            }

            return false;
        }

        private string CreateGeneratedId(string baseId) {
            string id = baseId;
            int index = 2;
            while (IsIdInUse(id) || _generatedIds.Contains(id)) {
                id = baseId + "-" + index;
                index++;
            }

            _generatedIds.Add(id);
            return id;
        }

        private void EnsureKnownNode(string id, string parameterName) {
            if (!_nodesById.ContainsKey(id)) {
                throw new ArgumentException($"Unknown graph node id '{id}'.", parameterName);
            }
        }

        private static IReadOnlyList<string> NormalizeZoneNodeIds(string[] nodeIds) {
            if (nodeIds == null) throw new ArgumentNullException(nameof(nodeIds));
            List<string> normalizedNodeIds = new();
            HashSet<string> seen = new(StringComparer.Ordinal);
            for (int i = 0; i < nodeIds.Length; i++) {
                string normalizedId = RequireId(nodeIds[i], nameof(nodeIds), "Zone node id");
                if (seen.Add(normalizedId)) {
                    normalizedNodeIds.Add(normalizedId);
                }
            }

            if (normalizedNodeIds.Count == 0) {
                throw new ArgumentException("A graph zone requires at least one node id.", nameof(nodeIds));
            }

            return normalizedNodeIds.AsReadOnly();
        }

        private static string RequireId(string id, string parameterName, string label) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException(label + " cannot be null or whitespace.", parameterName);
            }

            return id.Trim();
        }

        private static void ValidatePositive(double value, string parameterName) {
            if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
                throw new ArgumentOutOfRangeException(parameterName, "Value must be positive.");
            }
        }

        private static void ValidateNonNegative(double value, string parameterName) {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
                throw new ArgumentOutOfRangeException(parameterName, "Value must be zero or greater.");
            }
        }
    }
}
