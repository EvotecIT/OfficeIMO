using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level builder for network topologies where OfficeIMO places devices from links.
    /// </summary>
    public sealed class VisioNetworkTopologyDiagramBuilder {
        private sealed class NodeItem {
            public NodeItem(string id, string text, VisioNetworkNodeKind kind) {
                Id = id;
                Text = text;
                Kind = kind;
            }

            public string Id { get; }

            public string Text { get; }

            public VisioNetworkNodeKind Kind { get; }

            public int Layer { get; set; }

            public int Row { get; set; }

            public VisioShape? Shape { get; set; }
        }

        private sealed class LinkItem {
            public LinkItem(string fromId, string toId, VisioNetworkLinkKind kind, string? label) {
                FromId = fromId;
                ToId = toId;
                Kind = kind;
                Label = label;
            }

            public string FromId { get; }

            public string ToId { get; }

            public VisioNetworkLinkKind Kind { get; }

            public string? Label { get; }
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
        private readonly List<string> _rootIds = new();
        private readonly HashSet<string> _rootIdSet = new(StringComparer.Ordinal);
        private readonly List<ZoneItem> _zones = new();
        private readonly HashSet<string> _zoneIds = new(StringComparer.Ordinal);
        private readonly List<LinkItem> _links = new();
        private VisioStyleTheme _theme = VisioStyleTheme.Technical();
        private VisioMeasurementUnit _unit = VisioMeasurementUnit.Inches;
        private double _pageWidth = 11;
        private double _pageHeight = 8.5;
        private double _leftMargin = 0.8;
        private double _topMargin = 0.8;
        private double _rightMargin = 0.8;
        private double _bottomMargin = 0.8;
        private double _nodeWidth = 1.45;
        private double _nodeHeight = 0.85;
        private double _columnGap = 1.15;
        private double _rowGap = 0.85;
        private string? _titleText;
        private string _titleId = "title";
        private double _titleHeight = 0.45;
        private double _titleGap = 0.35;
        private int _maximumLayerRows = 1;
        private bool _fitPageToTopology = true;
        private bool _built;

        internal VisioNetworkTopologyDiagramBuilder(VisioDocument document, string pageName) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _pageName = string.IsNullOrWhiteSpace(pageName) ? "Network Topology" : pageName;
        }

        /// <summary>Sets the page size used by the generated topology page.</summary>
        public VisioNetworkTopologyDiagramBuilder PageSize(double width, double height, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _pageWidth = width;
            _pageHeight = height;
            _unit = unit;
            return this;
        }

        /// <summary>Sets whether the builder can grow the page to fit the topology. Enabled by default.</summary>
        public VisioNetworkTopologyDiagramBuilder FitPageToTopology(bool enabled = true) {
            _fitPageToTopology = enabled;
            return this;
        }

        /// <summary>Sets the visual theme.</summary>
        public VisioNetworkTopologyDiagramBuilder Theme(VisioStyleTheme theme) {
            _theme = (theme ?? throw new ArgumentNullException(nameof(theme))).Clone();
            return this;
        }

        /// <summary>Sets outer page margins used by the automatic layout.</summary>
        public VisioNetworkTopologyDiagramBuilder Margins(double left, double top, double right = 0.8, double bottom = 0.8) {
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

        /// <summary>Sets the default network node size.</summary>
        public VisioNetworkTopologyDiagramBuilder NodeSize(double width, double height) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _nodeWidth = width;
            _nodeHeight = height;
            return this;
        }

        /// <summary>Sets spacing between automatic layers and rows.</summary>
        public VisioNetworkTopologyDiagramBuilder Spacing(double columnGap, double rowGap) {
            ValidateNonNegative(columnGap, nameof(columnGap));
            ValidateNonNegative(rowGap, nameof(rowGap));
            _columnGap = columnGap;
            _rowGap = rowGap;
            return this;
        }

        /// <summary>Adds a centered editable title above the automatically placed topology.</summary>
        public VisioNetworkTopologyDiagramBuilder Title(string? text = null, string id = "title", double height = 0.45, double gap = 0.35) {
            string normalizedId = RequireId(id, nameof(id), "Title id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A network topology item with id '{normalizedId}' already exists.", nameof(id));
            }

            ValidatePositive(height, nameof(height));
            ValidateNonNegative(gap, nameof(gap));
            _titleText = string.IsNullOrWhiteSpace(text) ? _pageName : text;
            _titleId = normalizedId;
            _titleHeight = height;
            _titleGap = gap;
            return this;
        }

        /// <summary>Adds a background zone around automatically placed topology nodes.</summary>
        public VisioNetworkTopologyDiagramBuilder Zone(string id, string text, params string[] nodeIds) {
            string normalizedId = RequireId(id, nameof(id), "Zone id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A network topology item with id '{normalizedId}' already exists.", nameof(id));
            }

            IReadOnlyList<string> normalizedNodeIds = NormalizeZoneNodeIds(nodeIds);
            _zones.Add(new ZoneItem(normalizedId, text ?? string.Empty, normalizedNodeIds));
            _zoneIds.Add(normalizedId);
            return this;
        }

        /// <summary>Adds a subnet-style background zone around automatically placed topology nodes.</summary>
        public VisioNetworkTopologyDiagramBuilder Subnet(string id, string text, params string[] nodeIds) => Zone(id, text, nodeIds);

        /// <summary>Adds and marks a root node used as the starting point for automatic layout.</summary>
        public VisioNetworkTopologyDiagramBuilder Root(string id, string text, VisioNetworkNodeKind kind = VisioNetworkNodeKind.Internet) {
            Node(id, text, kind);
            AddRoot(RequireId(id, nameof(id), "Root node id"));
            return this;
        }

        /// <summary>Marks an existing node as a root used by automatic layout.</summary>
        public VisioNetworkTopologyDiagramBuilder Root(string id) {
            string normalizedId = RequireId(id, nameof(id), "Root node id");
            EnsureKnownNode(normalizedId, nameof(id));
            AddRoot(normalizedId);
            return this;
        }

        /// <summary>Adds a network node that will be placed by automatic topology layout.</summary>
        public VisioNetworkTopologyDiagramBuilder Node(string id, string text, VisioNetworkNodeKind kind = VisioNetworkNodeKind.Server) {
            string normalizedId = RequireId(id, nameof(id), "Node id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A network topology item with id '{normalizedId}' already exists.", nameof(id));
            }

            if (!Enum.IsDefined(typeof(VisioNetworkNodeKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            NodeItem node = new(normalizedId, text ?? string.Empty, kind);
            _nodes.Add(node);
            _nodesById.Add(normalizedId, node);
            return this;
        }

        /// <summary>Adds a user/client node.</summary>
        public VisioNetworkTopologyDiagramBuilder User(string id, string text) => Node(id, text, VisioNetworkNodeKind.User);

        /// <summary>Adds a workstation node.</summary>
        public VisioNetworkTopologyDiagramBuilder Workstation(string id, string text) => Node(id, text, VisioNetworkNodeKind.Workstation);

        /// <summary>Adds a server node.</summary>
        public VisioNetworkTopologyDiagramBuilder Server(string id, string text) => Node(id, text, VisioNetworkNodeKind.Server);

        /// <summary>Adds a switch node.</summary>
        public VisioNetworkTopologyDiagramBuilder Switch(string id, string text) => Node(id, text, VisioNetworkNodeKind.Switch);

        /// <summary>Adds a router node.</summary>
        public VisioNetworkTopologyDiagramBuilder Router(string id, string text) => Node(id, text, VisioNetworkNodeKind.Router);

        /// <summary>Adds a firewall node.</summary>
        public VisioNetworkTopologyDiagramBuilder Firewall(string id, string text) => Node(id, text, VisioNetworkNodeKind.Firewall);

        /// <summary>Adds an Internet/external network node.</summary>
        public VisioNetworkTopologyDiagramBuilder Internet(string id, string text) => Node(id, text, VisioNetworkNodeKind.Internet);

        /// <summary>Adds a printer node.</summary>
        public VisioNetworkTopologyDiagramBuilder Printer(string id, string text) => Node(id, text, VisioNetworkNodeKind.Printer);

        /// <summary>Adds a storage node.</summary>
        public VisioNetworkTopologyDiagramBuilder Storage(string id, string text) => Node(id, text, VisioNetworkNodeKind.Storage);

        /// <summary>Adds a database node.</summary>
        public VisioNetworkTopologyDiagramBuilder Database(string id, string text) => Node(id, text, VisioNetworkNodeKind.Database);

        /// <summary>Adds a wireless access point node.</summary>
        public VisioNetworkTopologyDiagramBuilder Wireless(string id, string text) => Node(id, text, VisioNetworkNodeKind.Wireless);

        /// <summary>Adds a note or legend node.</summary>
        public VisioNetworkTopologyDiagramBuilder Legend(string id, string text) => Node(id, text, VisioNetworkNodeKind.Note);

        /// <summary>Adds a standard network link.</summary>
        public VisioNetworkTopologyDiagramBuilder Ethernet(string fromId, string toId, string? label = null) => Link(fromId, toId, VisioNetworkLinkKind.Ethernet, label);

        /// <summary>Adds a trunk/uplink connection.</summary>
        public VisioNetworkTopologyDiagramBuilder Trunk(string fromId, string toId, string? label = null) => Link(fromId, toId, VisioNetworkLinkKind.Trunk, label);

        /// <summary>Adds a wireless connection.</summary>
        public VisioNetworkTopologyDiagramBuilder WirelessLink(string fromId, string toId, string? label = null) => Link(fromId, toId, VisioNetworkLinkKind.Wireless, label);

        /// <summary>Adds a management connection.</summary>
        public VisioNetworkTopologyDiagramBuilder Management(string fromId, string toId, string? label = null) => Link(fromId, toId, VisioNetworkLinkKind.Management, label);

        /// <summary>Adds a link between two known network nodes.</summary>
        public VisioNetworkTopologyDiagramBuilder Link(string fromId, string toId, VisioNetworkLinkKind kind, string? label = null) {
            string normalizedFromId = RequireId(fromId, nameof(fromId), "From node id");
            string normalizedToId = RequireId(toId, nameof(toId), "To node id");
            EnsureKnownNode(normalizedFromId, nameof(fromId));
            EnsureKnownNode(normalizedToId, nameof(toId));
            if (!Enum.IsDefined(typeof(VisioNetworkLinkKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            _links.Add(new LinkItem(normalizedFromId, normalizedToId, kind, label));
            return this;
        }

        internal VisioPage Build() {
            if (_built) {
                throw new InvalidOperationException("This network topology diagram builder has already produced a page.");
            }

            _built = true;
            if (_nodes.Count == 0) {
                throw new InvalidOperationException("A network topology diagram requires at least one node.");
            }

            AssignLayout();
            ValidateZones();
            SizePageForLayout();

            VisioPage page = _document.AddPage(_pageName, _pageWidth, _pageHeight, _unit);
            page.Grid(visible: false, snap: true);
            AddZones(page);
            AddNodes(page);
            AddLinks(page);
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

            _maximumLayerRows = Math.Max(1, _nodes.GroupBy(node => node.Layer).Max(group => group.Count()));
        }

        private void AssignLayout() {
            Dictionary<string, List<string>> adjacency = _nodes.ToDictionary(node => node.Id, _ => new List<string>(), StringComparer.Ordinal);
            foreach (LinkItem link in _links) {
                adjacency[link.FromId].Add(link.ToId);
                adjacency[link.ToId].Add(link.FromId);
            }

            HashSet<string> assigned = new(StringComparer.Ordinal);
            Queue<NodeItem> ready = new();

            void Enqueue(NodeItem node, int layer) {
                if (assigned.Add(node.Id)) {
                    node.Layer = layer;
                    ready.Enqueue(node);
                }
            }

            if (_rootIds.Count == 0) {
                Enqueue(_nodes[0], 0);
            } else {
                foreach (string rootId in _rootIds) {
                    Enqueue(_nodesById[rootId], 0);
                }
            }

            while (assigned.Count < _nodes.Count || ready.Count > 0) {
                while (ready.Count > 0) {
                    NodeItem node = ready.Dequeue();
                    foreach (string neighborId in adjacency[node.Id]) {
                        if (!assigned.Contains(neighborId)) {
                            Enqueue(_nodesById[neighborId], node.Layer + 1);
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
        }

        private void SizePageForLayout() {
            if (!_fitPageToTopology) {
                return;
            }

            int layerCount = _nodes.Max(node => node.Layer) + 1;
            int rowCount = _nodes.GroupBy(node => node.Layer).Max(group => group.Count());
            double requiredWidth = _leftMargin + _rightMargin + (layerCount * _nodeWidth) + Math.Max(0, layerCount - 1) * _columnGap;
            double requiredHeight = _topMargin + _bottomMargin + HeaderHeight + (rowCount * _nodeHeight) + Math.Max(0, rowCount - 1) * _rowGap;
            _pageWidth = Math.Max(_pageWidth, requiredWidth);
            _pageHeight = Math.Max(_pageHeight, requiredHeight);
        }

        private void AddTitle(VisioPage page) {
            if (string.IsNullOrWhiteSpace(_titleText)) {
                return;
            }

            double y = _pageHeight - _topMargin - (_titleHeight / 2D);
            VisioShape title = page.AddTextBox(_titleId, _pageWidth / 2D, y, Math.Max(1D, _pageWidth - _leftMargin - _rightMargin), _titleHeight, _titleText, _unit);
            title.TextStyle = CreateTitleTextStyle();
        }

        private VisioTextStyle CreateTitleTextStyle() {
            VisioTextStyle style = _theme.Emphasis.TextStyle?.Clone() ?? new VisioTextStyle();
            style.FontFamily = string.IsNullOrWhiteSpace(style.FontFamily) ? "Aptos Display" : style.FontFamily;
            style.Size = Math.Max(style.Size ?? 0D, 20D);
            style.Bold = true;
            style.HorizontalAlignment = VisioTextHorizontalAlignment.Center;
            style.VerticalAlignment = VisioTextVerticalAlignment.Middle;
            return style;
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
                    zone.Text,
                    _theme);
                page.Shapes.Add(shape);
            }
        }

        private void AddNodes(VisioPage page) {
            foreach (NodeItem node in _nodes) {
                VisioNetworkDiagramVisuals.GetNodeShape(node.Kind, _nodeWidth, _nodeHeight, out string masterNameU, out double width, out double height);
                VisioShape shape = new(node.Id, XForLayer(node.Layer), YForRow(node.Layer, node.Row), width, height, node.Text) {
                    NameU = masterNameU,
                    Master = _document.EnsureBuiltinMaster(masterNameU)
                };
                VisioNetworkDiagramVisuals.GetNodeStyle(_theme, node.Kind).ApplyTo(shape);
                node.Shape = shape;
                page.Shapes.Add(shape);
            }
        }

        private void AddLinks(VisioPage page) {
            int routeIndex = 0;
            foreach (LinkItem link in _links) {
                NodeItem from = _nodesById[link.FromId];
                NodeItem to = _nodesById[link.ToId];
                if (from.Shape == null || to.Shape == null) {
                    throw new InvalidOperationException("Nodes must be placed before topology links are created.");
                }

                VisioNetworkDiagramVisuals.ResolveSides(from.Shape, to.Shape, out VisioSide fromSide, out VisioSide toSide);
                VisioConnector connector = page.AddConnector(from.Shape, to.Shape, ConnectorKind.RightAngle, fromSide, toSide);
                VisioNetworkDiagramVisuals.GetConnectorStyle(_theme, link.Kind).ApplyTo(connector);
                connector.Label = link.Label;
                connector.RouteOrthogonal(offset: (routeIndex % 5) * 0.06);
                if (!string.IsNullOrWhiteSpace(link.Label)) {
                    connector.PlaceLabel(0.5, offsetY: 0.16);
                    connector.ResizeLabelToText(maximumWidth: 1.4);
                }

                routeIndex++;
            }
        }

        private double XForLayer(int layer) {
            return _leftMargin + (_nodeWidth / 2D) + layer * (_nodeWidth + _columnGap);
        }

        private double YForRow(int layer, int row) {
            double contentHeight = _maximumLayerRows * _nodeHeight + Math.Max(0, _maximumLayerRows - 1) * _rowGap;
            double top = _pageHeight - _topMargin - HeaderHeight;
            double availableHeight = _pageHeight - _topMargin - _bottomMargin - HeaderHeight;
            double layerTop = top - Math.Max(0D, (availableHeight - contentHeight) / 2D);
            return layerTop - (_nodeHeight / 2D) - row * (_nodeHeight + _rowGap);
        }

        private double HeaderHeight => string.IsNullOrWhiteSpace(_titleText) ? 0D : _titleHeight + _titleGap;

        private void GetZoneBounds(ZoneItem zone, out double left, out double bottom, out double right, out double top) {
            const double horizontalPadding = 0.45D;
            const double verticalPadding = 0.35D;

            left = double.MaxValue;
            bottom = double.MaxValue;
            right = double.MinValue;
            top = double.MinValue;

            foreach (string nodeId in zone.NodeIds) {
                NodeItem node = _nodesById[nodeId];
                GetNodeBounds(node, out double nodeLeft, out double nodeBottom, out double nodeRight, out double nodeTop);
                left = Math.Min(left, nodeLeft);
                bottom = Math.Min(bottom, nodeBottom);
                right = Math.Max(right, nodeRight);
                top = Math.Max(top, nodeTop);
            }

            left -= horizontalPadding;
            bottom -= verticalPadding;
            right += horizontalPadding;
            top += verticalPadding;
        }

        private void GetNodeBounds(NodeItem node, out double left, out double bottom, out double right, out double top) {
            VisioNetworkDiagramVisuals.GetNodeShape(node.Kind, _nodeWidth, _nodeHeight, out _, out double width, out double height);
            double pinX = XForLayer(node.Layer);
            double pinY = YForRow(node.Layer, node.Row);
            left = pinX - width / 2D;
            bottom = pinY - height / 2D;
            right = pinX + width / 2D;
            top = pinY + height / 2D;
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

            return _nodesById.ContainsKey(id) || _zoneIds.Contains(id);
        }

        private void EnsureKnownNode(string id, string parameterName) {
            if (!_nodesById.ContainsKey(id)) {
                throw new ArgumentException($"Unknown network node id '{id}'.", parameterName);
            }
        }

        private static IReadOnlyList<string> NormalizeZoneNodeIds(string[] nodeIds) {
            if (nodeIds == null) {
                throw new ArgumentNullException(nameof(nodeIds));
            }

            List<string> normalizedNodeIds = new();
            HashSet<string> seen = new(StringComparer.Ordinal);
            for (int i = 0; i < nodeIds.Length; i++) {
                string normalizedId = RequireId(nodeIds[i], nameof(nodeIds), "Zone node id");
                if (seen.Add(normalizedId)) {
                    normalizedNodeIds.Add(normalizedId);
                }
            }

            if (normalizedNodeIds.Count == 0) {
                throw new ArgumentException("A network topology zone requires at least one node id.", nameof(nodeIds));
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
