using System;
using System.Collections.Generic;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level builder for dependency-free network diagrams with zones,
    /// devices, links, and legends.
    /// </summary>
    public sealed class VisioNetworkDiagramBuilder {
        private sealed class NodeItem {
            public NodeItem(string id, string text, int column, int row, VisioNetworkNodeKind kind, IDictionary<string, string?>? shapeData = null, string? hyperlinkAddress = null, string? hyperlinkDescription = null) {
                Id = id;
                Text = text;
                Column = column;
                Row = row;
                Kind = kind;
                ShapeData = CopyShapeData(shapeData);
                HyperlinkAddress = hyperlinkAddress;
                HyperlinkDescription = hyperlinkDescription;
            }

            public string Id { get; }

            public string Text { get; }

            public int Column { get; }

            public int Row { get; }

            public VisioNetworkNodeKind Kind { get; }

            public VisioShape? Shape { get; set; }

            public IReadOnlyDictionary<string, string?> ShapeData { get; }

            public string? HyperlinkAddress { get; }

            public string? HyperlinkDescription { get; }
        }

        private sealed class ZoneItem {
            public ZoneItem(string id, string text, int column, int row, int columnSpan, int rowSpan, IDictionary<string, string?>? shapeData = null, string? hyperlinkAddress = null, string? hyperlinkDescription = null) {
                Id = id;
                Text = text;
                Column = column;
                Row = row;
                ColumnSpan = columnSpan;
                RowSpan = rowSpan;
                ShapeData = CopyShapeData(shapeData);
                HyperlinkAddress = hyperlinkAddress;
                HyperlinkDescription = hyperlinkDescription;
            }

            public string Id { get; }

            public string Text { get; }

            public int Column { get; }

            public int Row { get; }

            public int ColumnSpan { get; }

            public int RowSpan { get; }

            public IReadOnlyDictionary<string, string?> ShapeData { get; }

            public string? HyperlinkAddress { get; }

            public string? HyperlinkDescription { get; }
        }

        private sealed class LinkItem {
            public LinkItem(string? id, string fromId, string toId, VisioNetworkLinkKind kind, string? label, IDictionary<string, string?>? shapeData = null, string? hyperlinkAddress = null, string? hyperlinkDescription = null) {
                Id = id;
                FromId = fromId;
                ToId = toId;
                Kind = kind;
                Label = label;
                ShapeData = CopyShapeData(shapeData);
                HyperlinkAddress = hyperlinkAddress;
                HyperlinkDescription = hyperlinkDescription;
            }

            public string? Id { get; }

            public string FromId { get; }

            public string ToId { get; }

            public VisioNetworkLinkKind Kind { get; }

            public string? Label { get; }

            public IReadOnlyDictionary<string, string?> ShapeData { get; }

            public string? HyperlinkAddress { get; }

            public string? HyperlinkDescription { get; }
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
        private readonly List<NodeItem> _nodes = new();
        private readonly Dictionary<string, NodeItem> _nodesById = new(StringComparer.Ordinal);
        private readonly List<ZoneItem> _zones = new();
        private readonly List<LinkItem> _links = new();
        private readonly HashSet<string> _linkIds = new(StringComparer.Ordinal);
        private readonly List<CalloutItem> _callouts = new();
        private VisioStyleTheme _theme = VisioStyleTheme.Technical();
        private VisioMeasurementUnit _unit = VisioMeasurementUnit.Inches;
        private double _pageWidth = 14;
        private double _pageHeight = 8.5;
        private double _leftMargin = 0.9;
        private double _topMargin = 0.8;
        private double _columnGap = 0.75;
        private double _rowGap = 0.85;
        private double _nodeWidth = 1.45;
        private double _nodeHeight = 0.85;
        private string? _titleText;
        private string _titleId = "title";
        private double _titleHeight = 0.45;
        private double _titleGap = 0.35;
        private bool _built;

        internal VisioNetworkDiagramBuilder(VisioDocument document, string pageName) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _pageName = string.IsNullOrWhiteSpace(pageName) ? "Network Diagram" : pageName;
        }

        /// <summary>Sets the page size used by the generated network page.</summary>
        public VisioNetworkDiagramBuilder PageSize(double width, double height, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _pageWidth = width;
            _pageHeight = height;
            _unit = unit;
            return this;
        }

        /// <summary>Sets the visual theme.</summary>
        public VisioNetworkDiagramBuilder Theme(VisioStyleTheme theme) {
            _theme = (theme ?? throw new ArgumentNullException(nameof(theme))).Clone();
            return this;
        }

        /// <summary>Sets outer page margins used by the grid layout.</summary>
        public VisioNetworkDiagramBuilder Margins(double left, double top) {
            ValidateNonNegative(left, nameof(left));
            ValidateNonNegative(top, nameof(top));
            _leftMargin = left;
            _topMargin = top;
            return this;
        }

        /// <summary>Sets grid spacing between nodes.</summary>
        public VisioNetworkDiagramBuilder Spacing(double columnGap, double rowGap) {
            ValidateNonNegative(columnGap, nameof(columnGap));
            ValidateNonNegative(rowGap, nameof(rowGap));
            _columnGap = columnGap;
            _rowGap = rowGap;
            return this;
        }

        /// <summary>Sets the default network node size.</summary>
        public VisioNetworkDiagramBuilder NodeSize(double width, double height) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _nodeWidth = width;
            _nodeHeight = height;
            return this;
        }

        /// <summary>Adds a centered editable title above the generated network grid.</summary>
        public VisioNetworkDiagramBuilder Title(string? text = null, string id = "title", double height = 0.45, double gap = 0.35) {
            string normalizedId = RequireId(id, nameof(id), "Title id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A network item with id '{normalizedId}' already exists.", nameof(id));
            }

            ValidatePositive(height, nameof(height));
            ValidateNonNegative(gap, nameof(gap));
            _titleText = string.IsNullOrWhiteSpace(text) ? _pageName : text;
            _titleId = normalizedId;
            _titleHeight = height;
            _titleGap = gap;
            return this;
        }

        /// <summary>Adds a background zone around a grid area.</summary>
        public VisioNetworkDiagramBuilder Zone(string id, string text, int column, int row, int columnSpan, int rowSpan) {
            string normalizedId = RequireId(id, nameof(id), "Zone id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A network item with id '{normalizedId}' already exists.", nameof(id));
            }

            ValidateGridPosition(column, row);
            if (columnSpan <= 0) throw new ArgumentOutOfRangeException(nameof(columnSpan), "Column span must be positive.");
            if (rowSpan <= 0) throw new ArgumentOutOfRangeException(nameof(rowSpan), "Row span must be positive.");
            _zones.Add(new ZoneItem(normalizedId, text ?? string.Empty, column, row, columnSpan, rowSpan));
            return this;
        }

        /// <summary>Adds a node at a deterministic grid position.</summary>
        public VisioNetworkDiagramBuilder Node(string id, string text, int column, int row, VisioNetworkNodeKind kind = VisioNetworkNodeKind.Server) {
            string normalizedId = RequireId(id, nameof(id), "Node id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A network item with id '{normalizedId}' already exists.", nameof(id));
            }

            ValidateGridPosition(column, row);
            if (!Enum.IsDefined(typeof(VisioNetworkNodeKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            NodeItem item = new(normalizedId, text ?? string.Empty, column, row, kind);
            _nodes.Add(item);
            _nodesById.Add(normalizedId, item);
            return this;
        }

        /// <summary>Adds a user/client node.</summary>
        public VisioNetworkDiagramBuilder User(string id, string text, int column, int row) => Node(id, text, column, row, VisioNetworkNodeKind.User);

        /// <summary>Adds a workstation node.</summary>
        public VisioNetworkDiagramBuilder Workstation(string id, string text, int column, int row) => Node(id, text, column, row, VisioNetworkNodeKind.Workstation);

        /// <summary>Adds a server node.</summary>
        public VisioNetworkDiagramBuilder Server(string id, string text, int column, int row) => Node(id, text, column, row, VisioNetworkNodeKind.Server);

        /// <summary>Adds a switch node.</summary>
        public VisioNetworkDiagramBuilder Switch(string id, string text, int column, int row) => Node(id, text, column, row, VisioNetworkNodeKind.Switch);

        /// <summary>Adds a router node.</summary>
        public VisioNetworkDiagramBuilder Router(string id, string text, int column, int row) => Node(id, text, column, row, VisioNetworkNodeKind.Router);

        /// <summary>Adds a firewall node.</summary>
        public VisioNetworkDiagramBuilder Firewall(string id, string text, int column, int row) => Node(id, text, column, row, VisioNetworkNodeKind.Firewall);

        /// <summary>Adds an Internet/external network node.</summary>
        public VisioNetworkDiagramBuilder Internet(string id, string text, int column, int row) => Node(id, text, column, row, VisioNetworkNodeKind.Internet);

        /// <summary>Adds a printer node.</summary>
        public VisioNetworkDiagramBuilder Printer(string id, string text, int column, int row) => Node(id, text, column, row, VisioNetworkNodeKind.Printer);

        /// <summary>Adds a storage node.</summary>
        public VisioNetworkDiagramBuilder Storage(string id, string text, int column, int row) => Node(id, text, column, row, VisioNetworkNodeKind.Storage);

        /// <summary>Adds a database node.</summary>
        public VisioNetworkDiagramBuilder Database(string id, string text, int column, int row) => Node(id, text, column, row, VisioNetworkNodeKind.Database);

        /// <summary>Adds a wireless access point node.</summary>
        public VisioNetworkDiagramBuilder Wireless(string id, string text, int column, int row) => Node(id, text, column, row, VisioNetworkNodeKind.Wireless);

        /// <summary>Adds a note or legend node.</summary>
        public VisioNetworkDiagramBuilder Legend(string id, string text, int column, int row) => Node(id, text, column, row, VisioNetworkNodeKind.Note);

        /// <summary>Adds a standard network link.</summary>
        public VisioNetworkDiagramBuilder Ethernet(string fromId, string toId, string? label = null) => Link(fromId, toId, VisioNetworkLinkKind.Ethernet, label);

        /// <summary>Adds a trunk/uplink connection.</summary>
        public VisioNetworkDiagramBuilder Trunk(string fromId, string toId, string? label = null) => Link(fromId, toId, VisioNetworkLinkKind.Trunk, label);

        /// <summary>Adds a wireless connection.</summary>
        public VisioNetworkDiagramBuilder WirelessLink(string fromId, string toId, string? label = null) => Link(fromId, toId, VisioNetworkLinkKind.Wireless, label);

        /// <summary>Adds a management connection.</summary>
        public VisioNetworkDiagramBuilder Management(string fromId, string toId, string? label = null) => Link(fromId, toId, VisioNetworkLinkKind.Management, label);

        /// <summary>Adds a link between two known network nodes.</summary>
        public VisioNetworkDiagramBuilder Link(string fromId, string toId, VisioNetworkLinkKind kind, string? label = null) {
            string normalizedFromId = RequireId(fromId, nameof(fromId), "From node id");
            string normalizedToId = RequireId(toId, nameof(toId), "To node id");
            EnsureKnownNode(normalizedFromId, nameof(fromId));
            EnsureKnownNode(normalizedToId, nameof(toId));
            if (!Enum.IsDefined(typeof(VisioNetworkLinkKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            _links.Add(new LinkItem(null, normalizedFromId, normalizedToId, kind, label));
            return this;
        }

        /// <summary>Adds a semantic callout connected to a known network node using a generated callout id.</summary>
        public VisioNetworkDiagramBuilder Callout(string targetId, string text, double pinX, double pinY, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            EnsureKnownNode(normalizedTargetId, nameof(targetId));
            return Callout(normalizedTargetId, CreateCalloutId(normalizedTargetId), text, pinX, pinY, configure);
        }

        /// <summary>Adds a semantic callout connected to a known network node.</summary>
        public VisioNetworkDiagramBuilder Callout(string targetId, string id, string text, double pinX, double pinY, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            string normalizedId = RequireId(id, nameof(id), "Callout id");
            EnsureKnownNode(normalizedTargetId, nameof(targetId));
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A network item with id '{normalizedId}' already exists.", nameof(id));
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

        /// <summary>Adds a semantic callout placed beside a known network node using a generated callout id.</summary>
        public VisioNetworkDiagramBuilder Callout(string targetId, string text, VisioSide placement, double gap = 0.35D, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            EnsureKnownNode(normalizedTargetId, nameof(targetId));
            return Callout(normalizedTargetId, CreateCalloutId(normalizedTargetId), text, placement, gap, configure);
        }

        /// <summary>Adds a semantic callout placed beside a known network node.</summary>
        public VisioNetworkDiagramBuilder Callout(string targetId, string id, string text, VisioSide placement, double gap = 0.35D, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            string normalizedId = RequireId(id, nameof(id), "Callout id");
            EnsureKnownNode(normalizedTargetId, nameof(targetId));
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A network item with id '{normalizedId}' already exists.", nameof(id));
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

        /// <summary>Imports zone records into the network diagram.</summary>
        public VisioNetworkDiagramBuilder Zones(IEnumerable<VisioNetworkZoneRecord> zones) {
            if (zones == null) {
                throw new ArgumentNullException(nameof(zones));
            }

            foreach (VisioNetworkZoneRecord zone in zones) {
                string normalizedId = RequireId(zone.Id, nameof(zones), "Zone id");
                if (IsIdInUse(normalizedId)) {
                    throw new ArgumentException($"A network item with id '{normalizedId}' already exists.", nameof(zones));
                }

                ValidateGridPosition(zone.Column, zone.Row);
                if (zone.ColumnSpan <= 0) throw new ArgumentOutOfRangeException(nameof(zones), "Column span must be positive.");
                if (zone.RowSpan <= 0) throw new ArgumentOutOfRangeException(nameof(zones), "Row span must be positive.");
                _zones.Add(new ZoneItem(normalizedId, zone.Text, zone.Column, zone.Row, zone.ColumnSpan, zone.RowSpan, zone.ShapeData, zone.HyperlinkAddress, zone.HyperlinkDescription));
            }

            return this;
        }

        /// <summary>Imports node records into the network diagram.</summary>
        public VisioNetworkDiagramBuilder Nodes(IEnumerable<VisioNetworkNodeRecord> nodes) {
            if (nodes == null) {
                throw new ArgumentNullException(nameof(nodes));
            }

            foreach (VisioNetworkNodeRecord node in nodes) {
                string normalizedId = RequireId(node.Id, nameof(nodes), "Node id");
                if (IsIdInUse(normalizedId)) {
                    throw new ArgumentException($"A network item with id '{normalizedId}' already exists.", nameof(nodes));
                }

                ValidateGridPosition(node.Column, node.Row);
                if (!Enum.IsDefined(typeof(VisioNetworkNodeKind), node.Kind)) {
                    throw new ArgumentOutOfRangeException(nameof(nodes));
                }

                NodeItem item = new(normalizedId, node.Text, node.Column, node.Row, node.Kind, node.ShapeData, node.HyperlinkAddress, node.HyperlinkDescription);
                _nodes.Add(item);
                _nodesById.Add(normalizedId, item);
            }

            return this;
        }

        /// <summary>Imports link records into the network diagram.</summary>
        public VisioNetworkDiagramBuilder Links(IEnumerable<VisioNetworkLinkRecord> links) {
            if (links == null) {
                throw new ArgumentNullException(nameof(links));
            }

            foreach (VisioNetworkLinkRecord link in links) {
                string normalizedId = RequireId(link.Id, nameof(links), "Link id");
                string normalizedFromId = RequireId(link.FromId, nameof(links), "From node id");
                string normalizedToId = RequireId(link.ToId, nameof(links), "To node id");
                if (IsIdInUse(normalizedId)) {
                    throw new ArgumentException($"A network item with id '{normalizedId}' already exists.", nameof(links));
                }

                EnsureKnownNode(normalizedFromId, nameof(links));
                EnsureKnownNode(normalizedToId, nameof(links));
                if (!Enum.IsDefined(typeof(VisioNetworkLinkKind), link.Kind)) {
                    throw new ArgumentOutOfRangeException(nameof(links));
                }

                _links.Add(new LinkItem(normalizedId, normalizedFromId, normalizedToId, link.Kind, link.Label, link.ShapeData, link.HyperlinkAddress, link.HyperlinkDescription));
                _linkIds.Add(normalizedId);
            }

            return this;
        }

        /// <summary>Imports callout records into the network diagram.</summary>
        public VisioNetworkDiagramBuilder Callouts(IEnumerable<VisioNetworkCalloutRecord> callouts) {
            if (callouts == null) {
                throw new ArgumentNullException(nameof(callouts));
            }

            foreach (VisioNetworkCalloutRecord callout in callouts) {
                Action<VisioCalloutOptions>? configure = callout.CreateOptionsConfigurator();
                if (callout.UsePlacement) {
                    Callout(callout.TargetId, callout.Id, callout.Text, callout.Placement, callout.Gap, configure);
                } else {
                    Callout(callout.TargetId, callout.Id, callout.Text, callout.PinX, callout.PinY, configure);
                }
            }

            return this;
        }

        /// <summary>Imports network zones, nodes, links, and callouts from simple data records.</summary>
        public VisioNetworkDiagramBuilder Import(
            IEnumerable<VisioNetworkZoneRecord> zones,
            IEnumerable<VisioNetworkNodeRecord> nodes,
            IEnumerable<VisioNetworkLinkRecord> links,
            IEnumerable<VisioNetworkCalloutRecord>? callouts = null) {
            Zones(zones);
            Nodes(nodes);
            Links(links);
            if (callouts != null) {
                Callouts(callouts);
            }

            return this;
        }

        internal VisioPage Build() {
            if (_built) {
                throw new InvalidOperationException("This network diagram builder has already produced a page.");
            }

            _built = true;
            if (_nodes.Count == 0) {
                throw new InvalidOperationException("A network diagram requires at least one node.");
            }

            VisioPage page = _document.AddPage(_pageName, _pageWidth, _pageHeight, _unit);
            page.Grid(visible: false, snap: true);
            AddZones(page);
            AddNodes(page);
            AddLinks(page);
            AddCallouts(page);
            AddTitle(page);
            _document.RequestRecalcOnOpen();
            return page;
        }

        private void AddTitle(VisioPage page) {
            if (string.IsNullOrWhiteSpace(_titleText)) {
                return;
            }

            double y = _pageHeight - _topMargin - (_titleHeight / 2D);
            VisioShape title = page.AddTextBox(_titleId, _pageWidth / 2D, y, Math.Max(1D, _pageWidth - 1.6D), _titleHeight, _titleText, _unit);
            title.TextStyle = CreateTitleTextStyle();
            VisioSemanticUserCells.MarkGeneratedAdornment(title);
        }

        private VisioTextStyle CreateTitleTextStyle() => VisioDiagramTitleStyles.Create(_theme);

        private void AddZones(VisioPage page) {
            foreach (ZoneItem zone in _zones) {
                double width = (zone.ColumnSpan * _nodeWidth) + ((zone.ColumnSpan - 1) * _columnGap) + 0.8;
                double height = (zone.RowSpan * _nodeHeight) + ((zone.RowSpan - 1) * _rowGap) + 0.65;
                double x = GridX(zone.Column, zone.ColumnSpan);
                double y = GridY(zone.Row, zone.RowSpan);
                VisioShape shape = VisioNetworkDiagramVisuals.CreateBackgroundZone(
                    _document,
                    zone.Id,
                    x,
                    y,
                    width,
                    height,
                    string.Empty,
                    _theme,
                    _unit);
                page.Shapes.Add(shape);
                ApplyMetadata(shape, zone.ShapeData, zone.HyperlinkAddress, zone.HyperlinkDescription);
                VisioNetworkDiagramVisuals.AddBackgroundZoneCaption(
                    page,
                    CreateGeneratedAdornmentId(VisioNetworkDiagramVisuals.CreateBackgroundZoneCaptionId(zone.Id), page),
                    zone.Text,
                    x - width / 2D,
                    y + height / 2D,
                    width,
                    _theme);
            }
        }

        private void AddNodes(VisioPage page) {
            foreach (NodeItem node in _nodes) {
                VisioNetworkDiagramVisuals.GetNodeShape(node.Kind, _nodeWidth, _nodeHeight, out string masterNameU, out double width, out double height);
                VisioShape shape = page.AddStencilShape(
                    VisioStencils.Network,
                    VisioNetworkDiagramVisuals.GetNodeStencilId(node.Kind),
                    node.Id,
                    GridX(node.Column, 1),
                    GridY(node.Row, 1),
                    width,
                    height,
                    node.Text);
                shape.NameU = masterNameU;
                VisioNetworkDiagramVisuals.GetNodeStyle(_theme, node.Kind).ApplyTo(shape);
                ApplyMetadata(shape, node.ShapeData, node.HyperlinkAddress, node.HyperlinkDescription);
                node.Shape = shape;
            }
        }

        private void AddLinks(VisioPage page) {
            int routeIndex = 0;
            HashSet<string> reservedConnectorIds = BuildReservedConnectorIds(page);
            foreach (LinkItem link in _links) {
                NodeItem from = _nodesById[link.FromId];
                NodeItem to = _nodesById[link.ToId];
                if (from.Shape == null || to.Shape == null) {
                    throw new InvalidOperationException("Nodes must be placed before links are created.");
                }

                VisioNetworkDiagramVisuals.ResolveSides(from.Shape, to.Shape, out VisioSide fromSide, out VisioSide toSide);
                string connectorId = link.Id ?? ReserveGeneratedConnectorId(reservedConnectorIds);
                VisioConnector connector = page.AddConnector(connectorId, from.Shape, to.Shape, ConnectorKind.RightAngle, fromSide, toSide);

                VisioNetworkDiagramVisuals.GetConnectorStyle(_theme, link.Kind).ApplyTo(connector);
                connector.Label = link.Label;
                ApplyMetadata(connector, link.ShapeData, link.HyperlinkAddress, link.HyperlinkDescription);
                connector.RouteOrthogonal(offset: (routeIndex % 4) * 0.06);
                if (!string.IsNullOrWhiteSpace(link.Label)) {
                    connector.PlaceLabel(0.5, offsetY: 0.15);
                }

                routeIndex++;
            }
        }

        private HashSet<string> BuildReservedConnectorIds(VisioPage page) {
            HashSet<string> ids = new(StringComparer.OrdinalIgnoreCase);
            foreach (VisioShape shape in page.Shapes) {
                ReserveShapeIds(shape, ids);
            }

            foreach (VisioConnector connector in page.Connectors) {
                ids.Add(connector.Id);
            }

            foreach (LinkItem link in _links) {
                if (link.Id != null) {
                    ids.Add(link.Id);
                }
            }

            return ids;
        }

        private static void ReserveShapeIds(VisioShape shape, HashSet<string> ids) {
            ids.Add(shape.Id);
            foreach (VisioShape child in shape.Children) {
                ReserveShapeIds(child, ids);
            }
        }

        private static string ReserveGeneratedConnectorId(HashSet<string> reservedIds) {
            int nextId = 1;
            while (true) {
                string candidate = nextId.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (reservedIds.Add(candidate)) {
                    return candidate;
                }

                nextId++;
            }
        }

        private void AddCallouts(VisioPage page) {
            foreach (CalloutItem callout in _callouts) {
                NodeItem target = _nodesById[callout.TargetId];
                if (target.Shape == null) {
                    throw new InvalidOperationException("Nodes must be placed before callouts are created.");
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

        private double GridX(int column, int span) {
            double left = _leftMargin + column * (_nodeWidth + _columnGap);
            double width = span * _nodeWidth + (span - 1) * _columnGap;
            return left + width / 2D;
        }

        private double GridY(int row, int span) {
            double top = _pageHeight - _topMargin - HeaderHeight - row * (_nodeHeight + _rowGap);
            double height = span * _nodeHeight + (span - 1) * _rowGap;
            return top - height / 2D;
        }

        private double HeaderHeight {
            get {
                double height = string.IsNullOrWhiteSpace(_titleText) ? 0D : _titleHeight + _titleGap;
                if (_zones.Any(zone => !string.IsNullOrWhiteSpace(zone.Text))) {
                    height += VisioNetworkDiagramVisuals.BackgroundZoneCaptionHeaderClearance;
                }

                return height;
            }
        }

        private void EnsureKnownNode(string id, string parameterName) {
            string normalizedId = RequireId(id, parameterName, "Network node id");
            if (!_nodesById.ContainsKey(normalizedId)) {
                throw new ArgumentException($"Unknown network node id '{normalizedId}'.", parameterName);
            }
        }

        private static IReadOnlyDictionary<string, string?> CopyShapeData(IDictionary<string, string?>? shapeData) {
            if (shapeData == null || shapeData.Count == 0) {
                return new Dictionary<string, string?>(StringComparer.Ordinal);
            }

            return shapeData.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal);
        }

        private static void ApplyMetadata(VisioShape shape, IReadOnlyDictionary<string, string?> shapeData, string? hyperlinkAddress, string? hyperlinkDescription) {
            foreach (KeyValuePair<string, string?> pair in shapeData) {
                if (!string.IsNullOrWhiteSpace(pair.Key)) {
                    shape.SetShapeData(pair.Key, pair.Value);
                }
            }

            string? address = hyperlinkAddress;
            if (!string.IsNullOrWhiteSpace(address)) {
                shape.AddHyperlink(address!, hyperlinkDescription);
            }
        }

        private static void ApplyMetadata(VisioConnector connector, IReadOnlyDictionary<string, string?> shapeData, string? hyperlinkAddress, string? hyperlinkDescription) {
            foreach (KeyValuePair<string, string?> pair in shapeData) {
                if (!string.IsNullOrWhiteSpace(pair.Key)) {
                    connector.SetShapeData(pair.Key, pair.Value);
                }
            }

            string? address = hyperlinkAddress;
            if (!string.IsNullOrWhiteSpace(address)) {
                connector.AddHyperlink(address!, hyperlinkDescription);
            }
        }

        private static string RequireId(string id, string parameterName, string label) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException(label + " cannot be null or whitespace.", parameterName);
            }

            return id.Trim();
        }

        private bool IsIdInUse(string id) {
            if (!string.IsNullOrWhiteSpace(_titleText) && string.Equals(_titleId, id, StringComparison.Ordinal)) {
                return true;
            }

            if (_nodesById.ContainsKey(id)) {
                return true;
            }

            foreach (ZoneItem zone in _zones) {
                if (string.Equals(zone.Id, id, StringComparison.Ordinal)) {
                    return true;
                }
            }

            if (_linkIds.Contains(id)) {
                return true;
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

        private string CreateGeneratedAdornmentId(string baseId, VisioPage page) {
            string id = baseId;
            int index = 2;
            while (IsIdInUse(id) || page.Shapes.Any(shape => string.Equals(shape.Id, id, StringComparison.Ordinal))) {
                id = baseId + "-" + index;
                index++;
            }

            return id;
        }

        private static void ValidateGridPosition(int column, int row) {
            if (column < 0) throw new ArgumentOutOfRangeException(nameof(column), "Column must be zero or greater.");
            if (row < 0) throw new ArgumentOutOfRangeException(nameof(row), "Row must be zero or greater.");
        }

        private static void ValidateFinite(double value, string parameterName) {
            if (double.IsNaN(value) || double.IsInfinity(value)) {
                throw new ArgumentOutOfRangeException(parameterName, "Value must be a finite number.");
            }
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

        private static void ValidatePlacement(VisioSide placement, string parameterName) {
            if (placement == VisioSide.Auto || !Enum.IsDefined(typeof(VisioSide), placement)) {
                throw new ArgumentOutOfRangeException(parameterName, "Placement must be Left, Right, Bottom, or Top.");
            }
        }

    }
}
