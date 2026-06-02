using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level builder for generic graph diagrams where OfficeIMO lays out arbitrary nodes and edges.
    /// </summary>
    public sealed partial class VisioGraphDiagramBuilder {
        private readonly VisioDocument _document;
        private readonly string _pageName;
        private readonly List<NodeItem> _nodes = new();
        private readonly Dictionary<string, NodeItem> _nodesById = new(StringComparer.Ordinal);
        private readonly List<EdgeItem> _edges = new();
        private readonly Dictionary<string, EdgeItem> _edgesById = new(StringComparer.Ordinal);
        private readonly HashSet<string> _edgeIds = new(StringComparer.Ordinal);
        private readonly List<string> _rootIds = new();
        private readonly HashSet<string> _rootIdSet = new(StringComparer.Ordinal);
        private readonly List<ZoneItem> _zones = new();
        private readonly Dictionary<string, ZoneItem> _zonesById = new(StringComparer.Ordinal);
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
        private bool _showLegend;
        private string _legendTitle = "Legend";
        private bool _legendIncludeNodeKinds = true;
        private bool _legendIncludeConnectorKinds = true;
        private const double StencilCaptionBottomOverflow = 0.52D;
        private const double LegendTitleHeight = 0.24D;
        private const double LegendRowHeight = 0.32D;
        private const double LegendGap = 0.22D;
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

        /// <summary>Adds an automatic legend based on the graph node and connector kinds used by the diagram.</summary>
        public VisioGraphDiagramBuilder Legend(bool enabled = true, string title = "Legend", bool includeNodeKinds = true, bool includeConnectorKinds = true) {
            _showLegend = enabled;
            _legendTitle = string.IsNullOrWhiteSpace(title) ? "Legend" : title;
            _legendIncludeNodeKinds = includeNodeKinds;
            _legendIncludeConnectorKinds = includeConnectorKinds;
            return this;
        }

        /// <summary>Adds a background zone around graph nodes.</summary>
        public VisioGraphDiagramBuilder Zone(string id, string text, params string[] nodeIds) {
            string normalizedId = RequireId(id, nameof(id), "Zone id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A graph diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            IReadOnlyList<string> normalizedNodeIds = NormalizeZoneNodeIds(nodeIds, nameof(nodeIds), "Zone node id");
            ZoneItem zone = new(normalizedId, text ?? string.Empty, normalizedNodeIds);
            _zones.Add(zone);
            _zonesById.Add(normalizedId, zone);
            _zoneIds.Add(normalizedId);
            return this;
        }

        /// <summary>Adds a semantic cluster around graph nodes. Clusters render as graph background zones and can carry Shape Data and hyperlinks.</summary>
        public VisioGraphDiagramBuilder Cluster(string id, string text, params string[] nodeIds) =>
            Zone(id, text, nodeIds);

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

            NodeItem node = new(normalizedId, text ?? string.Empty, kind, null, null);
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

            NodeItem node = new(normalizedId, text ?? string.Empty, InferStencilNodeKind(stencil), stencil, null);
            _nodes.Add(node);
            _nodesById.Add(normalizedId, node);
            return this;
        }

        /// <summary>Adds a graph node backed by the first matching stencil in a catalog.</summary>
        public VisioGraphDiagramBuilder StencilNode(string id, string text, VisioStencilCatalog catalog, params string[] stencilQueries) {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));
            VisioStencilShape stencil = catalog.FindBest(stencilQueries);
            string normalizedId = RequireId(id, nameof(id), "Node id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A graph diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            NodeItem node = new(normalizedId, text ?? string.Empty, InferStencilNodeKind(stencil), stencil, catalog.Name);
            _nodes.Add(node);
            _nodesById.Add(normalizedId, node);
            return this;
        }

        /// <summary>Imports graph nodes from simple data records.</summary>
        public VisioGraphDiagramBuilder Nodes(IEnumerable<VisioGraphNodeRecord> nodes) {
            if (nodes == null) throw new ArgumentNullException(nameof(nodes));
            foreach (VisioGraphNodeRecord node in nodes) {
                AddNodeRecord(node);
            }

            return this;
        }

        /// <summary>Imports graph edges from simple data records.</summary>
        public VisioGraphDiagramBuilder Edges(IEnumerable<VisioGraphEdgeRecord> edges) {
            if (edges == null) throw new ArgumentNullException(nameof(edges));
            List<VisioGraphEdgeRecord> edgeRecords = edges.ToList();
            HashSet<string> reservedExplicitIds = ReserveExplicitEdgeRecordIds(edgeRecords);
            foreach (VisioGraphEdgeRecord edge in edgeRecords) {
                AddEdgeRecord(edge, reservedExplicitIds);
            }

            return this;
        }

        /// <summary>Imports graph nodes and edges from simple data records.</summary>
        public VisioGraphDiagramBuilder Import(IEnumerable<VisioGraphNodeRecord> nodes, IEnumerable<VisioGraphEdgeRecord> edges) {
            return Nodes(nodes).Edges(edges);
        }

        /// <summary>Imports graph clusters from simple data records.</summary>
        public VisioGraphDiagramBuilder Clusters(IEnumerable<VisioGraphClusterRecord> clusters) {
            if (clusters == null) throw new ArgumentNullException(nameof(clusters));
            foreach (VisioGraphClusterRecord cluster in clusters) {
                AddClusterRecord(cluster);
            }

            return this;
        }

        /// <summary>Imports graph nodes, edges, and clusters from simple data records.</summary>
        public VisioGraphDiagramBuilder Import(IEnumerable<VisioGraphNodeRecord> nodes, IEnumerable<VisioGraphEdgeRecord> edges, IEnumerable<VisioGraphClusterRecord> clusters) {
            return Nodes(nodes).Edges(edges).Clusters(clusters);
        }

        /// <summary>Adds or updates Shape Data metadata that will be written to a graph node.</summary>
        public VisioGraphDiagramBuilder NodeShapeData(string nodeId, string name, string? value, string? label = null, VisioShapeDataType? type = null, string? prompt = null, string? format = null) {
            string normalizedNodeId = RequireId(nodeId, nameof(nodeId), "Node id");
            NodeItem node = GetKnownNode(normalizedNodeId, nameof(nodeId));
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Shape data name cannot be null or whitespace.", nameof(name));
            }

            string normalizedName = name.Trim();
            node.ShapeData.RemoveAll(row => string.Equals(row.Name, normalizedName, StringComparison.OrdinalIgnoreCase));
            node.ShapeData.Add(new NodeShapeDataItem(normalizedName, value, label, type, prompt, format));
            return this;
        }

        /// <summary>Overrides the visual style for a graph node.</summary>
        public VisioGraphDiagramBuilder NodeStyle(string nodeId, VisioShapeStyle style) {
            if (style == null) throw new ArgumentNullException(nameof(style));
            string normalizedNodeId = RequireId(nodeId, nameof(nodeId), "Node id");
            NodeItem node = GetKnownNode(normalizedNodeId, nameof(nodeId));
            node.StyleOverride = style.Clone();
            return this;
        }

        /// <summary>Overrides the visual style for a graph node by editing a theme-derived style copy.</summary>
        public VisioGraphDiagramBuilder NodeStyle(string nodeId, Action<VisioShapeStyle> configure) {
            if (configure == null) throw new ArgumentNullException(nameof(configure));
            string normalizedNodeId = RequireId(nodeId, nameof(nodeId), "Node id");
            NodeItem node = GetKnownNode(normalizedNodeId, nameof(nodeId));
            VisioShapeStyle style = (node.StyleOverride ?? GetNodeStyle(node.Kind)).Clone();
            configure(style);
            node.StyleOverride = style;
            return this;
        }

        /// <summary>Adds a hyperlink that will be written to a graph node.</summary>
        public VisioGraphDiagramBuilder NodeHyperlink(string nodeId, string address, string? description = null, string? subAddress = null) {
            string normalizedNodeId = RequireId(nodeId, nameof(nodeId), "Node id");
            NodeItem node = GetKnownNode(normalizedNodeId, nameof(nodeId));
            if (string.IsNullOrWhiteSpace(address)) {
                throw new ArgumentException("Hyperlink address cannot be null or whitespace.", nameof(address));
            }

            node.Hyperlinks.Add(new VisioHyperlink(address, description, subAddress));
            return this;
        }

        /// <summary>Adds a hyperlink that will be written to a graph node.</summary>
        public VisioGraphDiagramBuilder NodeHyperlink(string nodeId, Uri address, string? description = null, string? subAddress = null) {
            if (address == null) {
                throw new ArgumentNullException(nameof(address));
            }

            return NodeHyperlink(nodeId, address.ToString(), description, subAddress);
        }

        /// <summary>Adds a standard directed graph edge.</summary>
        public VisioGraphDiagramBuilder Edge(string fromId, string toId, string? label = null) =>
            Edge(fromId, toId, VisioGraphConnectorKind.Standard, label, directed: true);

        /// <summary>Adds a named standard directed graph edge.</summary>
        public VisioGraphDiagramBuilder Edge(string edgeId, string fromId, string toId, string? label) =>
            Edge(edgeId, fromId, toId, VisioGraphConnectorKind.Standard, label, directed: true);

        /// <summary>Adds an unlabeled named standard directed graph edge.</summary>
        public VisioGraphDiagramBuilder NamedEdge(string edgeId, string fromId, string toId) =>
            Edge(edgeId, fromId, toId, VisioGraphConnectorKind.Standard, label: null, directed: true);

        /// <summary>Adds a named standard directed graph edge.</summary>
        public VisioGraphDiagramBuilder NamedEdge(string edgeId, string fromId, string toId, string? label) =>
            Edge(edgeId, fromId, toId, VisioGraphConnectorKind.Standard, label, directed: true);

        /// <summary>Adds a data-flow graph edge.</summary>
        public VisioGraphDiagramBuilder DataEdge(string fromId, string toId, string? label = null) =>
            Edge(fromId, toId, VisioGraphConnectorKind.Data, label, directed: true);

        /// <summary>Adds a named data-flow graph edge.</summary>
        public VisioGraphDiagramBuilder DataEdge(string edgeId, string fromId, string toId, string? label) =>
            Edge(edgeId, fromId, toId, VisioGraphConnectorKind.Data, label, directed: true);

        /// <summary>Adds an unlabeled named data-flow graph edge.</summary>
        public VisioGraphDiagramBuilder NamedDataEdge(string edgeId, string fromId, string toId) =>
            Edge(edgeId, fromId, toId, VisioGraphConnectorKind.Data, label: null, directed: true);

        /// <summary>Adds a named data-flow graph edge.</summary>
        public VisioGraphDiagramBuilder NamedDataEdge(string edgeId, string fromId, string toId, string? label) =>
            Edge(edgeId, fromId, toId, VisioGraphConnectorKind.Data, label, directed: true);

        /// <summary>Adds a control-flow graph edge.</summary>
        public VisioGraphDiagramBuilder ControlEdge(string fromId, string toId, string? label = null) =>
            Edge(fromId, toId, VisioGraphConnectorKind.Control, label, directed: true);

        /// <summary>Adds a named control-flow graph edge.</summary>
        public VisioGraphDiagramBuilder ControlEdge(string edgeId, string fromId, string toId, string? label) =>
            Edge(edgeId, fromId, toId, VisioGraphConnectorKind.Control, label, directed: true);

        /// <summary>Adds an unlabeled named control-flow graph edge.</summary>
        public VisioGraphDiagramBuilder NamedControlEdge(string edgeId, string fromId, string toId) =>
            Edge(edgeId, fromId, toId, VisioGraphConnectorKind.Control, label: null, directed: true);

        /// <summary>Adds a named control-flow graph edge.</summary>
        public VisioGraphDiagramBuilder NamedControlEdge(string edgeId, string fromId, string toId, string? label) =>
            Edge(edgeId, fromId, toId, VisioGraphConnectorKind.Control, label, directed: true);

        /// <summary>Adds an emphasized graph edge.</summary>
        public VisioGraphDiagramBuilder EmphasisEdge(string fromId, string toId, string? label = null) =>
            Edge(fromId, toId, VisioGraphConnectorKind.Emphasis, label, directed: true);

        /// <summary>Adds a named emphasized graph edge.</summary>
        public VisioGraphDiagramBuilder EmphasisEdge(string edgeId, string fromId, string toId, string? label) =>
            Edge(edgeId, fromId, toId, VisioGraphConnectorKind.Emphasis, label, directed: true);

        /// <summary>Adds an unlabeled named emphasized graph edge.</summary>
        public VisioGraphDiagramBuilder NamedEmphasisEdge(string edgeId, string fromId, string toId) =>
            Edge(edgeId, fromId, toId, VisioGraphConnectorKind.Emphasis, label: null, directed: true);

        /// <summary>Adds a named emphasized graph edge.</summary>
        public VisioGraphDiagramBuilder NamedEmphasisEdge(string edgeId, string fromId, string toId, string? label) =>
            Edge(edgeId, fromId, toId, VisioGraphConnectorKind.Emphasis, label, directed: true);

        /// <summary>Adds an undirected graph edge.</summary>
        public VisioGraphDiagramBuilder Relationship(string fromId, string toId, string? label = null) =>
            Edge(fromId, toId, VisioGraphConnectorKind.Standard, label, directed: false);

        /// <summary>Adds a named undirected graph edge.</summary>
        public VisioGraphDiagramBuilder Relationship(string edgeId, string fromId, string toId, string? label) =>
            Edge(edgeId, fromId, toId, VisioGraphConnectorKind.Standard, label, directed: false);

        /// <summary>Adds an unlabeled named undirected graph edge.</summary>
        public VisioGraphDiagramBuilder NamedRelationship(string edgeId, string fromId, string toId) =>
            Edge(edgeId, fromId, toId, VisioGraphConnectorKind.Standard, label: null, directed: false);

        /// <summary>Adds a named undirected graph edge.</summary>
        public VisioGraphDiagramBuilder NamedRelationship(string edgeId, string fromId, string toId, string? label) =>
            Edge(edgeId, fromId, toId, VisioGraphConnectorKind.Standard, label, directed: false);

        /// <summary>Adds a graph edge between two known nodes.</summary>
        public VisioGraphDiagramBuilder Edge(string fromId, string toId, VisioGraphConnectorKind kind, string? label = null, bool directed = true) {
            AddEdge(null, fromId, toId, kind, label, directed);
            return this;
        }

        /// <summary>Adds a named graph edge between two known nodes.</summary>
        public VisioGraphDiagramBuilder Edge(string edgeId, string fromId, string toId, VisioGraphConnectorKind kind, string? label = null, bool directed = true) {
            string normalizedEdgeId = RequireId(edgeId, nameof(edgeId), "Edge id");
            if (IsIdInUse(normalizedEdgeId)) {
                throw new ArgumentException($"A graph diagram item with id '{normalizedEdgeId}' already exists.", nameof(edgeId));
            }

            AddEdge(normalizedEdgeId, fromId, toId, kind, label, directed);
            return this;
        }

        /// <summary>Adds a hyperlink that will be written to a named graph edge connector.</summary>
        public VisioGraphDiagramBuilder EdgeHyperlink(string edgeId, string address, string? description = null, string? subAddress = null) {
            string normalizedEdgeId = RequireId(edgeId, nameof(edgeId), "Edge id");
            EdgeItem edge = GetKnownEdge(normalizedEdgeId, nameof(edgeId));
            if (string.IsNullOrWhiteSpace(address)) {
                throw new ArgumentException("Hyperlink address cannot be null or whitespace.", nameof(address));
            }

            edge.Hyperlinks.Add(new VisioHyperlink(address, description, subAddress));
            return this;
        }

        /// <summary>Adds a hyperlink that will be written to a named graph edge connector.</summary>
        public VisioGraphDiagramBuilder EdgeHyperlink(string edgeId, Uri address, string? description = null, string? subAddress = null) {
            if (address == null) {
                throw new ArgumentNullException(nameof(address));
            }

            return EdgeHyperlink(edgeId, address.ToString(), description, subAddress);
        }

        /// <summary>Adds or replaces Shape Data on a named graph edge connector.</summary>
        public VisioGraphDiagramBuilder EdgeShapeData(string edgeId, string name, string? value, string? label = null, VisioShapeDataType? type = null, string? prompt = null, string? format = null) {
            string normalizedEdgeId = RequireId(edgeId, nameof(edgeId), "Edge id");
            EdgeItem edge = GetKnownEdge(normalizedEdgeId, nameof(edgeId));
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Shape data name cannot be null or whitespace.", nameof(name));
            }

            string normalizedName = name.Trim();
            edge.ShapeData.RemoveAll(row => string.Equals(row.Name, normalizedName, StringComparison.OrdinalIgnoreCase));
            edge.ShapeData.Add(new NodeShapeDataItem(normalizedName, value, label, type, prompt, format));
            return this;
        }

        /// <summary>Overrides the visual style for a named graph edge connector.</summary>
        public VisioGraphDiagramBuilder EdgeStyle(string edgeId, VisioConnectorStyle style) {
            if (style == null) throw new ArgumentNullException(nameof(style));
            string normalizedEdgeId = RequireId(edgeId, nameof(edgeId), "Edge id");
            EdgeItem edge = GetKnownEdge(normalizedEdgeId, nameof(edgeId));
            edge.StyleOverride = style.Clone();
            return this;
        }

        /// <summary>Overrides the visual style for a named graph edge connector by editing a theme-derived style copy.</summary>
        public VisioGraphDiagramBuilder EdgeStyle(string edgeId, Action<VisioConnectorStyle> configure) {
            if (configure == null) throw new ArgumentNullException(nameof(configure));
            string normalizedEdgeId = RequireId(edgeId, nameof(edgeId), "Edge id");
            EdgeItem edge = GetKnownEdge(normalizedEdgeId, nameof(edgeId));
            VisioConnectorStyle style = (edge.StyleOverride ?? GetConnectorStyle(edge.Kind, edge.Directed)).Clone();
            configure(style);
            edge.StyleOverride = style;
            return this;
        }

        /// <summary>Adds or replaces Shape Data on a graph zone or cluster background.</summary>
        public VisioGraphDiagramBuilder ZoneShapeData(string zoneId, string name, string? value, string? label = null, VisioShapeDataType? type = null, string? prompt = null, string? format = null) {
            string normalizedZoneId = RequireId(zoneId, nameof(zoneId), "Zone id");
            ZoneItem zone = GetKnownZone(normalizedZoneId, nameof(zoneId));
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Shape data name cannot be null or whitespace.", nameof(name));
            }

            string normalizedName = name.Trim();
            zone.ShapeData.RemoveAll(row => string.Equals(row.Name, normalizedName, StringComparison.OrdinalIgnoreCase));
            zone.ShapeData.Add(new NodeShapeDataItem(normalizedName, value, label, type, prompt, format));
            return this;
        }

        /// <summary>Adds a hyperlink that will be written to a graph zone or cluster background.</summary>
        public VisioGraphDiagramBuilder ZoneHyperlink(string zoneId, string address, string? description = null, string? subAddress = null) {
            string normalizedZoneId = RequireId(zoneId, nameof(zoneId), "Zone id");
            ZoneItem zone = GetKnownZone(normalizedZoneId, nameof(zoneId));
            if (string.IsNullOrWhiteSpace(address)) {
                throw new ArgumentException("Hyperlink address cannot be null or whitespace.", nameof(address));
            }

            zone.Hyperlinks.Add(new VisioHyperlink(address, description, subAddress));
            return this;
        }

        /// <summary>Adds a hyperlink that will be written to a graph zone or cluster background.</summary>
        public VisioGraphDiagramBuilder ZoneHyperlink(string zoneId, Uri address, string? description = null, string? subAddress = null) {
            if (address == null) {
                throw new ArgumentNullException(nameof(address));
            }

            return ZoneHyperlink(zoneId, address.ToString(), description, subAddress);
        }

        private void AddNodeRecord(VisioGraphNodeRecord record) {
            if (record == null) throw new ArgumentNullException(nameof(record));
            if (record.Stencil != null) {
                StencilNode(record.Id, record.Text, record.Stencil);
            } else if (record.StencilCatalog != null && record.StencilQueries.Count > 0) {
                StencilNode(record.Id, record.Text, record.StencilCatalog, record.StencilQueries.ToArray());
            } else {
                Node(record.Id, record.Text, record.Kind);
            }

            foreach (KeyValuePair<string, string?> item in record.ShapeData) {
                NodeShapeData(record.Id, item.Key, item.Value);
            }

            if (!string.IsNullOrWhiteSpace(record.HyperlinkAddress)) {
                NodeHyperlink(record.Id, record.HyperlinkAddress!, record.HyperlinkDescription, record.HyperlinkSubAddress);
            }

            if (record.IsRoot) {
                Root(record.Id);
            }
        }

        private void AddEdgeRecord(VisioGraphEdgeRecord record, ISet<string>? reservedExplicitIds = null) {
            if (record == null) throw new ArgumentNullException(nameof(record));
            string edgeId = string.IsNullOrWhiteSpace(record.Id)
                ? CreateStableEdgeId(record.FromId, record.ToId, record.Kind, reservedExplicitIds)
                : RequireId(record.Id!, nameof(record.Id), "Edge id");
            Edge(edgeId, record.FromId, record.ToId, record.Kind, record.Label, record.Directed);
            foreach (KeyValuePair<string, string?> item in record.ShapeData) {
                EdgeShapeData(edgeId, item.Key, item.Value);
            }

            if (!string.IsNullOrWhiteSpace(record.HyperlinkAddress)) {
                EdgeHyperlink(edgeId, record.HyperlinkAddress!, record.HyperlinkDescription, record.HyperlinkSubAddress);
            }
        }

        private void AddClusterRecord(VisioGraphClusterRecord record) {
            if (record == null) throw new ArgumentNullException(nameof(record));
            IReadOnlyList<string> nodeIds = NormalizeZoneNodeIds(record.NodeIds, nameof(record.NodeIds), "Cluster node id");
            Cluster(record.Id, record.Text, nodeIds.ToArray());
            foreach (KeyValuePair<string, string?> item in record.ShapeData) {
                ZoneShapeData(record.Id, item.Key, item.Value);
            }

            if (!string.IsNullOrWhiteSpace(record.HyperlinkAddress)) {
                ZoneHyperlink(record.Id, record.HyperlinkAddress!, record.HyperlinkDescription, record.HyperlinkSubAddress);
            }
        }

        private void AddEdge(string? edgeId, string fromId, string toId, VisioGraphConnectorKind kind, string? label, bool directed) {
            string normalizedFromId = RequireId(fromId, nameof(fromId), "From node id");
            string normalizedToId = RequireId(toId, nameof(toId), "To node id");
            EnsureKnownNode(normalizedFromId, nameof(fromId));
            EnsureKnownNode(normalizedToId, nameof(toId));
            if (!Enum.IsDefined(typeof(VisioGraphConnectorKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            EdgeItem edge = new(edgeId, normalizedFromId, normalizedToId, kind, label, directed);
            _edges.Add(edge);
            if (edgeId != null) {
                _edgesById.Add(edgeId, edge);
                _edgeIds.Add(edgeId);
            }
        }

        private string CreateStableEdgeId(string fromId, string toId, VisioGraphConnectorKind kind, ISet<string>? reservedExplicitIds = null) {
            string baseId = SlugId(fromId) + "-" + SlugId(kind.ToString()) + "-" + SlugId(toId);
            if (string.IsNullOrWhiteSpace(baseId.Replace("-", string.Empty))) {
                baseId = "edge";
            }

            string candidate = baseId;
            int index = 2;
            while (IsIdInUse(candidate) || reservedExplicitIds?.Contains(candidate) == true) {
                candidate = baseId + "-" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
                index++;
            }

            return candidate;
        }

        private HashSet<string> ReserveExplicitEdgeRecordIds(IEnumerable<VisioGraphEdgeRecord> records) {
            HashSet<string> ids = new(StringComparer.Ordinal);
            foreach (VisioGraphEdgeRecord record in records) {
                if (record == null) {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(record.Id)) {
                    continue;
                }

                string normalizedId = RequireId(record.Id!, nameof(record.Id), "Edge id");
                if (IsIdInUse(normalizedId) || !ids.Add(normalizedId)) {
                    throw new ArgumentException($"A graph item with id '{normalizedId}' already exists.", nameof(records));
                }
            }

            return ids;
        }
    }
}
