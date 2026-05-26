using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level builder for dependency diagrams where OfficeIMO lays out nodes from directed relationships.
    /// </summary>
    public sealed class VisioDependencyDiagramBuilder {
        private sealed class NodeItem {
            public NodeItem(string id, string text, VisioDependencyNodeKind kind) {
                Id = id;
                Text = text;
                Kind = kind;
            }

            public string Id { get; }

            public string Text { get; }

            public VisioDependencyNodeKind Kind { get; }

            public int Layer { get; set; }

            public int Row { get; set; }

            public VisioShape? Shape { get; set; }
        }

        private sealed class DependencyItem {
            public DependencyItem(string fromId, string toId, VisioDependencyConnectorKind kind, string? label) {
                FromId = fromId;
                ToId = toId;
                Kind = kind;
                Label = label;
            }

            public string FromId { get; }

            public string ToId { get; }

            public VisioDependencyConnectorKind Kind { get; }

            public string? Label { get; }
        }

        private readonly VisioDocument _document;
        private readonly string _pageName;
        private readonly List<NodeItem> _nodes = new();
        private readonly Dictionary<string, NodeItem> _nodesById = new(StringComparer.Ordinal);
        private readonly List<DependencyItem> _dependencies = new();
        private VisioStyleTheme _theme = VisioStyleTheme.Technical();
        private VisioMeasurementUnit _unit = VisioMeasurementUnit.Inches;
        private double _pageWidth = 11;
        private double _pageHeight = 8.5;
        private double _leftMargin = 0.8;
        private double _topMargin = 0.8;
        private double _rightMargin = 0.8;
        private double _bottomMargin = 0.8;
        private double _nodeWidth = 1.8;
        private double _nodeHeight = 0.85;
        private double _columnGap = 1.15;
        private double _rowGap = 0.55;
        private bool _fitPageToGraph = true;
        private bool _built;

        internal VisioDependencyDiagramBuilder(VisioDocument document, string pageName) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _pageName = string.IsNullOrWhiteSpace(pageName) ? "Dependency Diagram" : pageName;
        }

        /// <summary>Sets the page size used by the generated dependency diagram page.</summary>
        public VisioDependencyDiagramBuilder PageSize(double width, double height, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _pageWidth = width;
            _pageHeight = height;
            _unit = unit;
            return this;
        }

        /// <summary>Sets whether the builder can grow the page to fit the graph. Enabled by default.</summary>
        public VisioDependencyDiagramBuilder FitPageToGraph(bool enabled = true) {
            _fitPageToGraph = enabled;
            return this;
        }

        /// <summary>Sets the visual theme.</summary>
        public VisioDependencyDiagramBuilder Theme(VisioStyleTheme theme) {
            _theme = (theme ?? throw new ArgumentNullException(nameof(theme))).Clone();
            return this;
        }

        /// <summary>Sets outer page margins used by the automatic layout.</summary>
        public VisioDependencyDiagramBuilder Margins(double left, double top, double right = 0.8, double bottom = 0.8) {
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

        /// <summary>Sets default node size.</summary>
        public VisioDependencyDiagramBuilder NodeSize(double width, double height) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _nodeWidth = width;
            _nodeHeight = height;
            return this;
        }

        /// <summary>Sets spacing between automatic layers and rows.</summary>
        public VisioDependencyDiagramBuilder Spacing(double columnGap, double rowGap) {
            ValidateNonNegative(columnGap, nameof(columnGap));
            ValidateNonNegative(rowGap, nameof(rowGap));
            _columnGap = columnGap;
            _rowGap = rowGap;
            return this;
        }

        /// <summary>Adds a component node.</summary>
        public VisioDependencyDiagramBuilder Component(string id, string text) => Node(id, text, VisioDependencyNodeKind.Component);

        /// <summary>Adds a data node.</summary>
        public VisioDependencyDiagramBuilder Data(string id, string text) => Node(id, text, VisioDependencyNodeKind.Data);

        /// <summary>Adds an external actor or system node.</summary>
        public VisioDependencyDiagramBuilder External(string id, string text) => Node(id, text, VisioDependencyNodeKind.External);

        /// <summary>Adds a decision or policy node.</summary>
        public VisioDependencyDiagramBuilder Decision(string id, string text) => Node(id, text, VisioDependencyNodeKind.Decision);

        /// <summary>Adds a dependency node.</summary>
        public VisioDependencyDiagramBuilder Node(string id, string text, VisioDependencyNodeKind kind = VisioDependencyNodeKind.Component) {
            string normalizedId = RequireId(id, nameof(id), "Node id");
            if (_nodesById.ContainsKey(normalizedId)) {
                throw new ArgumentException($"A dependency node with id '{normalizedId}' already exists.", nameof(id));
            }

            if (!Enum.IsDefined(typeof(VisioDependencyNodeKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            NodeItem node = new(normalizedId, text ?? string.Empty, kind);
            _nodes.Add(node);
            _nodesById.Add(normalizedId, node);
            return this;
        }

        /// <summary>Adds a standard dependency connector.</summary>
        public VisioDependencyDiagramBuilder DependsOn(string fromId, string toId, string? label = null) =>
            Dependency(fromId, toId, VisioDependencyConnectorKind.Dependency, label);

        /// <summary>Adds a data dependency connector.</summary>
        public VisioDependencyDiagramBuilder DataDependency(string fromId, string toId, string? label = null) =>
            Dependency(fromId, toId, VisioDependencyConnectorKind.Data, label);

        /// <summary>Adds a control/policy dependency connector.</summary>
        public VisioDependencyDiagramBuilder ControlDependency(string fromId, string toId, string? label = null) =>
            Dependency(fromId, toId, VisioDependencyConnectorKind.Control, label);

        /// <summary>Adds a dependency connector between two known nodes.</summary>
        public VisioDependencyDiagramBuilder Dependency(string fromId, string toId, VisioDependencyConnectorKind kind, string? label = null) {
            EnsureKnownNode(fromId, nameof(fromId));
            EnsureKnownNode(toId, nameof(toId));
            if (!Enum.IsDefined(typeof(VisioDependencyConnectorKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            _dependencies.Add(new DependencyItem(fromId, toId, kind, label));
            return this;
        }

        internal VisioPage Build() {
            if (_built) {
                throw new InvalidOperationException("This dependency diagram builder has already produced a page.");
            }

            _built = true;
            if (_nodes.Count == 0) {
                throw new InvalidOperationException("A dependency diagram requires at least one node.");
            }

            AssignLayers();
            SizePageForLayout();

            VisioPage page = _document.AddPage(_pageName, _pageWidth, _pageHeight, _unit);
            page.Grid(visible: false, snap: true);
            AddNodes(page);
            AddDependencies(page);
            _document.RequestRecalcOnOpen();
            return page;
        }

        private void AssignLayers() {
            Dictionary<string, int> indegree = _nodes.ToDictionary(node => node.Id, _ => 0, StringComparer.Ordinal);
            Dictionary<string, List<string>> outgoing = _nodes.ToDictionary(node => node.Id, _ => new List<string>(), StringComparer.Ordinal);
            foreach (DependencyItem dependency in _dependencies) {
                outgoing[dependency.FromId].Add(dependency.ToId);
                indegree[dependency.ToId]++;
            }

            Queue<NodeItem> ready = new(_nodes.Where(node => indegree[node.Id] == 0));
            List<NodeItem> ordered = new();
            while (ready.Count > 0) {
                NodeItem node = ready.Dequeue();
                ordered.Add(node);
                foreach (string targetId in outgoing[node.Id]) {
                    NodeItem target = _nodesById[targetId];
                    target.Layer = Math.Max(target.Layer, node.Layer + 1);
                    indegree[targetId]--;
                    if (indegree[targetId] == 0) {
                        ready.Enqueue(target);
                    }
                }
            }

            if (ordered.Count != _nodes.Count) {
                throw new InvalidOperationException("Dependency diagram contains a cycle. Automatic layered layout requires an acyclic graph.");
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
            if (!_fitPageToGraph) {
                return;
            }

            int layerCount = _nodes.Max(node => node.Layer) + 1;
            int rowCount = _nodes.GroupBy(node => node.Layer).Max(group => group.Count());
            double requiredWidth = _leftMargin + _rightMargin + (layerCount * _nodeWidth) + Math.Max(0, layerCount - 1) * _columnGap;
            double requiredHeight = _topMargin + _bottomMargin + (rowCount * _nodeHeight) + Math.Max(0, rowCount - 1) * _rowGap;
            _pageWidth = Math.Max(_pageWidth, requiredWidth);
            _pageHeight = Math.Max(_pageHeight, requiredHeight);
        }

        private void AddNodes(VisioPage page) {
            foreach (NodeItem node in _nodes) {
                GetNodeShape(node.Kind, out string masterNameU, out double width, out double height);
                VisioShape shape = new(node.Id, XForLayer(node.Layer), YForRow(node.Layer, node.Row), width, height, node.Text) {
                    NameU = masterNameU,
                    Master = _document.EnsureBuiltinMaster(masterNameU)
                };
                GetNodeStyle(node.Kind).ApplyTo(shape);
                node.Shape = shape;
                page.Shapes.Add(shape);
            }
        }

        private void AddDependencies(VisioPage page) {
            int routeIndex = 0;
            foreach (DependencyItem dependency in _dependencies) {
                NodeItem from = _nodesById[dependency.FromId];
                NodeItem to = _nodesById[dependency.ToId];
                if (from.Shape == null || to.Shape == null) {
                    throw new InvalidOperationException("Nodes must be placed before dependency connectors are created.");
                }

                ResolveSides(from.Shape, to.Shape, out VisioSide fromSide, out VisioSide toSide);
                VisioConnector connector = page.AddConnector(from.Shape, to.Shape, ConnectorKind.RightAngle, fromSide, toSide);
                GetConnectorStyle(dependency.Kind).ApplyTo(connector);
                connector.Label = dependency.Label;
                connector.RouteOrthogonal(offset: (routeIndex % 5) * 0.06);
                if (!string.IsNullOrWhiteSpace(dependency.Label)) {
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
            int rowCount = _nodes.Count(node => node.Layer == layer);
            double contentHeight = rowCount * _nodeHeight + Math.Max(0, rowCount - 1) * _rowGap;
            double top = _pageHeight - _topMargin;
            double layerTop = top - Math.Max(0D, ((_pageHeight - _topMargin - _bottomMargin) - contentHeight) / 2D);
            return layerTop - (_nodeHeight / 2D) - row * (_nodeHeight + _rowGap);
        }

        private void GetNodeShape(VisioDependencyNodeKind kind, out string masterNameU, out double width, out double height) {
            width = _nodeWidth;
            height = _nodeHeight;
            switch (kind) {
                case VisioDependencyNodeKind.Data:
                    masterNameU = "Data";
                    break;
                case VisioDependencyNodeKind.External:
                    masterNameU = "Circle";
                    width = Math.Min(_nodeWidth, _nodeHeight * 1.15);
                    height = width;
                    break;
                case VisioDependencyNodeKind.Decision:
                    masterNameU = "Decision";
                    height = _nodeHeight * 1.2;
                    break;
                default:
                    masterNameU = "Process";
                    break;
            }
        }

        private VisioShapeStyle GetNodeStyle(VisioDependencyNodeKind kind) {
            switch (kind) {
                case VisioDependencyNodeKind.Data:
                    return _theme.Marker;
                case VisioDependencyNodeKind.External:
                    return _theme.Success;
                case VisioDependencyNodeKind.Decision:
                    return _theme.Decision;
                default:
                    return _theme.Primary;
            }
        }

        private VisioConnectorStyle GetConnectorStyle(VisioDependencyConnectorKind kind) {
            switch (kind) {
                case VisioDependencyConnectorKind.Data:
                    return _theme.DataConnector;
                case VisioDependencyConnectorKind.Control:
                    return _theme.ControlConnector;
                default:
                    return _theme.Connector;
            }
        }

        private static void ResolveSides(VisioShape from, VisioShape to, out VisioSide fromSide, out VisioSide toSide) {
            if (Math.Abs(from.PinX - to.PinX) >= Math.Abs(from.PinY - to.PinY)) {
                bool toRight = to.PinX >= from.PinX;
                fromSide = toRight ? VisioSide.Right : VisioSide.Left;
                toSide = toRight ? VisioSide.Left : VisioSide.Right;
                return;
            }

            bool toAbove = to.PinY >= from.PinY;
            fromSide = toAbove ? VisioSide.Top : VisioSide.Bottom;
            toSide = toAbove ? VisioSide.Bottom : VisioSide.Top;
        }

        private void EnsureKnownNode(string id, string parameterName) {
            string normalizedId = RequireId(id, parameterName, "Node id");
            if (!_nodesById.ContainsKey(normalizedId)) {
                throw new ArgumentException($"Unknown dependency node id '{normalizedId}'.", parameterName);
            }
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
