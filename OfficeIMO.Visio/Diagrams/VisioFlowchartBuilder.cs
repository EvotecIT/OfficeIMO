using System;
using System.Collections.Generic;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level flowchart authoring helper that creates normal Visio pages,
    /// shapes, and connectors from semantic nodes.
    /// </summary>
    public sealed class VisioFlowchartBuilder {
        private sealed class Node {
            public Node(string id, string text, VisioFlowchartNodeKind kind) {
                Id = id;
                Text = text;
                Kind = kind;
            }

            public string Id { get; }

            public string Text { get; }

            public VisioFlowchartNodeKind Kind { get; }

            public VisioShape? Shape { get; set; }
        }

        private sealed class Edge {
            public Edge(string fromId, string toId, string? label, bool automatic) {
                FromId = fromId;
                ToId = toId;
                Label = label;
                Automatic = automatic;
            }

            public string FromId { get; }

            public string ToId { get; }

            public string? Label { get; }

            public bool Automatic { get; }
        }

        private readonly List<Node> _nodes = new List<Node>();
        private readonly Dictionary<string, Node> _nodesById = new Dictionary<string, Node>(StringComparer.Ordinal);
        private readonly List<Edge> _edges = new List<Edge>();
        private readonly VisioDocument _document;
        private readonly string _pageName;
        private VisioFlowchartTheme _theme = VisioFlowchartTheme.ModernBlueGreen();
        private VisioFlowchartLayout _layout = VisioFlowchartLayout.Vertical;
        private VisioMeasurementUnit _unit = VisioMeasurementUnit.Inches;
        private double _pageWidth = 8.5;
        private double _pageHeight = 11;
        private double _topMargin = 0.75;
        private double _bottomMargin = 0.75;
        private double _verticalGap = 0.55;
        private bool _routeBranches = true;
        private double _branchLaneSpacing = 0.45;
        private bool _built;

        internal VisioFlowchartBuilder(VisioDocument document, string pageName) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _pageName = string.IsNullOrWhiteSpace(pageName) ? "Flowchart" : pageName;
        }

        /// <summary>Sets the page size used by the generated flowchart page.</summary>
        public VisioFlowchartBuilder PageSize(double width, double height, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _pageWidth = width;
            _pageHeight = height;
            _unit = unit;
            return this;
        }

        /// <summary>Sets the visual theme used for generated shapes and connectors.</summary>
        public VisioFlowchartBuilder Theme(VisioFlowchartTheme theme) {
            _theme = (theme ?? throw new ArgumentNullException(nameof(theme))).Clone();
            return this;
        }

        /// <summary>Sets the visual theme from a reusable OfficeIMO Visio style theme.</summary>
        public VisioFlowchartBuilder Theme(VisioStyleTheme theme) {
            if (theme == null) {
                throw new ArgumentNullException(nameof(theme));
            }

            return Theme(theme.ToFlowchartTheme());
        }

        /// <summary>Sets the deterministic layout strategy.</summary>
        public VisioFlowchartBuilder Layout(VisioFlowchartLayout layout) {
            if (!Enum.IsDefined(typeof(VisioFlowchartLayout), layout)) {
                throw new ArgumentOutOfRangeException(nameof(layout));
            }

            _layout = layout;
            return this;
        }

        /// <summary>Sets vertical spacing between generated nodes.</summary>
        public VisioFlowchartBuilder Spacing(double verticalGap) {
            ValidateNonNegative(verticalGap, nameof(verticalGap));
            _verticalGap = verticalGap;
            return this;
        }

        /// <summary>
        /// Controls deterministic side-lane routing for explicit branch and loop connectors.
        /// </summary>
        /// <param name="enabled">Whether explicit non-linear connectors should be routed around the main flow.</param>
        /// <param name="laneSpacing">Distance from the connected shapes to the generated side lane.</param>
        public VisioFlowchartBuilder RouteBranches(bool enabled = true, double laneSpacing = 0.45) {
            ValidatePositive(laneSpacing, nameof(laneSpacing));
            _routeBranches = enabled;
            _branchLaneSpacing = laneSpacing;
            return this;
        }

        /// <summary>Adds a start node.</summary>
        public VisioFlowchartBuilder Start(string id, string text) => AddNode(id, text, VisioFlowchartNodeKind.Start, connectFromPrevious: true);

        /// <summary>Adds a process step.</summary>
        public VisioFlowchartBuilder Step(string id, string text) => AddNode(id, text, VisioFlowchartNodeKind.Process, connectFromPrevious: true);

        /// <summary>Adds a decision node.</summary>
        public VisioFlowchartBuilder Decision(string id, string text) => AddNode(id, text, VisioFlowchartNodeKind.Decision, connectFromPrevious: true);

        /// <summary>Adds an input/output data node.</summary>
        public VisioFlowchartBuilder Data(string id, string text) => AddNode(id, text, VisioFlowchartNodeKind.Data, connectFromPrevious: true);

        /// <summary>Adds an off-page reference marker.</summary>
        public VisioFlowchartBuilder OffPage(string id, string text) => AddNode(id, text, VisioFlowchartNodeKind.OffPageReference, connectFromPrevious: true);

        /// <summary>Adds a continuation marker, usually for a second column or page region.</summary>
        public VisioFlowchartBuilder Continue(string id, string text) => AddNode(id, text, VisioFlowchartNodeKind.Continuation, connectFromPrevious: false);

        /// <summary>Adds an end node.</summary>
        public VisioFlowchartBuilder End(string id, string text) => AddNode(id, text, VisioFlowchartNodeKind.End, connectFromPrevious: true);

        /// <summary>Adds an explicit connector between two nodes.</summary>
        public VisioFlowchartBuilder Connect(string fromId, string toId, string? label = null) {
            EnsureKnownNode(fromId, nameof(fromId));
            EnsureKnownNode(toId, nameof(toId));
            _edges.Add(new Edge(fromId, toId, label, automatic: false));
            return this;
        }

        /// <summary>Adds a labeled branch connector between two nodes.</summary>
        public VisioFlowchartBuilder Branch(string fromId, string label, string toId) => Connect(fromId, toId, label);

        internal VisioPage Build() {
            if (_built) {
                throw new InvalidOperationException("This flowchart builder has already produced a page.");
            }

            _built = true;
            if (_nodes.Count == 0) {
                throw new InvalidOperationException("A flowchart requires at least one node.");
            }

            bool previousMastersByDefault = _document.UseMastersByDefault;
            _document.UseMastersByDefault = true;
            try {
                VisioPage page = _document.AddPage(_pageName, _pageWidth, _pageHeight, _unit);
                page.Grid(visible: false, snap: true);
                PlaceNodes(page);
                ConnectNodes(page);
                _document.RequestRecalcOnOpen();
                return page;
            } finally {
                _document.UseMastersByDefault = previousMastersByDefault;
            }
        }

        private VisioFlowchartBuilder AddNode(string id, string text, VisioFlowchartNodeKind kind, bool connectFromPrevious) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Flowchart node id cannot be null or whitespace.", nameof(id));
            }

            if (_nodesById.ContainsKey(id)) {
                throw new ArgumentException($"A flowchart node with id '{id}' already exists.", nameof(id));
            }

            Node node = new Node(id, text ?? string.Empty, kind);
            if (connectFromPrevious && _nodes.Count > 0) {
                _edges.Add(new Edge(_nodes[_nodes.Count - 1].Id, id, null, automatic: true));
            }

            _nodes.Add(node);
            _nodesById.Add(id, node);
            return this;
        }

        private void PlaceNodes(VisioPage page) {
            if (_layout == VisioFlowchartLayout.TwoColumnContinuation && TryGetContinuationSplit(out int splitIndex)) {
                PlaceColumn(page, 0, splitIndex, _pageWidth * 0.28);
                PlaceColumn(page, splitIndex, _nodes.Count, _pageWidth * 0.72);
                return;
            }

            PlaceColumn(page, 0, _nodes.Count, _pageWidth / 2D);
        }

        private bool TryGetContinuationSplit(out int splitIndex) {
            for (int i = 0; i < _nodes.Count; i++) {
                if (_nodes[i].Kind == VisioFlowchartNodeKind.Continuation) {
                    splitIndex = i;
                    return i > 0;
                }
            }

            splitIndex = -1;
            return false;
        }

        private void PlaceColumn(VisioPage page, int startIndex, int endIndex, double x) {
            double y = _pageHeight - _topMargin;
            for (int i = startIndex; i < endIndex; i++) {
                Node node = _nodes[i];
                GetNodeSize(node.Kind, out double width, out double height);
                y -= height / 2D;
                node.Shape = CreateShape(page, node, x, y, width, height);
                y -= height / 2D + _verticalGap;
            }

            if (y < _bottomMargin) {
                page.Height = (_pageHeight + (_bottomMargin - y)).ToInches(_unit);
            }
        }

        private VisioShape CreateShape(VisioPage page, Node node, double x, double y, double width, double height) {
            string nameU;
            Color fill;
            Color stroke;
            VisioTextStyle? textStyle;
            switch (node.Kind) {
                case VisioFlowchartNodeKind.Start:
                case VisioFlowchartNodeKind.End:
                    nameU = "Ellipse";
                    fill = _theme.TerminatorFill;
                    stroke = _theme.TerminatorStroke;
                    textStyle = _theme.TerminatorTextStyle;
                    break;
                case VisioFlowchartNodeKind.Decision:
                    nameU = "Decision";
                    fill = _theme.DecisionFill;
                    stroke = _theme.DecisionStroke;
                    textStyle = _theme.DecisionTextStyle;
                    break;
                case VisioFlowchartNodeKind.Data:
                    nameU = "Data";
                    fill = _theme.ProcessFill;
                    stroke = _theme.ProcessStroke;
                    textStyle = _theme.ProcessTextStyle;
                    break;
                case VisioFlowchartNodeKind.OffPageReference:
                    nameU = "Off-page reference";
                    fill = _theme.MarkerFill;
                    stroke = _theme.MarkerStroke;
                    textStyle = _theme.MarkerTextStyle;
                    break;
                case VisioFlowchartNodeKind.Continuation:
                    nameU = "Circle";
                    fill = _theme.MarkerFill;
                    stroke = _theme.MarkerStroke;
                    textStyle = _theme.MarkerTextStyle;
                    break;
                default:
                    nameU = "Process";
                    fill = _theme.ProcessFill;
                    stroke = _theme.ProcessStroke;
                    textStyle = _theme.ProcessTextStyle;
                    break;
            }

            VisioShape shape = new VisioShape(node.Id, x, y, width, height, node.Text) { NameU = nameU };
            ApplyStyle(shape, fill, stroke);
            if (textStyle != null) {
                shape.TextStyle = textStyle.Clone();
            }

            shape.Master = _document.EnsureBuiltinMaster(nameU);
            page.Shapes.Add(shape);

            return shape;
        }

        private void ConnectNodes(VisioPage page) {
            int branchRouteIndex = 0;
            for (int i = 0; i < _edges.Count; i++) {
                Edge edge = _edges[i];
                Node from = _nodesById[edge.FromId];
                Node to = _nodesById[edge.ToId];
                if (from.Shape == null || to.Shape == null) {
                    throw new InvalidOperationException("Flowchart nodes must be placed before connectors are created.");
                }

                bool routeBranch = ShouldRouteBranch(edge, from, to);
                if (routeBranch) {
                    ResolveBranchSides(from.Shape, to.Shape, out VisioSide branchFromSide, out VisioSide branchToSide);
                    AddConnector(page, edge, from.Shape, to.Shape, branchFromSide, branchToSide, routeBranch, branchRouteIndex++);
                    continue;
                }

                ResolveSides(from.Shape, to.Shape, out VisioSide fromSide, out VisioSide toSide);
                AddConnector(page, edge, from.Shape, to.Shape, fromSide, toSide, routeBranch: false, branchRouteIndex: 0);
            }
        }

        private VisioConnector AddConnector(
            VisioPage page,
            Edge edge,
            VisioShape from,
            VisioShape to,
            VisioSide fromSide,
            VisioSide toSide,
            bool routeBranch,
            int branchRouteIndex) {
            VisioConnector connector = page.AddConnector(from, to, ConnectorKind.RightAngle, fromSide, toSide);
            connector.EndArrow = EndArrow.Triangle;
            connector.LineColor = _theme.ConnectorColor;
            connector.LineWeight = _theme.LineWeight;
            connector.Label = edge.Label;
            if (_theme.ConnectorTextStyle != null) {
                connector.TextStyle = _theme.ConnectorTextStyle.Clone();
            }

            if (routeBranch) {
                RouteBranchConnector(page, connector, branchRouteIndex);
            }

            if (!string.IsNullOrWhiteSpace(edge.Label)) {
                connector.PlaceLabel(0.5D, offsetY: 0.18D);
            }

            return connector;
        }

        private bool ShouldRouteBranch(Edge edge, Node from, Node to) {
            if (!_routeBranches || edge.Automatic) {
                return false;
            }

            int fromIndex = _nodes.IndexOf(from);
            int toIndex = _nodes.IndexOf(to);
            if (fromIndex < 0 || toIndex < 0) {
                return false;
            }

            if (Math.Abs(fromIndex - toIndex) > 1 || toIndex < fromIndex) {
                return true;
            }

            if (from.Shape != null && to.Shape != null) {
                double horizontalDistance = Math.Abs(to.Shape.PinX - from.Shape.PinX);
                return horizontalDistance > Math.Max(_theme.ProcessWidth, _theme.DecisionWidth) * 0.9D;
            }

            return false;
        }

        private void RouteBranchConnector(VisioPage page, VisioConnector connector, int branchRouteIndex) {
            VisioShapeBounds fromBounds = connector.From.GetShapeBounds();
            VisioShapeBounds toBounds = connector.To.GetShapeBounds();
            double routeOffset = (branchRouteIndex % 3) * (_branchLaneSpacing * 0.5D);

            if (fromBounds.Right < toBounds.Left || toBounds.Right < fromBounds.Left) {
                double laneX = fromBounds.Right < toBounds.Left
                    ? (fromBounds.Right + toBounds.Left) / 2D
                    : (toBounds.Right + fromBounds.Left) / 2D;
                connector.RouteThrough(
                    VisioConnectorWaypoint.At(laneX, fromBounds.CenterY),
                    VisioConnectorWaypoint.At(laneX, toBounds.CenterY));
                return;
            }

            bool routeLeft = toBounds.CenterY >= fromBounds.CenterY;
            double laneXCandidate = routeLeft
                ? Math.Min(fromBounds.Left, toBounds.Left) - _branchLaneSpacing - routeOffset
                : Math.Max(fromBounds.Right, toBounds.Right) + _branchLaneSpacing + routeOffset;

            if (laneXCandidate < _branchLaneSpacing) {
                laneXCandidate = Math.Max(fromBounds.Right, toBounds.Right) + _branchLaneSpacing + routeOffset;
            } else if (laneXCandidate > page.Width - _branchLaneSpacing) {
                laneXCandidate = Math.Min(fromBounds.Left, toBounds.Left) - _branchLaneSpacing - routeOffset;
            }

            connector.RouteThrough(
                VisioConnectorWaypoint.At(laneXCandidate, fromBounds.CenterY),
                VisioConnectorWaypoint.At(laneXCandidate, toBounds.CenterY));
        }

        private static void ResolveBranchSides(VisioShape from, VisioShape to, out VisioSide fromSide, out VisioSide toSide) {
            VisioShapeBounds fromBounds = from.GetShapeBounds();
            VisioShapeBounds toBounds = to.GetShapeBounds();
            if (fromBounds.Right < toBounds.Left) {
                fromSide = VisioSide.Right;
                toSide = VisioSide.Left;
                return;
            }

            if (toBounds.Right < fromBounds.Left) {
                fromSide = VisioSide.Left;
                toSide = VisioSide.Right;
                return;
            }

            bool routeLeft = toBounds.CenterY >= fromBounds.CenterY;
            fromSide = routeLeft ? VisioSide.Left : VisioSide.Right;
            toSide = routeLeft ? VisioSide.Left : VisioSide.Right;
        }

        private void GetNodeSize(VisioFlowchartNodeKind kind, out double width, out double height) {
            switch (kind) {
                case VisioFlowchartNodeKind.Decision:
                    width = _theme.DecisionWidth;
                    height = _theme.DecisionHeight;
                    break;
                case VisioFlowchartNodeKind.Start:
                case VisioFlowchartNodeKind.End:
                    width = _theme.TerminatorWidth;
                    height = _theme.TerminatorHeight;
                    break;
                case VisioFlowchartNodeKind.Continuation:
                    width = _theme.MarkerDiameter;
                    height = _theme.MarkerDiameter;
                    break;
                case VisioFlowchartNodeKind.OffPageReference:
                    width = _theme.MarkerDiameter * 1.35;
                    height = _theme.MarkerDiameter * 1.35;
                    break;
                default:
                    width = _theme.ProcessWidth;
                    height = _theme.ProcessHeight;
                    break;
            }
        }

        private void ApplyStyle(VisioShape shape, Color fill, Color stroke) {
            shape.FillColor = fill;
            shape.LineColor = stroke;
            shape.LineWeight = _theme.LineWeight;
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

        private void EnsureKnownNode(string id, string parameterName) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Flowchart node id cannot be null or whitespace.", parameterName);
            }

            if (!_nodesById.ContainsKey(id)) {
                throw new ArgumentException($"Unknown flowchart node id '{id}'.", parameterName);
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
    }
}
