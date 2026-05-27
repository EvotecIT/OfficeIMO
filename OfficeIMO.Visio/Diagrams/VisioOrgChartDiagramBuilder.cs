using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level builder for dependency-free organization charts with hierarchy,
    /// assistant placement, and team bands.
    /// </summary>
    public sealed class VisioOrgChartDiagramBuilder {
        private sealed class OrgNode {
            public OrgNode(string id, string name, string title, string? managerId, string? bandId, VisioOrgChartNodeKind kind) {
                Id = id;
                Name = name;
                Title = title;
                ManagerId = managerId;
                BandId = bandId;
                Kind = kind;
            }

            public string Id { get; }

            public string Name { get; }

            public string Title { get; }

            public string? ManagerId { get; }

            public string? BandId { get; }

            public VisioOrgChartNodeKind Kind { get; }

            public int Depth { get; set; }

            public double X { get; set; }

            public double Y { get; set; }

            public double Width { get; set; }

            public double Height { get; set; }

            public VisioShape? Shape { get; set; }
        }

        private sealed class TeamBandItem {
            public TeamBandItem(string id, string text, string managerId) {
                Id = id;
                Text = text;
                ManagerId = managerId;
            }

            public string Id { get; }

            public string Text { get; }

            public string ManagerId { get; }
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
        private readonly List<OrgNode> _nodes = new();
        private readonly Dictionary<string, OrgNode> _nodesById = new(StringComparer.Ordinal);
        private readonly List<TeamBandItem> _bands = new();
        private readonly Dictionary<string, TeamBandItem> _bandsById = new(StringComparer.Ordinal);
        private readonly List<CalloutItem> _callouts = new();
        private VisioStyleTheme _theme = VisioStyleTheme.Modern();
        private VisioMeasurementUnit _unit = VisioMeasurementUnit.Inches;
        private double _pageWidth = 14;
        private double _pageHeight = 8.5;
        private double _leftMargin = 0.7;
        private double _rightMargin = 0.7;
        private double _topMargin = 0.7;
        private double _bottomMargin = 0.7;
        private double _nodeWidth = 1.85;
        private double _nodeHeight = 0.82;
        private double _columnGap = 0.55;
        private double _levelGap = 0.82;
        private double _assistantGap = 0.35;
        private double _bandPadding = 0.28;
        private string? _titleText;
        private string _titleId = "title";
        private double _titleHeight = 0.45;
        private double _titleGap = 0.35;
        private bool _built;

        internal VisioOrgChartDiagramBuilder(VisioDocument document, string pageName) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _pageName = string.IsNullOrWhiteSpace(pageName) ? "Org Chart" : pageName;
        }

        /// <summary>Sets the page size used by the generated org chart page.</summary>
        public VisioOrgChartDiagramBuilder PageSize(double width, double height, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _pageWidth = width;
            _pageHeight = height;
            _unit = unit;
            return this;
        }

        /// <summary>Sets the visual theme.</summary>
        public VisioOrgChartDiagramBuilder Theme(VisioStyleTheme theme) {
            _theme = (theme ?? throw new ArgumentNullException(nameof(theme))).Clone();
            return this;
        }

        /// <summary>Sets outer page margins.</summary>
        public VisioOrgChartDiagramBuilder Margins(double left, double top, double right = 0.7, double bottom = 0.7) {
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

        /// <summary>Sets the default org chart card size.</summary>
        public VisioOrgChartDiagramBuilder NodeSize(double width, double height) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _nodeWidth = width;
            _nodeHeight = height;
            return this;
        }

        /// <summary>Sets spacing between org chart cards and reporting levels.</summary>
        public VisioOrgChartDiagramBuilder Spacing(double columnGap, double levelGap, double assistantGap = 0.35) {
            ValidateNonNegative(columnGap, nameof(columnGap));
            ValidateNonNegative(levelGap, nameof(levelGap));
            ValidateNonNegative(assistantGap, nameof(assistantGap));
            _columnGap = columnGap;
            _levelGap = levelGap;
            _assistantGap = assistantGap;
            return this;
        }

        /// <summary>Sets padding around generated team bands.</summary>
        public VisioOrgChartDiagramBuilder TeamBandPadding(double padding) {
            ValidateNonNegative(padding, nameof(padding));
            _bandPadding = padding;
            return this;
        }

        /// <summary>Adds a centered editable title above the generated organization chart.</summary>
        public VisioOrgChartDiagramBuilder Title(string? text = null, string id = "title", double height = 0.45, double gap = 0.35) {
            string normalizedId = RequireId(id, nameof(id), "Title id");
            if (IsShapeIdInUse(normalizedId)) {
                throw new ArgumentException($"An org chart item with shape id '{normalizedId}' already exists.", nameof(id));
            }

            ValidatePositive(height, nameof(height));
            ValidateNonNegative(gap, nameof(gap));
            _titleText = string.IsNullOrWhiteSpace(text) ? _pageName : text;
            _titleId = normalizedId;
            _titleHeight = height;
            _titleGap = gap;
            return this;
        }

        /// <summary>Adds the root executive node.</summary>
        public VisioOrgChartDiagramBuilder Root(string id, string name, string title = "") =>
            AddNode(id, name, title, managerId: null, bandId: null, VisioOrgChartNodeKind.Executive);

        /// <summary>Adds a manager node below another node.</summary>
        public VisioOrgChartDiagramBuilder Manager(string id, string name, string title, string managerId, string? bandId = null) =>
            AddNode(id, name, title, managerId, bandId, VisioOrgChartNodeKind.Manager);

        /// <summary>Adds a standard position below another node.</summary>
        public VisioOrgChartDiagramBuilder Position(string id, string name, string title, string managerId, string? bandId = null) =>
            AddNode(id, name, title, managerId, bandId, VisioOrgChartNodeKind.Position);

        /// <summary>Adds an assistant beside a manager.</summary>
        public VisioOrgChartDiagramBuilder Assistant(string id, string name, string title, string managerId) =>
            AddNode(id, name, title, managerId, bandId: null, VisioOrgChartNodeKind.Assistant);

        /// <summary>Adds an open position below another node.</summary>
        public VisioOrgChartDiagramBuilder Vacancy(string id, string text, string managerId, string? bandId = null) =>
            AddNode(id, text, string.Empty, managerId, bandId, VisioOrgChartNodeKind.Vacancy);

        /// <summary>Adds an external advisor, vendor, or partner role below another node.</summary>
        public VisioOrgChartDiagramBuilder External(string id, string name, string title, string managerId, string? bandId = null) =>
            AddNode(id, name, title, managerId, bandId, VisioOrgChartNodeKind.External);

        /// <summary>Adds a background band around positions tagged with the band id.</summary>
        public VisioOrgChartDiagramBuilder TeamBand(string id, string text, string managerId) {
            string normalizedId = RequireId(id, nameof(id), "Team band id");
            string normalizedManagerId = RequireId(managerId, nameof(managerId), "Manager node id");
            EnsureKnownNode(normalizedManagerId, nameof(managerId));
            if (_bandsById.ContainsKey(normalizedId)) {
                throw new ArgumentException($"An org chart team band with id '{normalizedId}' already exists.", nameof(id));
            }

            string shapeId = GetBandShapeId(normalizedId);
            if (IsShapeIdInUse(shapeId)) {
                throw new ArgumentException($"An org chart item with shape id '{shapeId}' already exists.", nameof(id));
            }

            TeamBandItem band = new(normalizedId, text ?? string.Empty, normalizedManagerId);
            _bands.Add(band);
            _bandsById.Add(normalizedId, band);
            return this;
        }

        /// <summary>Adds a semantic callout connected to a known org chart node using a generated callout id.</summary>
        public VisioOrgChartDiagramBuilder Callout(string targetId, string text, double pinX, double pinY, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            EnsureKnownNode(normalizedTargetId, nameof(targetId));
            return Callout(normalizedTargetId, CreateCalloutId(normalizedTargetId), text, pinX, pinY, configure);
        }

        /// <summary>Adds a semantic callout connected to a known org chart node.</summary>
        public VisioOrgChartDiagramBuilder Callout(string targetId, string id, string text, double pinX, double pinY, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            string normalizedId = RequireId(id, nameof(id), "Callout id");
            EnsureKnownNode(normalizedTargetId, nameof(targetId));
            if (IsShapeIdInUse(normalizedId)) {
                throw new ArgumentException($"An org chart item with shape id '{normalizedId}' already exists.", nameof(id));
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

        /// <summary>Adds a semantic callout placed beside a known org chart node using a generated callout id.</summary>
        public VisioOrgChartDiagramBuilder Callout(string targetId, string text, VisioSide placement, double gap = 0.35D, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            EnsureKnownNode(normalizedTargetId, nameof(targetId));
            return Callout(normalizedTargetId, CreateCalloutId(normalizedTargetId), text, placement, gap, configure);
        }

        /// <summary>Adds a semantic callout placed beside a known org chart node.</summary>
        public VisioOrgChartDiagramBuilder Callout(string targetId, string id, string text, VisioSide placement, double gap = 0.35D, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            string normalizedId = RequireId(id, nameof(id), "Callout id");
            EnsureKnownNode(normalizedTargetId, nameof(targetId));
            if (IsShapeIdInUse(normalizedId)) {
                throw new ArgumentException($"An org chart item with shape id '{normalizedId}' already exists.", nameof(id));
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
                throw new InvalidOperationException("This org chart builder has already produced a page.");
            }

            _built = true;
            OrgNode root = GetRootNode();
            ValidateReferences();
            Layout(root);

            VisioPage page = _document.AddPage(_pageName, _pageWidth, _pageHeight, _unit);
            page.Grid(visible: false, snap: true);
            AddTeamBands(page);
            AddNodes(page);
            AddReportingLines(page);
            AddCallouts(page);
            AddTitle(page);
            _document.RequestRecalcOnOpen();
            return page;
        }

        private VisioOrgChartDiagramBuilder AddNode(string id, string name, string title, string? managerId, string? bandId, VisioOrgChartNodeKind kind) {
            string normalizedId = RequireId(id, nameof(id), "Org chart node id");
            if (IsShapeIdInUse(normalizedId)) {
                throw new ArgumentException($"An org chart item with shape id '{normalizedId}' already exists.", nameof(id));
            }

            if (!Enum.IsDefined(typeof(VisioOrgChartNodeKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            if (kind == VisioOrgChartNodeKind.Executive && _nodes.Exists(node => node.Kind == VisioOrgChartNodeKind.Executive)) {
                throw new InvalidOperationException("An org chart can only have one root executive node.");
            }

            string? normalizedManagerId = null;
            if (kind != VisioOrgChartNodeKind.Executive) {
                normalizedManagerId = RequireId(managerId!, nameof(managerId), "Manager node id");
                EnsureKnownNode(normalizedManagerId, nameof(managerId));
            } else if (!string.IsNullOrWhiteSpace(managerId)) {
                throw new ArgumentException("The root executive node cannot have a manager.", nameof(managerId));
            }

            string? normalizedBandId = null;
            if (!string.IsNullOrWhiteSpace(bandId)) {
                normalizedBandId = RequireId(bandId!, nameof(bandId), "Team band id");
                EnsureKnownBand(normalizedBandId, nameof(bandId));
            }

            OrgNode node = new(normalizedId, name ?? string.Empty, title ?? string.Empty, normalizedManagerId, normalizedBandId, kind);
            _nodes.Add(node);
            _nodesById.Add(normalizedId, node);
            return this;
        }

        private OrgNode GetRootNode() {
            OrgNode? root = null;
            foreach (OrgNode node in _nodes) {
                if (node.Kind != VisioOrgChartNodeKind.Executive) {
                    continue;
                }

                if (root != null) {
                    throw new InvalidOperationException("An org chart can only have one root executive node.");
                }

                root = node;
            }

            if (root == null) {
                throw new InvalidOperationException("An org chart requires a root executive node.");
            }

            return root;
        }

        private void ValidateReferences() {
            foreach (TeamBandItem band in _bands) {
                EnsureKnownNode(band.ManagerId, nameof(band.ManagerId));
            }

            foreach (OrgNode node in _nodes) {
                if (!string.IsNullOrWhiteSpace(node.ManagerId)) {
                    EnsureKnownNode(node.ManagerId!, nameof(node.ManagerId));
                }

                if (!string.IsNullOrWhiteSpace(node.BandId)) {
                    EnsureKnownBand(node.BandId!, nameof(node.BandId));
                }
            }
        }

        private void Layout(OrgNode root) {
            int maxDepth = GetMaxDepth(root, new HashSet<string>(StringComparer.Ordinal));
            int leafCount = Math.Max(1, GetLeafCount(root, new HashSet<string>(StringComparer.Ordinal)));
            double treeWidth = (leafCount * _nodeWidth) + ((leafCount - 1) * _columnGap);
            _pageWidth = Math.Max(_pageWidth, _leftMargin + treeWidth + _rightMargin);
            _pageHeight = Math.Max(_pageHeight, _topMargin + HeaderHeight + ((maxDepth + 1) * _nodeHeight) + (maxDepth * _levelGap) + _bottomMargin);

            double nextLeafX = _leftMargin + (_nodeWidth / 2D);
            AssignTreePositions(root, 0, ref nextLeafX, new HashSet<string>(StringComparer.Ordinal));
            AssignAssistantPositions();
            ExpandPageToFitNodes();
        }

        private int GetLeafCount(OrgNode node, HashSet<string> visiting) {
            EnsureNoCycle(node, visiting);
            List<OrgNode> children = GetDirectReports(node.Id);
            if (children.Count == 0) {
                visiting.Remove(node.Id);
                return 1;
            }

            int count = 0;
            foreach (OrgNode child in children) {
                count += GetLeafCount(child, visiting);
            }

            visiting.Remove(node.Id);
            return count;
        }

        private int GetMaxDepth(OrgNode node, HashSet<string> visiting) {
            EnsureNoCycle(node, visiting);
            int maxDepth = 0;
            foreach (OrgNode child in GetDirectReports(node.Id)) {
                maxDepth = Math.Max(maxDepth, 1 + GetMaxDepth(child, visiting));
            }

            visiting.Remove(node.Id);
            return maxDepth;
        }

        private void AssignTreePositions(OrgNode node, int depth, ref double nextLeafX, HashSet<string> visiting) {
            EnsureNoCycle(node, visiting);
            List<OrgNode> children = GetDirectReports(node.Id);
            node.Depth = depth;
            GetNodeShape(node.Kind, out _, out double width, out double height);
            node.Width = width;
            node.Height = height;
            node.Y = LevelCenterY(depth);

            if (children.Count == 0) {
                node.X = nextLeafX;
                nextLeafX += _nodeWidth + _columnGap;
                visiting.Remove(node.Id);
                return;
            }

            double firstChildX = 0D;
            double lastChildX = 0D;
            for (int i = 0; i < children.Count; i++) {
                AssignTreePositions(children[i], depth + 1, ref nextLeafX, visiting);
                if (i == 0) {
                    firstChildX = children[i].X;
                }

                lastChildX = children[i].X;
            }

            node.X = (firstChildX + lastChildX) / 2D;
            visiting.Remove(node.Id);
        }

        private void AssignAssistantPositions() {
            Dictionary<string, int> assistantCountsByManager = new(StringComparer.Ordinal);
            foreach (OrgNode assistant in _nodes) {
                if (assistant.Kind != VisioOrgChartNodeKind.Assistant || string.IsNullOrWhiteSpace(assistant.ManagerId)) {
                    continue;
                }

                OrgNode manager = _nodesById[assistant.ManagerId!];
                GetNodeShape(assistant.Kind, out _, out double width, out double height);
                assistant.Width = width;
                assistant.Height = height;
                assistant.Depth = manager.Depth;
                assistantCountsByManager.TryGetValue(manager.Id, out int index);
                assistant.X = manager.X + (manager.Width / 2D) + _assistantGap + (assistant.Width / 2D);
                assistant.Y = manager.Y - (index * (assistant.Height + 0.18D));
                assistantCountsByManager[manager.Id] = index + 1;
            }
        }

        private void ExpandPageToFitNodes() {
            double maxRight = 0D;
            double minBottom = double.MaxValue;
            foreach (OrgNode node in _nodes) {
                maxRight = Math.Max(maxRight, node.X + (node.Width / 2D));
                minBottom = Math.Min(minBottom, node.Y - (node.Height / 2D));
            }

            _pageWidth = Math.Max(_pageWidth, maxRight + _rightMargin);
            if (minBottom < _bottomMargin) {
                double delta = _bottomMargin - minBottom;
                _pageHeight += delta;
                foreach (OrgNode node in _nodes) {
                    node.Y += delta;
                }
            }
        }

        private void AddTeamBands(VisioPage page) {
            foreach (TeamBandItem band in _bands) {
                List<OrgNode> members = GetBandMembers(band.Id);
                if (members.Count == 0) {
                    continue;
                }

                GetBounds(members, out double left, out double bottom, out double right, out double top);
                double width = (right - left) + (_bandPadding * 2D);
                double height = (top - bottom) + (_bandPadding * 2D);
                double x = (left + right) / 2D;
                double y = (bottom + top) / 2D;
                VisioShape shape = new(GetBandShapeId(band.Id), x, y, width, height, band.Text) {
                    NameU = "Rectangle",
                    Master = _document.EnsureBuiltinMaster("Rectangle")
                };
                _theme.Container.ApplyTo(shape);
                shape.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.BackgroundSurfaceKind, "STR", prompt: "OfficeIMO semantic kind");
                page.Shapes.Add(shape);
            }
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

        private void AddNodes(VisioPage page) {
            foreach (OrgNode node in _nodes) {
                GetNodeShape(node.Kind, out string masterNameU, out double width, out double height);
                VisioShape shape = new(node.Id, node.X, node.Y, width, height, GetNodeText(node)) {
                    NameU = masterNameU,
                    Master = _document.EnsureBuiltinMaster(masterNameU)
                };
                GetNodeStyle(node.Kind).ApplyTo(shape);
                node.Shape = shape;
                page.Shapes.Add(shape);
            }
        }

        private void AddReportingLines(VisioPage page) {
            int routeIndex = 0;
            foreach (OrgNode node in _nodes) {
                if (string.IsNullOrWhiteSpace(node.ManagerId)) {
                    continue;
                }

                OrgNode manager = _nodesById[node.ManagerId!];
                if (manager.Shape == null || node.Shape == null) {
                    throw new InvalidOperationException("Org chart nodes must be placed before reporting lines are created.");
                }

                VisioConnector connector;
                if (node.Kind == VisioOrgChartNodeKind.Assistant) {
                    connector = page.AddConnector(manager.Shape, node.Shape, ConnectorKind.RightAngle, VisioSide.Right, VisioSide.Left);
                    _theme.ControlConnector.ApplyTo(connector);
                    connector.RouteOrthogonal(VisioConnectorRouteStyle.HorizontalThenVertical, (routeIndex % 2) * 0.04D);
                } else {
                    connector = page.AddConnector(manager.Shape, node.Shape, ConnectorKind.RightAngle, VisioSide.Bottom, VisioSide.Top);
                    _theme.Connector.ApplyTo(connector);
                    connector.RouteOrthogonal(VisioConnectorRouteStyle.VerticalThenHorizontal, (routeIndex % 3) * 0.04D);
                }

                routeIndex++;
            }
        }

        private void AddCallouts(VisioPage page) {
            foreach (CalloutItem callout in _callouts) {
                OrgNode target = _nodesById[callout.TargetId];
                if (target.Shape == null) {
                    throw new InvalidOperationException("Org chart nodes must be placed before callouts are created.");
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

        private List<OrgNode> GetDirectReports(string managerId) {
            List<OrgNode> children = new();
            foreach (OrgNode node in _nodes) {
                if (node.Kind == VisioOrgChartNodeKind.Assistant) {
                    continue;
                }

                if (string.Equals(node.ManagerId, managerId, StringComparison.Ordinal)) {
                    children.Add(node);
                }
            }

            return children;
        }

        private List<OrgNode> GetBandMembers(string bandId) {
            List<OrgNode> members = new();
            foreach (OrgNode node in _nodes) {
                if (string.Equals(node.BandId, bandId, StringComparison.Ordinal)) {
                    members.Add(node);
                }
            }

            return members;
        }

        private double LevelCenterY(int depth) {
            return _pageHeight - _topMargin - HeaderHeight - (_nodeHeight / 2D) - (depth * (_nodeHeight + _levelGap));
        }

        private double HeaderHeight => string.IsNullOrWhiteSpace(_titleText) ? 0D : _titleHeight + _titleGap;

        private void GetNodeShape(VisioOrgChartNodeKind kind, out string masterNameU, out double width, out double height) {
            width = _nodeWidth;
            height = _nodeHeight;
            switch (kind) {
                case VisioOrgChartNodeKind.Assistant:
                    masterNameU = "Rectangle";
                    width = _nodeWidth * 0.92D;
                    height = _nodeHeight * 0.78D;
                    break;
                case VisioOrgChartNodeKind.Vacancy:
                case VisioOrgChartNodeKind.External:
                    masterNameU = "Rectangle";
                    break;
                default:
                    masterNameU = "Process";
                    break;
            }
        }

        private VisioShapeStyle GetNodeStyle(VisioOrgChartNodeKind kind) {
            switch (kind) {
                case VisioOrgChartNodeKind.Executive:
                    return _theme.Emphasis;
                case VisioOrgChartNodeKind.Manager:
                    return _theme.Primary;
                case VisioOrgChartNodeKind.Assistant:
                    return _theme.Marker;
                case VisioOrgChartNodeKind.Vacancy:
                    return _theme.Container;
                case VisioOrgChartNodeKind.External:
                    return _theme.Decision;
                default:
                    return _theme.Success;
            }
        }

        private static string GetNodeText(OrgNode node) {
            if (string.IsNullOrWhiteSpace(node.Title)) {
                return node.Name;
            }

            return node.Name + Environment.NewLine + node.Title;
        }

        private static void GetBounds(IReadOnlyList<OrgNode> nodes, out double left, out double bottom, out double right, out double top) {
            left = double.MaxValue;
            bottom = double.MaxValue;
            right = double.MinValue;
            top = double.MinValue;
            foreach (OrgNode node in nodes) {
                left = Math.Min(left, node.X - (node.Width / 2D));
                bottom = Math.Min(bottom, node.Y - (node.Height / 2D));
                right = Math.Max(right, node.X + (node.Width / 2D));
                top = Math.Max(top, node.Y + (node.Height / 2D));
            }
        }

        private static void EnsureNoCycle(OrgNode node, HashSet<string> visiting) {
            if (!visiting.Add(node.Id)) {
                throw new InvalidOperationException($"The org chart contains a reporting cycle at node '{node.Id}'.");
            }
        }

        private void EnsureKnownNode(string? id, string parameterName) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Org chart node id cannot be null or whitespace.", parameterName);
            }

            string normalizedId = id!.Trim();
            if (!_nodesById.ContainsKey(normalizedId)) {
                throw new ArgumentException($"Unknown org chart node id '{normalizedId}'.", parameterName);
            }
        }

        private void EnsureKnownBand(string id, string parameterName) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Team band id cannot be null or whitespace.", parameterName);
            }

            string normalizedId = id.Trim();
            if (!_bandsById.ContainsKey(normalizedId)) {
                throw new ArgumentException($"Unknown org chart team band id '{normalizedId}'.", parameterName);
            }
        }

        private bool IsShapeIdInUse(string id) {
            if (!string.IsNullOrWhiteSpace(_titleText) && string.Equals(_titleId, id, StringComparison.Ordinal)) {
                return true;
            }

            if (_nodesById.ContainsKey(id)) {
                return true;
            }

            foreach (TeamBandItem band in _bands) {
                if (string.Equals(GetBandShapeId(band.Id), id, StringComparison.Ordinal)) {
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
            if (!IsShapeIdInUse(id)) {
                return id;
            }

            int index = 2;
            while (IsShapeIdInUse(id + "-" + index)) {
                index++;
            }

            return id + "-" + index;
        }

        private static string GetBandShapeId(string bandId) {
            return "org-band-" + bandId;
        }

        private static string RequireId(string id, string parameterName, string label) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException(label + " cannot be null or whitespace.", parameterName);
            }

            return id.Trim();
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
