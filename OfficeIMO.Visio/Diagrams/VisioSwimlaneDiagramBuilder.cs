using System;
using System.Collections.Generic;
using System.Linq;

using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level builder for dependency-free swimlane and cross-functional process diagrams.
    /// </summary>
    public sealed class VisioSwimlaneDiagramBuilder {
        private sealed class LaneItem {
            public LaneItem(string id, string text) {
                Id = id;
                Text = text;
            }

            public string Id { get; }

            public string Text { get; }
        }

        private sealed class PhaseItem {
            public PhaseItem(string id, string text) {
                Id = id;
                Text = text;
            }

            public string Id { get; }

            public string Text { get; }
        }

        private sealed class ActivityItem {
            public ActivityItem(string id, string text, string laneId, string phaseId, VisioSwimlaneActivityKind kind) {
                Id = id;
                Text = text;
                LaneId = laneId;
                PhaseId = phaseId;
                Kind = kind;
            }

            public string Id { get; }

            public string Text { get; }

            public string LaneId { get; }

            public string PhaseId { get; }

            public VisioSwimlaneActivityKind Kind { get; }

            public VisioShape? Shape { get; set; }
        }

        private sealed class FlowItem {
            public FlowItem(string fromId, string toId, VisioSwimlaneConnectorKind kind, string? label) {
                FromId = fromId;
                ToId = toId;
                Kind = kind;
                Label = label;
            }

            public string FromId { get; }

            public string ToId { get; }

            public VisioSwimlaneConnectorKind Kind { get; }

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
        private readonly List<LaneItem> _lanes = new();
        private readonly Dictionary<string, LaneItem> _lanesById = new(StringComparer.Ordinal);
        private readonly List<PhaseItem> _phases = new();
        private readonly Dictionary<string, PhaseItem> _phasesById = new(StringComparer.Ordinal);
        private readonly List<ActivityItem> _activities = new();
        private readonly Dictionary<string, ActivityItem> _activitiesById = new(StringComparer.Ordinal);
        private readonly List<FlowItem> _flows = new();
        private readonly List<CalloutItem> _callouts = new();
        private VisioStyleTheme _theme = VisioStyleTheme.Modern();
        private VisioMeasurementUnit _unit = VisioMeasurementUnit.Inches;
        private double _pageWidth = 14;
        private double _pageHeight = 8.5;
        private double _leftMargin = 0.6;
        private double _rightMargin = 0.6;
        private double _topMargin = 0.6;
        private double _bottomMargin = 0.6;
        private double _laneHeaderWidth = 1.35;
        private double _phaseHeaderHeight = 0.55;
        private double _phaseWidth = 2.4;
        private double _laneHeight = 1.45;
        private double _activityWidth = 1.6;
        private double _activityHeight = 0.72;
        private double _activityStackGap = 0.12;
        private string? _titleText;
        private string _titleId = "title";
        private double _titleHeight = 0.45;
        private double _titleGap = 0.35;
        private bool _built;

        internal VisioSwimlaneDiagramBuilder(VisioDocument document, string pageName) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _pageName = string.IsNullOrWhiteSpace(pageName) ? "Swimlane Diagram" : pageName;
        }

        /// <summary>Sets the page size used by the generated swimlane page.</summary>
        public VisioSwimlaneDiagramBuilder PageSize(double width, double height, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _pageWidth = width;
            _pageHeight = height;
            _unit = unit;
            return this;
        }

        /// <summary>Sets the visual theme.</summary>
        public VisioSwimlaneDiagramBuilder Theme(VisioStyleTheme theme) {
            _theme = (theme ?? throw new ArgumentNullException(nameof(theme))).Clone();
            return this;
        }

        /// <summary>Sets the outer page margins.</summary>
        public VisioSwimlaneDiagramBuilder Margins(double left, double top, double right = 0.6, double bottom = 0.6) {
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

        /// <summary>Sets lane and phase dimensions.</summary>
        public VisioSwimlaneDiagramBuilder GridSize(double phaseWidth, double laneHeight, double laneHeaderWidth = 1.35, double phaseHeaderHeight = 0.55) {
            ValidatePositive(phaseWidth, nameof(phaseWidth));
            ValidatePositive(laneHeight, nameof(laneHeight));
            ValidatePositive(laneHeaderWidth, nameof(laneHeaderWidth));
            ValidatePositive(phaseHeaderHeight, nameof(phaseHeaderHeight));
            _phaseWidth = phaseWidth;
            _laneHeight = laneHeight;
            _laneHeaderWidth = laneHeaderWidth;
            _phaseHeaderHeight = phaseHeaderHeight;
            return this;
        }

        /// <summary>Sets the default activity size.</summary>
        public VisioSwimlaneDiagramBuilder ActivitySize(double width, double height) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _activityWidth = width;
            _activityHeight = height;
            return this;
        }

        /// <summary>Sets the vertical gap used when multiple activities share the same lane/phase cell.</summary>
        public VisioSwimlaneDiagramBuilder ActivityStackGap(double gap) {
            ValidateNonNegative(gap, nameof(gap));
            _activityStackGap = gap;
            return this;
        }

        /// <summary>Adds a centered editable title above the generated swimlane grid.</summary>
        public VisioSwimlaneDiagramBuilder Title(string? text = null, string id = "title", double height = 0.45, double gap = 0.35) {
            string normalizedId = RequireId(id, nameof(id), "Title id");
            if (IsShapeIdInUse(normalizedId)) {
                throw new ArgumentException($"A swimlane shape with id '{normalizedId}' already exists.", nameof(id));
            }

            ValidatePositive(height, nameof(height));
            ValidateNonNegative(gap, nameof(gap));
            _titleText = string.IsNullOrWhiteSpace(text) ? _pageName : text;
            _titleId = normalizedId;
            _titleHeight = height;
            _titleGap = gap;
            return this;
        }

        /// <summary>Adds a horizontal lane for a role, team, or system.</summary>
        public VisioSwimlaneDiagramBuilder Lane(string id, string text) {
            string normalizedId = RequireId(id, nameof(id), "Lane id");
            if (_lanesById.ContainsKey(normalizedId)) {
                throw new ArgumentException($"A swimlane lane with id '{normalizedId}' already exists.", nameof(id));
            }

            if (IsShapeIdInUse("lane-" + normalizedId) || IsShapeIdInUse("lane-header-" + normalizedId)) {
                throw new ArgumentException($"A swimlane shape with id '{normalizedId}' already exists.", nameof(id));
            }

            _lanesById.Add(normalizedId, new LaneItem(normalizedId, text ?? string.Empty));
            _lanes.Add(_lanesById[normalizedId]);
            return this;
        }

        /// <summary>Adds a vertical phase/milestone column.</summary>
        public VisioSwimlaneDiagramBuilder Phase(string id, string text) {
            string normalizedId = RequireId(id, nameof(id), "Phase id");
            if (_phasesById.ContainsKey(normalizedId)) {
                throw new ArgumentException($"A swimlane phase with id '{normalizedId}' already exists.", nameof(id));
            }

            if (IsShapeIdInUse("phase-" + normalizedId)) {
                throw new ArgumentException($"A swimlane shape with id '{normalizedId}' already exists.", nameof(id));
            }

            _phasesById.Add(normalizedId, new PhaseItem(normalizedId, text ?? string.Empty));
            _phases.Add(_phasesById[normalizedId]);
            return this;
        }

        /// <summary>Adds a start activity.</summary>
        public VisioSwimlaneDiagramBuilder Start(string id, string text, string laneId, string phaseId) =>
            Activity(id, text, laneId, phaseId, VisioSwimlaneActivityKind.Start);

        /// <summary>Adds a process step.</summary>
        public VisioSwimlaneDiagramBuilder Step(string id, string text, string laneId, string phaseId) =>
            Activity(id, text, laneId, phaseId, VisioSwimlaneActivityKind.Step);

        /// <summary>Adds a decision activity.</summary>
        public VisioSwimlaneDiagramBuilder Decision(string id, string text, string laneId, string phaseId) =>
            Activity(id, text, laneId, phaseId, VisioSwimlaneActivityKind.Decision);

        /// <summary>Adds a data/input/output activity.</summary>
        public VisioSwimlaneDiagramBuilder Data(string id, string text, string laneId, string phaseId) =>
            Activity(id, text, laneId, phaseId, VisioSwimlaneActivityKind.Data);

        /// <summary>Adds an end activity.</summary>
        public VisioSwimlaneDiagramBuilder End(string id, string text, string laneId, string phaseId) =>
            Activity(id, text, laneId, phaseId, VisioSwimlaneActivityKind.End);

        /// <summary>Adds an activity in a lane and phase.</summary>
        public VisioSwimlaneDiagramBuilder Activity(string id, string text, string laneId, string phaseId, VisioSwimlaneActivityKind kind = VisioSwimlaneActivityKind.Step) {
            string normalizedId = RequireId(id, nameof(id), "Activity id");
            EnsureKnownLane(laneId, nameof(laneId));
            EnsureKnownPhase(phaseId, nameof(phaseId));
            if (_activitiesById.ContainsKey(normalizedId)) {
                throw new ArgumentException($"A swimlane activity with id '{normalizedId}' already exists.", nameof(id));
            }

            if (IsShapeIdInUse(normalizedId)) {
                throw new ArgumentException($"A swimlane shape with id '{normalizedId}' already exists.", nameof(id));
            }

            if (!Enum.IsDefined(typeof(VisioSwimlaneActivityKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            ActivityItem item = new(normalizedId, text ?? string.Empty, laneId, phaseId, kind);
            _activities.Add(item);
            _activitiesById.Add(normalizedId, item);
            return this;
        }

        /// <summary>Adds a standard process flow.</summary>
        public VisioSwimlaneDiagramBuilder Flow(string fromId, string toId, string? label = null) =>
            AddFlow(fromId, toId, VisioSwimlaneConnectorKind.Flow, label);

        /// <summary>Adds a cross-lane handoff flow.</summary>
        public VisioSwimlaneDiagramBuilder Handoff(string fromId, string toId, string? label = null) =>
            AddFlow(fromId, toId, VisioSwimlaneConnectorKind.Handoff, label);

        /// <summary>Adds an exception or alternate-path flow.</summary>
        public VisioSwimlaneDiagramBuilder Exception(string fromId, string toId, string? label = null) =>
            AddFlow(fromId, toId, VisioSwimlaneConnectorKind.Exception, label);

        /// <summary>Adds a semantic callout connected to a known swimlane activity using a generated callout id.</summary>
        public VisioSwimlaneDiagramBuilder Callout(string targetId, string text, double pinX, double pinY, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            EnsureKnownActivity(normalizedTargetId, nameof(targetId));
            return Callout(normalizedTargetId, CreateCalloutId(normalizedTargetId), text, pinX, pinY, configure);
        }

        /// <summary>Adds a semantic callout connected to a known swimlane activity.</summary>
        public VisioSwimlaneDiagramBuilder Callout(string targetId, string id, string text, double pinX, double pinY, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            string normalizedId = RequireId(id, nameof(id), "Callout id");
            EnsureKnownActivity(normalizedTargetId, nameof(targetId));
            if (IsShapeIdInUse(normalizedId)) {
                throw new ArgumentException($"A swimlane shape with id '{normalizedId}' already exists.", nameof(id));
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

        /// <summary>Adds a semantic callout placed beside a known swimlane activity using a generated callout id.</summary>
        public VisioSwimlaneDiagramBuilder Callout(string targetId, string text, VisioSide placement, double gap = 0.35D, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            EnsureKnownActivity(normalizedTargetId, nameof(targetId));
            return Callout(normalizedTargetId, CreateCalloutId(normalizedTargetId), text, placement, gap, configure);
        }

        /// <summary>Adds a semantic callout placed beside a known swimlane activity.</summary>
        public VisioSwimlaneDiagramBuilder Callout(string targetId, string id, string text, VisioSide placement, double gap = 0.35D, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            string normalizedId = RequireId(id, nameof(id), "Callout id");
            EnsureKnownActivity(normalizedTargetId, nameof(targetId));
            if (IsShapeIdInUse(normalizedId)) {
                throw new ArgumentException($"A swimlane shape with id '{normalizedId}' already exists.", nameof(id));
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
                throw new InvalidOperationException("This swimlane diagram builder has already produced a page.");
            }

            _built = true;
            if (_lanes.Count == 0) {
                throw new InvalidOperationException("A swimlane diagram requires at least one lane.");
            }

            if (_phases.Count == 0) {
                throw new InvalidOperationException("A swimlane diagram requires at least one phase.");
            }

            if (_activities.Count == 0) {
                throw new InvalidOperationException("A swimlane diagram requires at least one activity.");
            }

            EnsurePageCanFitLayout();

            VisioPage page = _document.AddPage(_pageName, _pageWidth, _pageHeight, _unit);
            page.Grid(visible: false, snap: true);
            AddLanesAndPhases(page);
            AddActivities(page);
            AddFlows(page);
            AddCallouts(page);
            AddTitle(page);
            EnsureSideCalloutsFitPage(page);
            _document.RequestRecalcOnOpen();
            return page;
        }

        private VisioSwimlaneDiagramBuilder AddFlow(string fromId, string toId, VisioSwimlaneConnectorKind kind, string? label) {
            EnsureKnownActivity(fromId, nameof(fromId));
            EnsureKnownActivity(toId, nameof(toId));
            if (!Enum.IsDefined(typeof(VisioSwimlaneConnectorKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            _flows.Add(new FlowItem(fromId, toId, kind, label));
            return this;
        }

        private void EnsurePageCanFitLayout() {
            _laneHeight = Math.Max(_laneHeight, GetRequiredLaneHeight());
            double requiredWidth = _leftMargin + _laneHeaderWidth + (_phases.Count * _phaseWidth) + _rightMargin;
            double requiredHeight = _topMargin + HeaderHeight + _phaseHeaderHeight + (_lanes.Count * _laneHeight) + _bottomMargin;
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

        private double GetRequiredLaneHeight() {
            Dictionary<string, double> heightsByCell = new(StringComparer.Ordinal);
            foreach (ActivityItem activity in _activities) {
                GetActivityShape(activity.Kind, out _, out _, out double height);
                string key = GetCellKey(activity.LaneId, activity.PhaseId);
                if (heightsByCell.TryGetValue(key, out double currentHeight)) {
                    heightsByCell[key] = currentHeight + _activityStackGap + height;
                } else {
                    heightsByCell[key] = height;
                }
            }

            double required = _laneHeight;
            foreach (double cellHeight in heightsByCell.Values) {
                required = Math.Max(required, cellHeight + 0.35);
            }

            return required;
        }

        private void AddLanesAndPhases(VisioPage page) {
            double processWidth = _phases.Count * _phaseWidth;
            double processCenterX = _leftMargin + _laneHeaderWidth + (processWidth / 2D);

            for (int i = 0; i < _lanes.Count; i++) {
                LaneItem lane = _lanes[i];
                double y = LaneCenterY(i);
                VisioShape laneHeader = new("lane-header-" + lane.Id, _leftMargin + (_laneHeaderWidth / 2D), y, _laneHeaderWidth, _laneHeight, lane.Text) {
                    NameU = "Rectangle",
                    Master = _document.EnsureBuiltinMaster("Rectangle")
                };
                _theme.Emphasis.ApplyTo(laneHeader);
                page.Shapes.Add(laneHeader);

                VisioShape laneBody = new("lane-" + lane.Id, processCenterX, y, processWidth, _laneHeight, string.Empty) {
                    NameU = "Rectangle",
                    Master = _document.EnsureBuiltinMaster("Rectangle")
                };
                ApplyLaneBodyStyle(laneBody, i);
                laneBody.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.BackgroundSurfaceKind, "STR", prompt: "OfficeIMO semantic kind");
                page.Shapes.Add(laneBody);
            }

            for (int i = 0; i < _phases.Count; i++) {
                PhaseItem phase = _phases[i];
                VisioShape phaseHeader = new("phase-" + phase.Id, PhaseCenterX(i), PhaseHeaderCenterY(), _phaseWidth, _phaseHeaderHeight, phase.Text) {
                    NameU = "Rectangle",
                    Master = _document.EnsureBuiltinMaster("Rectangle")
                };
                _theme.Container.ApplyTo(phaseHeader);
                phaseHeader.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.BackgroundSurfaceKind, "STR", prompt: "OfficeIMO semantic kind");
                page.Shapes.Add(phaseHeader);
            }
        }

        private void AddActivities(VisioPage page) {
            foreach (ActivityItem activity in _activities) {
                int laneIndex = IndexOfLane(activity.LaneId);
                int phaseIndex = IndexOfPhase(activity.PhaseId);
                GetActivityShape(activity.Kind, out string masterNameU, out double width, out double height);
                VisioShape shape = new(activity.Id, PhaseCenterX(phaseIndex), ActivityCenterY(activity, laneIndex, height), width, height, activity.Text) {
                    NameU = masterNameU,
                    Master = _document.EnsureBuiltinMaster(masterNameU)
                };
                GetActivityStyle(activity.Kind).ApplyTo(shape);
                activity.Shape = shape;
                page.Shapes.Add(shape);
            }
        }

        private double ActivityCenterY(ActivityItem activity, int laneIndex, double height) {
            List<ActivityItem> stack = GetActivitiesInCell(activity.LaneId, activity.PhaseId);
            if (stack.Count == 1) {
                return LaneCenterY(laneIndex);
            }

            double totalHeight = 0D;
            for (int i = 0; i < stack.Count; i++) {
                GetActivityShape(stack[i].Kind, out _, out _, out double stackHeight);
                totalHeight += stackHeight;
                if (i > 0) {
                    totalHeight += _activityStackGap;
                }
            }

            double y = LaneCenterY(laneIndex) + (totalHeight / 2D);
            for (int i = 0; i < stack.Count; i++) {
                GetActivityShape(stack[i].Kind, out _, out _, out double stackHeight);
                y -= stackHeight / 2D;
                if (ReferenceEquals(stack[i], activity)) {
                    return y;
                }

                y -= stackHeight / 2D + _activityStackGap;
            }

            return LaneCenterY(laneIndex);
        }

        private List<ActivityItem> GetActivitiesInCell(string laneId, string phaseId) {
            List<ActivityItem> activities = new();
            foreach (ActivityItem activity in _activities) {
                if (string.Equals(activity.LaneId, laneId, StringComparison.Ordinal) &&
                    string.Equals(activity.PhaseId, phaseId, StringComparison.Ordinal)) {
                    activities.Add(activity);
                }
            }

            return activities;
        }

        private void AddFlows(VisioPage page) {
            int routeIndex = 0;
            foreach (FlowItem flow in _flows) {
                ActivityItem from = _activitiesById[flow.FromId];
                ActivityItem to = _activitiesById[flow.ToId];
                if (from.Shape == null || to.Shape == null) {
                    throw new InvalidOperationException("Activities must be placed before flows are created.");
                }

                ResolveSides(from.Shape, to.Shape, out VisioSide fromSide, out VisioSide toSide);
                VisioConnector connector = page.AddConnector(from.Shape, to.Shape, ConnectorKind.RightAngle, fromSide, toSide);
                GetConnectorStyle(flow.Kind).ApplyTo(connector);
                connector.Label = flow.Label;
                RouteFlow(connector, from, to, routeIndex);
                if (!string.IsNullOrWhiteSpace(flow.Label)) {
                    connector.PlaceLabel(0.5, offsetY: 0.16, width: 1.1);
                }

                routeIndex++;
            }
        }

        private void AddCallouts(VisioPage page) {
            foreach (CalloutItem callout in _callouts) {
                ActivityItem target = _activitiesById[callout.TargetId];
                if (target.Shape == null) {
                    throw new InvalidOperationException("Activities must be placed before callouts are created.");
                }

                if (callout.UsePlacement) {
                    page.AddCallout(target.Shape, callout.Id, callout.Text, callout.Placement, callout.Gap, callout.Options);
                } else {
                    page.AddCallout(target.Shape, callout.Id, callout.Text, callout.PinX, callout.PinY, callout.Options);
                }
            }
        }

        private void EnsureSideCalloutsFitPage(VisioPage page) {
            if (!_callouts.Any(callout => callout.UsePlacement)) {
                return;
            }

            VisioShapeBounds bounds = page.GetContentBounds();
            if (bounds.IsEmpty) {
                return;
            }

            double horizontalMargin = Math.Min(_leftMargin, _rightMargin);
            double verticalMargin = Math.Min(_topMargin, _bottomMargin);
            bool overflows = bounds.Left < horizontalMargin ||
                             bounds.Bottom < verticalMargin ||
                             bounds.Right > page.Width - horizontalMargin ||
                             bounds.Top > page.Height - verticalMargin;
            if (overflows) {
                page.FitToContent(horizontalMargin, verticalMargin);
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

        private void RouteFlow(VisioConnector connector, ActivityItem from, ActivityItem to, int routeIndex) {
            int fromLane = IndexOfLane(from.LaneId);
            int toLane = IndexOfLane(to.LaneId);
            int fromPhase = IndexOfPhase(from.PhaseId);
            int toPhase = IndexOfPhase(to.PhaseId);
            double offset = (routeIndex % 3) * 0.04;

            if (fromLane != toLane && fromPhase != toPhase) {
                connector.RouteOrthogonal(VisioConnectorRouteStyle.HorizontalThenVertical, offset);
                return;
            }

            if (fromLane != toLane) {
                connector.RouteOrthogonal(VisioConnectorRouteStyle.VerticalThenHorizontal, offset);
                return;
            }

            connector.RouteOrthogonal(VisioConnectorRouteStyle.HorizontalThenVertical, offset);
        }

        private void ApplyLaneBodyStyle(VisioShape shape, int laneIndex) {
            _theme.Container.ApplyTo(shape);
            if (laneIndex % 2 == 1) {
                shape.FillColor = Blend(_theme.Container.FillColor, Color.White, 0.34D);
            }
        }

        private void GetActivityShape(VisioSwimlaneActivityKind kind, out string masterNameU, out double width, out double height) {
            width = _activityWidth;
            height = _activityHeight;
            switch (kind) {
                case VisioSwimlaneActivityKind.Start:
                case VisioSwimlaneActivityKind.End:
                    masterNameU = "Ellipse";
                    width = _activityWidth * 0.95;
                    break;
                case VisioSwimlaneActivityKind.Decision:
                    masterNameU = "Decision";
                    width = _activityWidth * 0.9;
                    height = _activityHeight * 1.35;
                    break;
                case VisioSwimlaneActivityKind.Data:
                    masterNameU = "Data";
                    break;
                default:
                    masterNameU = "Process";
                    break;
            }
        }

        private VisioShapeStyle GetActivityStyle(VisioSwimlaneActivityKind kind) {
            switch (kind) {
                case VisioSwimlaneActivityKind.Start:
                case VisioSwimlaneActivityKind.End:
                    return _theme.Success;
                case VisioSwimlaneActivityKind.Decision:
                    return _theme.Decision;
                case VisioSwimlaneActivityKind.Data:
                    return _theme.Emphasis;
                default:
                    return _theme.Primary;
            }
        }

        private VisioConnectorStyle GetConnectorStyle(VisioSwimlaneConnectorKind kind) {
            switch (kind) {
                case VisioSwimlaneConnectorKind.Handoff:
                    return _theme.DataConnector;
                case VisioSwimlaneConnectorKind.Exception:
                    return _theme.ControlConnector;
                default:
                    return _theme.Connector;
            }
        }

        private double PhaseCenterX(int phaseIndex) {
            return _leftMargin + _laneHeaderWidth + (phaseIndex * _phaseWidth) + (_phaseWidth / 2D);
        }

        private double LaneCenterY(int laneIndex) {
            return _pageHeight - _topMargin - HeaderHeight - _phaseHeaderHeight - (laneIndex * _laneHeight) - (_laneHeight / 2D);
        }

        private double PhaseHeaderCenterY() {
            return _pageHeight - _topMargin - HeaderHeight - (_phaseHeaderHeight / 2D);
        }

        private double HeaderHeight => string.IsNullOrWhiteSpace(_titleText) ? 0D : _titleHeight + _titleGap;

        private int IndexOfLane(string laneId) {
            for (int i = 0; i < _lanes.Count; i++) {
                if (string.Equals(_lanes[i].Id, laneId, StringComparison.Ordinal)) {
                    return i;
                }
            }

            return -1;
        }

        private int IndexOfPhase(string phaseId) {
            for (int i = 0; i < _phases.Count; i++) {
                if (string.Equals(_phases[i].Id, phaseId, StringComparison.Ordinal)) {
                    return i;
                }
            }

            return -1;
        }

        private void EnsureKnownLane(string id, string parameterName) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Lane id cannot be null or whitespace.", parameterName);
            }

            if (!_lanesById.ContainsKey(id)) {
                throw new ArgumentException($"Unknown swimlane lane id '{id}'.", parameterName);
            }
        }

        private void EnsureKnownPhase(string id, string parameterName) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Phase id cannot be null or whitespace.", parameterName);
            }

            if (!_phasesById.ContainsKey(id)) {
                throw new ArgumentException($"Unknown swimlane phase id '{id}'.", parameterName);
            }
        }

        private void EnsureKnownActivity(string id, string parameterName) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Activity id cannot be null or whitespace.", parameterName);
            }

            if (!_activitiesById.ContainsKey(id)) {
                throw new ArgumentException($"Unknown swimlane activity id '{id}'.", parameterName);
            }
        }

        private static string RequireId(string id, string parameterName, string label) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException(label + " cannot be null or whitespace.", parameterName);
            }

            return id;
        }

        private static string GetCellKey(string laneId, string phaseId) {
            return laneId + "\u001f" + phaseId;
        }

        private bool IsShapeIdInUse(string id) {
            if (!string.IsNullOrWhiteSpace(_titleText) && string.Equals(_titleId, id, StringComparison.Ordinal)) {
                return true;
            }

            foreach (LaneItem lane in _lanes) {
                if (string.Equals("lane-" + lane.Id, id, StringComparison.Ordinal) ||
                    string.Equals("lane-header-" + lane.Id, id, StringComparison.Ordinal)) {
                    return true;
                }
            }

            foreach (PhaseItem phase in _phases) {
                if (string.Equals("phase-" + phase.Id, id, StringComparison.Ordinal)) {
                    return true;
                }
            }

            if (_activitiesById.ContainsKey(id)) {
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
            if (!IsShapeIdInUse(id)) {
                return id;
            }

            int index = 2;
            while (IsShapeIdInUse(id + "-" + index)) {
                index++;
            }

            return id + "-" + index;
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

        private static Color Blend(Color first, Color second, double secondWeight) {
            double clamped = Math.Max(0D, Math.Min(1D, secondWeight));
            double firstWeight = 1D - clamped;
            return Color.FromRgb(
                (byte)Math.Round((first.R * firstWeight) + (second.R * clamped)),
                (byte)Math.Round((first.G * firstWeight) + (second.G * clamped)),
                (byte)Math.Round((first.B * firstWeight) + (second.B * clamped)));
        }
    }
}
