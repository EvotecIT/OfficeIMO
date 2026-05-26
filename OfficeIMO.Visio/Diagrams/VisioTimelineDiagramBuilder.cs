using System;
using System.Collections.Generic;
using System.Globalization;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level builder for dependency-free timelines with date-scaled
    /// milestones, spans, lanes, and clean label stacking.
    /// </summary>
    public sealed class VisioTimelineDiagramBuilder {
        private sealed class MilestoneItem {
            public MilestoneItem(string id, string text, DateTime date, VisioTimelineMilestoneKind kind, VisioTimelinePlacement placement) {
                Id = id;
                Text = text;
                Date = date;
                Kind = kind;
                Placement = placement;
            }

            public string Id { get; }

            public string Text { get; }

            public DateTime Date { get; }

            public VisioTimelineMilestoneKind Kind { get; }

            public VisioTimelinePlacement Placement { get; }

            public VisioShape? MarkerShape { get; set; }
        }

        private sealed class SpanItem {
            public SpanItem(string id, string text, DateTime start, DateTime end, int lane, VisioTimelinePlacement placement) {
                Id = id;
                Text = text;
                Start = start;
                End = end;
                Lane = lane;
                Placement = placement;
            }

            public string Id { get; }

            public string Text { get; }

            public DateTime Start { get; }

            public DateTime End { get; }

            public int Lane { get; }

            public VisioTimelinePlacement Placement { get; }
        }

        private readonly VisioDocument _document;
        private readonly string _pageName;
        private readonly List<MilestoneItem> _milestones = new();
        private readonly Dictionary<string, MilestoneItem> _milestonesById = new(StringComparer.Ordinal);
        private readonly List<SpanItem> _spans = new();
        private readonly Dictionary<string, SpanItem> _spansById = new(StringComparer.Ordinal);
        private VisioStyleTheme _theme = VisioStyleTheme.Modern();
        private VisioMeasurementUnit _unit = VisioMeasurementUnit.Inches;
        private DateTime? _startDate;
        private DateTime? _endDate;
        private double _pageWidth = 14;
        private double _pageHeight = 8.5;
        private double _leftMargin = 0.8;
        private double _rightMargin = 0.8;
        private double _topMargin = 0.7;
        private double _bottomMargin = 0.7;
        private double _axisY = 4.1;
        private double _axisHeight = 0.06;
        private double _milestoneSize = 0.24;
        private double _labelWidth = 1.45;
        private double _labelHeight = 0.48;
        private double _labelGap = 0.18;
        private double _spanHeight = 0.28;
        private double _spanLaneGap = 0.16;
        private string? _titleText;
        private string _titleId = "title";
        private double _titleHeight = 0.45;
        private double _titleGap = 0.35;
        private bool _built;

        internal VisioTimelineDiagramBuilder(VisioDocument document, string pageName) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _pageName = string.IsNullOrWhiteSpace(pageName) ? "Timeline" : pageName;
        }

        /// <summary>Sets the page size used by the generated timeline page.</summary>
        public VisioTimelineDiagramBuilder PageSize(double width, double height, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _pageWidth = width;
            _pageHeight = height;
            _unit = unit;
            return this;
        }

        /// <summary>Sets the visual theme.</summary>
        public VisioTimelineDiagramBuilder Theme(VisioStyleTheme theme) {
            _theme = (theme ?? throw new ArgumentNullException(nameof(theme))).Clone();
            return this;
        }

        /// <summary>Sets the timeline date range. If omitted, the range is inferred from milestones and spans.</summary>
        public VisioTimelineDiagramBuilder Range(DateTime start, DateTime end) {
            if (end <= start) {
                throw new ArgumentException("Timeline end date must be after the start date.", nameof(end));
            }

            _startDate = start.Date;
            _endDate = end.Date;
            return this;
        }

        /// <summary>Sets outer page margins.</summary>
        public VisioTimelineDiagramBuilder Margins(double left, double top, double right = 0.8, double bottom = 0.7) {
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

        /// <summary>Sets the vertical axis position from the bottom of the page.</summary>
        public VisioTimelineDiagramBuilder AxisY(double y) {
            ValidatePositive(y, nameof(y));
            _axisY = y;
            return this;
        }

        /// <summary>Sets default milestone marker and label sizes.</summary>
        public VisioTimelineDiagramBuilder MilestoneSize(double markerSize, double labelWidth, double labelHeight = 0.48) {
            ValidatePositive(markerSize, nameof(markerSize));
            ValidatePositive(labelWidth, nameof(labelWidth));
            ValidatePositive(labelHeight, nameof(labelHeight));
            _milestoneSize = markerSize;
            _labelWidth = labelWidth;
            _labelHeight = labelHeight;
            return this;
        }

        /// <summary>Sets span bar sizing.</summary>
        public VisioTimelineDiagramBuilder SpanSize(double height, double laneGap = 0.16) {
            ValidatePositive(height, nameof(height));
            ValidateNonNegative(laneGap, nameof(laneGap));
            _spanHeight = height;
            _spanLaneGap = laneGap;
            return this;
        }

        /// <summary>Adds a centered editable title above the generated timeline.</summary>
        public VisioTimelineDiagramBuilder Title(string? text = null, string id = "title", double height = 0.45, double gap = 0.35) {
            string normalizedId = RequireId(id, nameof(id), "Title id");
            if (IsShapeIdInUse(normalizedId)) {
                throw new ArgumentException($"A timeline item with shape id '{normalizedId}' already exists.", nameof(id));
            }

            ValidatePositive(height, nameof(height));
            ValidateNonNegative(gap, nameof(gap));
            _titleText = string.IsNullOrWhiteSpace(text) ? _pageName : text;
            _titleId = normalizedId;
            _titleHeight = height;
            _titleGap = gap;
            return this;
        }

        /// <summary>Adds a standard milestone.</summary>
        public VisioTimelineDiagramBuilder Milestone(string id, DateTime date, string text, VisioTimelinePlacement placement = VisioTimelinePlacement.Auto) =>
            AddMilestone(id, date, text, VisioTimelineMilestoneKind.Milestone, placement);

        /// <summary>Adds a release/delivery milestone.</summary>
        public VisioTimelineDiagramBuilder Release(string id, DateTime date, string text, VisioTimelinePlacement placement = VisioTimelinePlacement.Auto) =>
            AddMilestone(id, date, text, VisioTimelineMilestoneKind.Release, placement);

        /// <summary>Adds a decision/approval milestone.</summary>
        public VisioTimelineDiagramBuilder Decision(string id, DateTime date, string text, VisioTimelinePlacement placement = VisioTimelinePlacement.Auto) =>
            AddMilestone(id, date, text, VisioTimelineMilestoneKind.Decision, placement);

        /// <summary>Adds a risk/issue milestone.</summary>
        public VisioTimelineDiagramBuilder Risk(string id, DateTime date, string text, VisioTimelinePlacement placement = VisioTimelinePlacement.Auto) =>
            AddMilestone(id, date, text, VisioTimelineMilestoneKind.Risk, placement);

        /// <summary>Adds a milestone with an explicit semantic kind.</summary>
        public VisioTimelineDiagramBuilder AddMilestone(string id, DateTime date, string text, VisioTimelineMilestoneKind kind, VisioTimelinePlacement placement = VisioTimelinePlacement.Auto) {
            string normalizedId = RequireId(id, nameof(id), "Milestone id");
            if (IsTimelineItemIdInUse(normalizedId) || IsShapeIdInUse(normalizedId) || IsShapeIdInUse(GetMilestoneLabelId(normalizedId))) {
                throw new ArgumentException($"A timeline item with shape id '{normalizedId}' already exists.", nameof(id));
            }

            if (!Enum.IsDefined(typeof(VisioTimelineMilestoneKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            if (!Enum.IsDefined(typeof(VisioTimelinePlacement), placement)) {
                throw new ArgumentOutOfRangeException(nameof(placement));
            }

            MilestoneItem item = new(normalizedId, text ?? string.Empty, date.Date, kind, placement);
            _milestones.Add(item);
            _milestonesById.Add(normalizedId, item);
            return this;
        }

        /// <summary>Adds a span bar between two dates.</summary>
        public VisioTimelineDiagramBuilder Span(string id, DateTime start, DateTime end, string text, int lane = 0, VisioTimelinePlacement placement = VisioTimelinePlacement.Above) {
            string normalizedId = RequireId(id, nameof(id), "Span id");
            if (IsTimelineItemIdInUse(normalizedId) || IsShapeIdInUse(normalizedId)) {
                throw new ArgumentException($"A timeline item with shape id '{normalizedId}' already exists.", nameof(id));
            }

            if (end <= start) {
                throw new ArgumentException("Timeline span end date must be after the start date.", nameof(end));
            }

            if (lane < 0) {
                throw new ArgumentOutOfRangeException(nameof(lane), "Lane must be zero or greater.");
            }

            if (!Enum.IsDefined(typeof(VisioTimelinePlacement), placement) || placement == VisioTimelinePlacement.Auto) {
                throw new ArgumentOutOfRangeException(nameof(placement), "Timeline span placement must be Above or Below.");
            }

            SpanItem item = new(normalizedId, text ?? string.Empty, start.Date, end.Date, lane, placement);
            _spans.Add(item);
            _spansById.Add(normalizedId, item);
            return this;
        }

        internal VisioPage Build() {
            if (_built) {
                throw new InvalidOperationException("This timeline builder has already produced a page.");
            }

            _built = true;
            if (_milestones.Count == 0 && _spans.Count == 0) {
                throw new InvalidOperationException("A timeline requires at least one milestone or span.");
            }

            ResolveRange(out DateTime start, out DateTime end);
            EnsurePageCanFit();

            VisioPage page = _document.AddPage(_pageName, _pageWidth, _pageHeight, _unit);
            page.Grid(visible: false, snap: true);
            AddAxis(page, start, end);
            AddSpans(page, start, end);
            AddMilestones(page, start, end);
            AddTitle(page);
            _document.RequestRecalcOnOpen();
            return page;
        }

        private void ResolveRange(out DateTime start, out DateTime end) {
            if (_startDate.HasValue && _endDate.HasValue) {
                start = _startDate.Value;
                end = _endDate.Value;
            } else {
                start = DateTime.MaxValue;
                end = DateTime.MinValue;
                foreach (MilestoneItem milestone in _milestones) {
                    start = Min(start, milestone.Date);
                    end = Max(end, milestone.Date);
                }

                foreach (SpanItem span in _spans) {
                    start = Min(start, span.Start);
                    end = Max(end, span.End);
                }

                if (start == DateTime.MaxValue || end == DateTime.MinValue) {
                    throw new InvalidOperationException("A timeline requires at least one dated item.");
                }

                if (start == end) {
                    end = start.AddDays(1);
                }
            }

            foreach (MilestoneItem milestone in _milestones) {
                EnsureDateInRange(milestone.Date, start, end, milestone.Id);
            }

            foreach (SpanItem span in _spans) {
                EnsureDateInRange(span.Start, start, end, span.Id);
                EnsureDateInRange(span.End, start, end, span.Id);
            }
        }

        private void EnsurePageCanFit() {
            double minimumAxisY = _bottomMargin + 1.2D;
            double maximumAxisY = Math.Max(minimumAxisY, _pageHeight - _topMargin - HeaderHeight - 1.2D);
            _axisY = Math.Min(Math.Max(_axisY, minimumAxisY), maximumAxisY);
            int aboveSpanLanes = GetMaxSpanLane(VisioTimelinePlacement.Above) + 1;
            int belowSpanLanes = GetMaxSpanLane(VisioTimelinePlacement.Below) + 1;
            double aboveNeeded = _topMargin + HeaderHeight + RequiredDistanceFromAxis(Math.Max(0, aboveSpanLanes)) + 0.25D;
            double belowNeeded = _bottomMargin + RequiredDistanceFromAxis(Math.Max(0, belowSpanLanes)) + 0.25D;
            if (_pageHeight - _axisY < aboveNeeded) {
                _pageHeight = _axisY + aboveNeeded;
            }

            if (_axisY < belowNeeded) {
                double delta = belowNeeded - _axisY;
                _pageHeight += delta;
                _axisY += delta;
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

        private void AddAxis(VisioPage page, DateTime start, DateTime end) {
            double width = TimelineWidth();
            VisioShape axis = new("timeline-axis", _leftMargin + (width / 2D), _axisY, width, _axisHeight, string.Empty) {
                NameU = "Rectangle",
                Master = _document.EnsureBuiltinMaster("Rectangle")
            };
            _theme.Emphasis.ApplyTo(axis);
            page.Shapes.Add(axis);

            AddTick(page, "timeline-start", start, start);
            AddTick(page, "timeline-end", end, start);
        }

        private void AddTick(VisioPage page, string id, DateTime date, DateTime rangeStart) {
            double x = DateX(date, rangeStart, date == rangeStart ? date.AddDays(1) : date);
            if (id == "timeline-end") {
                x = _pageWidth - _rightMargin;
            }

            VisioShape label = new(id + "-label", x, _axisY - 0.42D, 1.05, 0.28, FormatShortDate(date)) {
                NameU = "Rectangle",
                Master = _document.EnsureBuiltinMaster("Rectangle")
            };
            _theme.Container.ApplyTo(label);
            page.Shapes.Add(label);
        }

        private void AddSpans(VisioPage page, DateTime start, DateTime end) {
            foreach (SpanItem span in _spans) {
                double startX = DateX(span.Start, start, end);
                double endX = DateX(span.End, start, end);
                double width = Math.Max(0.28D, endX - startX);
                double y = SpanY(span);
                VisioShape shape = new(span.Id, startX + (width / 2D), y, width, _spanHeight, span.Text) {
                    NameU = "Rectangle",
                    Master = _document.EnsureBuiltinMaster("Rectangle")
                };
                _theme.Primary.ApplyTo(shape);
                page.Shapes.Add(shape);
            }
        }

        private void AddMilestones(VisioPage page, DateTime start, DateTime end) {
            Dictionary<VisioTimelinePlacement, List<double>> levelRights = new();
            levelRights[VisioTimelinePlacement.Above] = new List<double>();
            levelRights[VisioTimelinePlacement.Below] = new List<double>();

            List<MilestoneItem> ordered = new(_milestones);
            ordered.Sort((first, second) => first.Date.CompareTo(second.Date));
            for (int i = 0; i < ordered.Count; i++) {
                MilestoneItem milestone = ordered[i];
                double x = DateX(milestone.Date, start, end);
                VisioTimelinePlacement placement = ResolvePlacement(milestone, i);
                int level = ResolveLabelLevel(levelRights[placement], x);
                double markerY = placement == VisioTimelinePlacement.Above
                    ? _axisY + 0.28D
                    : _axisY - 0.28D;
                double labelY = placement == VisioTimelinePlacement.Above
                    ? markerY + (_milestoneSize / 2D) + _labelGap + (_labelHeight / 2D) + (level * (_labelHeight + _labelGap))
                    : markerY - (_milestoneSize / 2D) - _labelGap - (_labelHeight / 2D) - (level * (_labelHeight + _labelGap));

                VisioShape marker = new(milestone.Id, x, markerY, _milestoneSize, _milestoneSize, string.Empty) {
                    NameU = GetMarkerMaster(milestone.Kind),
                    Master = _document.EnsureBuiltinMaster(GetMarkerMaster(milestone.Kind))
                };
                GetMilestoneStyle(milestone.Kind).ApplyTo(marker);
                page.Shapes.Add(marker);
                milestone.MarkerShape = marker;

                VisioShape label = new(GetMilestoneLabelId(milestone.Id), x, labelY, _labelWidth, _labelHeight, GetMilestoneText(milestone)) {
                    NameU = "Rectangle",
                    Master = _document.EnsureBuiltinMaster("Rectangle")
                };
                GetMilestoneLabelStyle(milestone.Kind).ApplyTo(label);
                page.Shapes.Add(label);
            }
        }

        private int ResolveLabelLevel(List<double> levelRights, double x) {
            double left = x - (_labelWidth / 2D);
            for (int i = 0; i < levelRights.Count; i++) {
                if (levelRights[i] + 0.08D <= left) {
                    levelRights[i] = x + (_labelWidth / 2D);
                    return i;
                }
            }

            levelRights.Add(x + (_labelWidth / 2D));
            return levelRights.Count - 1;
        }

        private VisioTimelinePlacement ResolvePlacement(MilestoneItem milestone, int index) {
            if (milestone.Placement == VisioTimelinePlacement.Above || milestone.Placement == VisioTimelinePlacement.Below) {
                return milestone.Placement;
            }

            return index % 2 == 0 ? VisioTimelinePlacement.Above : VisioTimelinePlacement.Below;
        }

        private double SpanY(SpanItem span) {
            double offset = SpanBaseOffset() + (span.Lane * (_spanHeight + _spanLaneGap));
            return span.Placement == VisioTimelinePlacement.Above
                ? _axisY + offset
                : _axisY - offset;
        }

        private double RequiredDistanceFromAxis(int spanLaneCount) {
            double labelDistance = 0.28D + (_milestoneSize / 2D) + _labelGap + _labelHeight;
            double spanDistance = spanLaneCount == 0
                ? 0D
                : SpanBaseOffset() + ((spanLaneCount - 1) * (_spanHeight + _spanLaneGap)) + (_spanHeight / 2D);
            return Math.Max(labelDistance, spanDistance);
        }

        private double SpanBaseOffset() {
            return _milestoneSize + (_labelGap * 2D) + _labelHeight + (_spanHeight / 2D) + 0.1D;
        }

        private double HeaderHeight => string.IsNullOrWhiteSpace(_titleText) ? 0D : _titleHeight + _titleGap;

        private double DateX(DateTime date, DateTime start, DateTime end) {
            double totalDays = Math.Max(1D, (end - start).TotalDays);
            double position = (date - start).TotalDays / totalDays;
            return _leftMargin + (TimelineWidth() * Math.Max(0D, Math.Min(1D, position)));
        }

        private double TimelineWidth() {
            return Math.Max(1D, _pageWidth - _leftMargin - _rightMargin);
        }

        private int GetMaxSpanLane(VisioTimelinePlacement placement) {
            int max = -1;
            foreach (SpanItem span in _spans) {
                if (span.Placement == placement) {
                    max = Math.Max(max, span.Lane);
                }
            }

            return max;
        }

        private static string GetMarkerMaster(VisioTimelineMilestoneKind kind) {
            return kind == VisioTimelineMilestoneKind.Milestone ? "Diamond" : "Circle";
        }

        private VisioShapeStyle GetMilestoneStyle(VisioTimelineMilestoneKind kind) {
            switch (kind) {
                case VisioTimelineMilestoneKind.Release:
                    return _theme.Success;
                case VisioTimelineMilestoneKind.Decision:
                    return _theme.Decision;
                case VisioTimelineMilestoneKind.Risk:
                    return _theme.Emphasis;
                default:
                    return _theme.Marker;
            }
        }

        private VisioShapeStyle GetMilestoneLabelStyle(VisioTimelineMilestoneKind kind) {
            return kind == VisioTimelineMilestoneKind.Risk ? _theme.Emphasis : _theme.Container;
        }

        private static string GetMilestoneText(MilestoneItem milestone) {
            return milestone.Text + Environment.NewLine + FormatShortDate(milestone.Date);
        }

        private bool IsTimelineItemIdInUse(string id) {
            return _milestonesById.ContainsKey(id) || _spansById.ContainsKey(id);
        }

        private bool IsShapeIdInUse(string id) {
            if (!string.IsNullOrWhiteSpace(_titleText) && string.Equals(_titleId, id, StringComparison.Ordinal)) {
                return true;
            }

            if (string.Equals(id, "timeline-axis", StringComparison.Ordinal) ||
                string.Equals(id, "timeline-start-label", StringComparison.Ordinal) ||
                string.Equals(id, "timeline-end-label", StringComparison.Ordinal)) {
                return true;
            }

            foreach (MilestoneItem milestone in _milestones) {
                if (string.Equals(milestone.Id, id, StringComparison.Ordinal) ||
                    string.Equals(GetMilestoneLabelId(milestone.Id), id, StringComparison.Ordinal)) {
                    return true;
                }
            }

            foreach (SpanItem span in _spans) {
                if (string.Equals(span.Id, id, StringComparison.Ordinal)) {
                    return true;
                }
            }

            return false;
        }

        private static string GetMilestoneLabelId(string milestoneId) {
            return milestoneId + "-label";
        }

        private static string FormatShortDate(DateTime date) {
            return date.ToString("MMM d", CultureInfo.InvariantCulture);
        }

        private static void EnsureDateInRange(DateTime date, DateTime start, DateTime end, string id) {
            if (date < start || date > end) {
                throw new ArgumentOutOfRangeException(nameof(date), $"Timeline item '{id}' is outside the configured date range.");
            }
        }

        private static DateTime Min(DateTime first, DateTime second) {
            return first <= second ? first : second;
        }

        private static DateTime Max(DateTime first, DateTime second) {
            return first >= second ? first : second;
        }

        private static string RequireId(string id, string parameterName, string label) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException(label + " cannot be null or whitespace.", parameterName);
            }

            return id.Trim();
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
