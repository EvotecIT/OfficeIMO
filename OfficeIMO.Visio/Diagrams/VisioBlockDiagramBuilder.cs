using System;
using System.Collections.Generic;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level builder for block and system diagrams with grid placement,
    /// grouped regions, and semantic data/control flow connectors.
    /// </summary>
    public sealed class VisioBlockDiagramBuilder {
        private sealed class BlockItem {
            public BlockItem(string id, string text, int column, int row, VisioBlockShapeKind kind, bool emphasis) {
                Id = id;
                Text = text;
                Column = column;
                Row = row;
                Kind = kind;
                Emphasis = emphasis;
            }

            public string Id { get; }

            public string Text { get; }

            public int Column { get; }

            public int Row { get; }

            public VisioBlockShapeKind Kind { get; }

            public bool Emphasis { get; }

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

        private sealed class Link {
            public Link(string fromId, string toId, VisioBlockConnectorKind kind, string? label) {
                FromId = fromId;
                ToId = toId;
                Kind = kind;
                Label = label;
            }

            public string FromId { get; }

            public string ToId { get; }

            public VisioBlockConnectorKind Kind { get; }

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
        private readonly List<BlockItem> _blocks = new List<BlockItem>();
        private readonly Dictionary<string, BlockItem> _blocksById = new Dictionary<string, BlockItem>(StringComparer.Ordinal);
        private readonly List<RegionItem> _regions = new List<RegionItem>();
        private readonly List<Link> _links = new List<Link>();
        private readonly List<CalloutItem> _callouts = new List<CalloutItem>();
        private VisioBlockDiagramTheme _theme = VisioBlockDiagramTheme.TechnicalBlue();
        private VisioMeasurementUnit _unit = VisioMeasurementUnit.Inches;
        private double _pageWidth = 11;
        private double _pageHeight = 8.5;
        private double _leftMargin = 0.8;
        private double _topMargin = 0.8;
        private string? _titleText;
        private string _titleId = "title";
        private double _titleHeight = 0.45;
        private double _titleGap = 0.35;
        private bool _showLegend;
        private string _dataFlowLegendLabel = "Data Flow";
        private string _controlFlowLegendLabel = "Control Flow";
        private bool _built;

        internal VisioBlockDiagramBuilder(VisioDocument document, string pageName) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _pageName = string.IsNullOrWhiteSpace(pageName) ? "Block Diagram" : pageName;
        }

        /// <summary>Sets the page size used by the generated block diagram page.</summary>
        public VisioBlockDiagramBuilder PageSize(double width, double height, VisioMeasurementUnit unit = VisioMeasurementUnit.Inches) {
            ValidatePositive(width, nameof(width));
            ValidatePositive(height, nameof(height));
            _pageWidth = width;
            _pageHeight = height;
            _unit = unit;
            return this;
        }

        /// <summary>Sets the visual theme.</summary>
        public VisioBlockDiagramBuilder Theme(VisioBlockDiagramTheme theme) {
            _theme = (theme ?? throw new ArgumentNullException(nameof(theme))).Clone();
            return this;
        }

        /// <summary>Sets the visual theme from a reusable OfficeIMO Visio style theme.</summary>
        public VisioBlockDiagramBuilder Theme(VisioStyleTheme theme) {
            if (theme == null) {
                throw new ArgumentNullException(nameof(theme));
            }

            return Theme(theme.ToBlockDiagramTheme());
        }

        /// <summary>Sets outer page margins used by the grid layout.</summary>
        public VisioBlockDiagramBuilder Margins(double left, double top) {
            ValidateNonNegative(left, nameof(left));
            ValidateNonNegative(top, nameof(top));
            _leftMargin = left;
            _topMargin = top;
            return this;
        }

        /// <summary>Adds a centered title above the generated grid.</summary>
        public VisioBlockDiagramBuilder Title(string? text = null, string id = "title", double height = 0.45, double gap = 0.35) {
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

        /// <summary>Adds a compact data/control flow legend above the generated grid.</summary>
        public VisioBlockDiagramBuilder Legend(bool enabled = true, string dataFlowLabel = "Data Flow", string controlFlowLabel = "Control Flow") {
            _showLegend = enabled;
            _dataFlowLegendLabel = string.IsNullOrWhiteSpace(dataFlowLabel) ? "Data Flow" : dataFlowLabel;
            _controlFlowLegendLabel = string.IsNullOrWhiteSpace(controlFlowLabel) ? "Control Flow" : controlFlowLabel;
            return this;
        }

        /// <summary>Adds a light background region around a grid area.</summary>
        public VisioBlockDiagramBuilder Region(string id, string text, int column, int row, int columnSpan, int rowSpan) {
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

        /// <summary>Adds a standard block at a deterministic grid position.</summary>
        public VisioBlockDiagramBuilder Block(string id, string text, int column, int row, VisioBlockShapeKind kind = VisioBlockShapeKind.Block) =>
            AddBlock(id, text, column, row, kind, emphasis: false);

        /// <summary>Adds an emphasized block at a deterministic grid position.</summary>
        public VisioBlockDiagramBuilder EmphasisBlock(string id, string text, int column, int row, VisioBlockShapeKind kind = VisioBlockShapeKind.Block) =>
            AddBlock(id, text, column, row, kind, emphasis: true);

        /// <summary>Adds a solid data-flow connector.</summary>
        public VisioBlockDiagramBuilder DataFlow(string fromId, string toId, string? label = null) =>
            AddLink(fromId, toId, VisioBlockConnectorKind.DataFlow, label);

        /// <summary>Adds a dashed control-flow connector.</summary>
        public VisioBlockDiagramBuilder ControlFlow(string fromId, string toId, string? label = null) =>
            AddLink(fromId, toId, VisioBlockConnectorKind.ControlFlow, label);

        /// <summary>Adds a semantic callout connected to a known block using a generated callout id.</summary>
        public VisioBlockDiagramBuilder Callout(string targetId, string text, double pinX, double pinY, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            EnsureKnownBlock(normalizedTargetId, nameof(targetId));
            return Callout(normalizedTargetId, CreateCalloutId(normalizedTargetId), text, pinX, pinY, configure);
        }

        /// <summary>Adds a semantic callout connected to a known block.</summary>
        public VisioBlockDiagramBuilder Callout(string targetId, string id, string text, double pinX, double pinY, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            string normalizedId = RequireId(id, nameof(id), "Callout id");
            EnsureKnownBlock(normalizedTargetId, nameof(targetId));
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

        /// <summary>Adds a semantic callout placed beside a known block using a generated callout id.</summary>
        public VisioBlockDiagramBuilder Callout(string targetId, string text, VisioSide placement, double gap = 0.35D, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            EnsureKnownBlock(normalizedTargetId, nameof(targetId));
            return Callout(normalizedTargetId, CreateCalloutId(normalizedTargetId), text, placement, gap, configure);
        }

        /// <summary>Adds a semantic callout placed beside a known block.</summary>
        public VisioBlockDiagramBuilder Callout(string targetId, string id, string text, VisioSide placement, double gap = 0.35D, Action<VisioCalloutOptions>? configure = null) {
            string normalizedTargetId = RequireId(targetId, nameof(targetId), "Callout target id");
            string normalizedId = RequireId(id, nameof(id), "Callout id");
            EnsureKnownBlock(normalizedTargetId, nameof(targetId));
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
                throw new InvalidOperationException("This block diagram builder has already produced a page.");
            }

            _built = true;
            if (_blocks.Count == 0) {
                throw new InvalidOperationException("A block diagram requires at least one block.");
            }

            VisioPage page = _document.AddPage(_pageName, _pageWidth, _pageHeight, _unit);
            page.Grid(visible: false, snap: true);
            AddRegions(page);
            AddBlocks(page);
            AddLinks(page);
            AddCallouts(page);
            AddAdornments(page);
            page.PolishDiagram(new VisioDiagramPolishOptions {
                FitToContent = false,
                ResizeShapesToText = false,
                ResizeConnectorLabelsToText = false,
                ResolveConnectorShapeIntersections = true,
                ResolveConnectorLabelOverlaps = true
            });
            _document.RequestRecalcOnOpen();
            return page;
        }

        private VisioBlockDiagramBuilder AddBlock(string id, string text, int column, int row, VisioBlockShapeKind kind, bool emphasis) {
            string normalizedId = RequireId(id, nameof(id), "Block id");
            if (IsIdInUse(normalizedId)) {
                throw new ArgumentException($"A diagram item with id '{normalizedId}' already exists.", nameof(id));
            }

            ValidateGridPosition(column, row);
            if (!Enum.IsDefined(typeof(VisioBlockShapeKind), kind)) {
                throw new ArgumentOutOfRangeException(nameof(kind));
            }

            BlockItem block = new BlockItem(normalizedId, text ?? string.Empty, column, row, kind, emphasis);
            _blocks.Add(block);
            _blocksById.Add(normalizedId, block);
            return this;
        }

        private VisioBlockDiagramBuilder AddLink(string fromId, string toId, VisioBlockConnectorKind kind, string? label) {
            string normalizedFromId = RequireId(fromId, nameof(fromId), "From block id");
            string normalizedToId = RequireId(toId, nameof(toId), "To block id");
            EnsureKnownBlock(normalizedFromId, nameof(fromId));
            EnsureKnownBlock(normalizedToId, nameof(toId));
            _links.Add(new Link(normalizedFromId, normalizedToId, kind, label));
            return this;
        }

        private void AddRegions(VisioPage page) {
            for (int i = 0; i < _regions.Count; i++) {
                RegionItem region = _regions[i];
                double x = GridX(region.Column, region.ColumnSpan);
                double y = GridY(region.Row, region.RowSpan);
                double width = region.ColumnSpan * _theme.BlockWidth + (region.ColumnSpan - 1) * _theme.ColumnGap + 0.7;
                double height = region.RowSpan * _theme.BlockHeight + (region.RowSpan - 1) * _theme.RowGap + 0.55;
                VisioShape shape = new VisioShape(region.Id, x.ToInches(_unit), y.ToInches(_unit), width.ToInches(_unit), height.ToInches(_unit), string.Empty) { NameU = "Rectangle" };
                shape.FillColor = _theme.RegionFill;
                shape.LineColor = _theme.RegionStroke;
                shape.LineWeight = _theme.LineWeight;
                if (_theme.RegionTextStyle != null) {
                    shape.TextStyle = _theme.RegionTextStyle.Clone();
                }
                shape.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.BackgroundSurfaceKind, "STR", prompt: "OfficeIMO semantic kind");
                page.Shapes.Add(shape);
                VisioNetworkDiagramVisuals.AddBackgroundZoneCaption(
                    page,
                    CreateGeneratedAdornmentId(VisioNetworkDiagramVisuals.CreateBackgroundZoneCaptionId(region.Id), page),
                    region.Text,
                    x - width / 2D,
                    y + height / 2D,
                    width,
                    CreateCaptionTheme());
            }
        }

        private void AddBlocks(VisioPage page) {
            for (int i = 0; i < _blocks.Count; i++) {
                BlockItem block = _blocks[i];
                double x = GridX(block.Column, 1);
                double y = GridY(block.Row, 1);
                VisioShape shape = CreateBlockShape(page, block, x, y);

                block.Shape = shape;
            }
        }

        private VisioShape CreateBlockShape(VisioPage page, BlockItem block, double x, double y) {
            string nameU;
            double height = _theme.BlockHeight;
            switch (block.Kind) {
                case VisioBlockShapeKind.Data:
                    nameU = "Data";
                    break;
                case VisioBlockShapeKind.Decision:
                    nameU = "Decision";
                    height = _theme.BlockHeight * 1.25;
                    break;
                default:
                    nameU = "Process";
                    break;
            }

            VisioShape shape = page.AddStencilShape(VisioStencils.BlockDiagram, GetBlockStencilId(block.Kind), block.Id, x, y, _theme.BlockWidth, height, block.Text);
            shape.NameU = nameU;
            ApplyBlockStyle(shape, block.Emphasis);
            return shape;
        }

        private void ApplyBlockStyle(VisioShape shape, bool emphasis) {
            shape.FillColor = emphasis ? _theme.EmphasisFill : _theme.BlockFill;
            shape.LineColor = emphasis ? _theme.EmphasisStroke : _theme.BlockStroke;
            shape.LineWeight = _theme.LineWeight;
            VisioTextStyle? textStyle = emphasis ? _theme.EmphasisTextStyle : _theme.BlockTextStyle;
            if (textStyle != null) {
                shape.TextStyle = textStyle.Clone();
            }
        }

        private static string GetBlockStencilId(VisioBlockShapeKind kind) {
            switch (kind) {
                case VisioBlockShapeKind.Data:
                    return "storage";
                case VisioBlockShapeKind.Decision:
                    return "decision";
                default:
                    return "block";
            }
        }

        private void AddLinks(VisioPage page) {
            for (int i = 0; i < _links.Count; i++) {
                Link link = _links[i];
                BlockItem from = _blocksById[link.FromId];
                BlockItem to = _blocksById[link.ToId];
                if (from.Shape == null || to.Shape == null) {
                    throw new InvalidOperationException("Blocks must be placed before connectors are created.");
                }

                ResolveSides(from.Shape, to.Shape, out VisioSide fromSide, out VisioSide toSide);
                VisioConnector connector = page.AddConnector(from.Shape, to.Shape, ConnectorKind.RightAngle, fromSide, toSide);
                connector.EndArrow = EndArrow.Triangle;
                connector.Label = link.Label;
                connector.LineWeight = _theme.LineWeight;
                if (link.Kind == VisioBlockConnectorKind.ControlFlow) {
                    connector.LineColor = _theme.ControlFlowColor;
                    connector.LinePattern = 2;
                } else {
                    connector.LineColor = _theme.DataFlowColor;
                    connector.LinePattern = 1;
                }

                if (_theme.ConnectorTextStyle != null) {
                    connector.TextStyle = _theme.ConnectorTextStyle.Clone();
                }

                if (!string.IsNullOrWhiteSpace(link.Label)) {
                    double labelWidth = Math.Max(1.25D, Math.Min(2.2D, link.Label!.Length * 0.14D));
                    connector.PlaceLabel(0.5D, offsetY: 0.18D, width: labelWidth, height: 0.35D);
                }
            }
        }

        private void AddCallouts(VisioPage page) {
            for (int i = 0; i < _callouts.Count; i++) {
                CalloutItem callout = _callouts[i];
                BlockItem target = _blocksById[callout.TargetId];
                if (target.Shape == null) {
                    throw new InvalidOperationException("Blocks must be placed before callouts are created.");
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
                ShapeStyle = new VisioShapeStyle(_theme.RegionFill, _theme.RegionStroke, Math.Max(0.012D, _theme.LineWeight)),
                LeaderStyle = new VisioConnectorStyle(_theme.DataFlowColor, Math.Max(0.012D, _theme.LineWeight), 2, EndArrow.None) {
                    Kind = ConnectorKind.RightAngle,
                    TextStyle = _theme.ConnectorTextStyle?.Clone()
                },
                RouteOffset = 0.08D
            };
        }

        private void AddAdornments(VisioPage page) {
            if (!string.IsNullOrWhiteSpace(_titleText)) {
                double y = _pageHeight - _topMargin - (_titleHeight / 2D);
                VisioShape title = page.AddTextBox(_titleId, _pageWidth / 2D, y, Math.Max(1D, _pageWidth - 1.6D), _titleHeight, _titleText, _unit);
                if (_theme.TitleTextStyle != null) {
                    title.TextStyle = _theme.TitleTextStyle.Clone();
                }
                VisioSemanticUserCells.MarkGeneratedAdornment(title);
            }

            if (_showLegend) {
                double y = _pageHeight - _topMargin - TitleHeaderHeight - (LegendHeaderHeight / 2D);
                AddLegendItem(page, Math.Max(0.8D, _leftMargin), y, _dataFlowLegendLabel, _theme.DataFlowColor, 1);
                AddLegendItem(page, Math.Max(0.8D, _pageWidth - 3.1D), y, _controlFlowLegendLabel, _theme.ControlFlowColor, 2);
            }
        }

        private void AddLegendItem(VisioPage page, double x, double y, string label, Color color, int linePattern) {
            VisioShape sample = page.AddRectangle(x + 0.32D, y, 0.64D, 0.08D, string.Empty, _unit);
            sample.NameU = "Rectangle";
            sample.FillPattern = 0;
            sample.LineColor = color;
            sample.LinePattern = linePattern;
            sample.LineWeight = Math.Max(0.018D, _theme.LineWeight);
            VisioSemanticUserCells.MarkGeneratedAdornment(sample);

            VisioShape text = page.AddTextBox(x + 1.5D, y, 1.4D, 0.28D, label, _unit);
            if (_theme.LegendTextStyle != null) {
                text.TextStyle = _theme.LegendTextStyle.Clone();
            }
            VisioSemanticUserCells.MarkGeneratedAdornment(text);
        }

        private double GridX(int column, int span) {
            double left = _leftMargin + column * (_theme.BlockWidth + _theme.ColumnGap);
            double width = span * _theme.BlockWidth + (span - 1) * _theme.ColumnGap;
            return left + width / 2D;
        }

        private double GridY(int row, int span) {
            double top = _pageHeight - _topMargin - HeaderHeight - row * (_theme.BlockHeight + _theme.RowGap);
            double height = span * _theme.BlockHeight + (span - 1) * _theme.RowGap;
            return top - height / 2D;
        }

        private double HeaderHeight {
            get {
                double height = TitleHeaderHeight + (_showLegend ? LegendHeaderHeight : 0D);
                if (_regions.Any(region => !string.IsNullOrWhiteSpace(region.Text))) {
                    height += VisioNetworkDiagramVisuals.BackgroundZoneCaptionHeaderClearance;
                }

                return height;
            }
        }

        private double TitleHeaderHeight => string.IsNullOrWhiteSpace(_titleText) ? 0D : _titleHeight + _titleGap;

        private const double LegendHeaderHeight = 0.45D;

        private void EnsureKnownBlock(string id, string parameterName) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Block id cannot be null or whitespace.", parameterName);
            }

            string normalizedId = id.Trim();
            if (!_blocksById.ContainsKey(normalizedId)) {
                throw new ArgumentException($"Unknown block id '{normalizedId}'.", parameterName);
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

            if (_blocksById.ContainsKey(id)) {
                return true;
            }

            for (int i = 0; i < _regions.Count; i++) {
                if (string.Equals(_regions[i].Id, id, StringComparison.Ordinal)) {
                    return true;
                }
            }

            for (int i = 0; i < _callouts.Count; i++) {
                if (string.Equals(_callouts[i].Id, id, StringComparison.Ordinal)) {
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

        private VisioStyleTheme CreateCaptionTheme() {
            VisioStyleTheme theme = VisioStyleTheme.Fluent();
            theme.Container = new VisioShapeStyle(_theme.RegionFill, _theme.RegionStroke, _theme.LineWeight) {
                TextStyle = _theme.RegionTextStyle?.Clone()
            };
            return theme;
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

        private static void ValidatePlacement(VisioSide placement, string parameterName) {
            if (placement == VisioSide.Auto || !Enum.IsDefined(typeof(VisioSide), placement)) {
                throw new ArgumentOutOfRangeException(parameterName, "Placement must be Left, Right, Bottom, or Top.");
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
