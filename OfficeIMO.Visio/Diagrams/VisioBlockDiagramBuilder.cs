using System;
using System.Collections.Generic;
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

        private readonly VisioDocument _document;
        private readonly string _pageName;
        private readonly List<BlockItem> _blocks = new List<BlockItem>();
        private readonly Dictionary<string, BlockItem> _blocksById = new Dictionary<string, BlockItem>(StringComparer.Ordinal);
        private readonly List<RegionItem> _regions = new List<RegionItem>();
        private readonly List<Link> _links = new List<Link>();
        private VisioBlockDiagramTheme _theme = VisioBlockDiagramTheme.TechnicalBlue();
        private VisioMeasurementUnit _unit = VisioMeasurementUnit.Inches;
        private double _pageWidth = 11;
        private double _pageHeight = 8.5;
        private double _leftMargin = 0.8;
        private double _topMargin = 0.8;
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
            EnsureKnownBlock(fromId, nameof(fromId));
            EnsureKnownBlock(toId, nameof(toId));
            _links.Add(new Link(fromId, toId, kind, label));
            return this;
        }

        private void AddRegions(VisioPage page) {
            for (int i = 0; i < _regions.Count; i++) {
                RegionItem region = _regions[i];
                double x = GridX(region.Column, region.ColumnSpan);
                double y = GridY(region.Row, region.RowSpan);
                double width = region.ColumnSpan * _theme.BlockWidth + (region.ColumnSpan - 1) * _theme.ColumnGap + 0.7;
                double height = region.RowSpan * _theme.BlockHeight + (region.RowSpan - 1) * _theme.RowGap + 0.55;
                VisioShape shape = new VisioShape(region.Id, x, y, width, height, region.Text) { NameU = "Rectangle" };
                shape.FillColor = _theme.RegionFill;
                shape.LineColor = _theme.RegionStroke;
                shape.LineWeight = _theme.LineWeight;
                if (_theme.RegionTextStyle != null) {
                    shape.TextStyle = _theme.RegionTextStyle.Clone();
                }

                shape.Master = _document.EnsureBuiltinMaster("Rectangle");
                shape.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.BackgroundSurfaceKind, "STR", prompt: "OfficeIMO semantic kind");
                page.Shapes.Add(shape);
            }
        }

        private void AddBlocks(VisioPage page) {
            for (int i = 0; i < _blocks.Count; i++) {
                BlockItem block = _blocks[i];
                double x = GridX(block.Column, 1);
                double y = GridY(block.Row, 1);
                VisioShape shape = CreateBlockShape(page, block, x, y);
                string? nameU = shape.NameU;
                if (!string.IsNullOrWhiteSpace(nameU)) {
                    shape.Master = _document.EnsureBuiltinMaster(nameU!);
                }

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

            VisioShape shape = new VisioShape(block.Id, x, y, _theme.BlockWidth, height, block.Text) { NameU = nameU };
            ApplyBlockStyle(shape, block.Emphasis);
            page.Shapes.Add(shape);
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
            }
        }

        private double GridX(int column, int span) {
            double left = _leftMargin + column * (_theme.BlockWidth + _theme.ColumnGap);
            double width = span * _theme.BlockWidth + (span - 1) * _theme.ColumnGap;
            return left + width / 2D;
        }

        private double GridY(int row, int span) {
            double top = _pageHeight - _topMargin - row * (_theme.BlockHeight + _theme.RowGap);
            double height = span * _theme.BlockHeight + (span - 1) * _theme.RowGap;
            return top - height / 2D;
        }

        private void EnsureKnownBlock(string id, string parameterName) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Block id cannot be null or whitespace.", parameterName);
            }

            if (!_blocksById.ContainsKey(id)) {
                throw new ArgumentException($"Unknown block id '{id}'.", parameterName);
            }
        }

        private static string RequireId(string id, string parameterName, string label) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException(label + " cannot be null or whitespace.", parameterName);
            }

            return id;
        }

        private bool IsIdInUse(string id) {
            if (_blocksById.ContainsKey(id)) {
                return true;
            }

            for (int i = 0; i < _regions.Count; i++) {
                if (string.Equals(_regions[i].Id, id, StringComparison.Ordinal)) {
                    return true;
                }
            }

            return false;
        }

        private static void ValidateGridPosition(int column, int row) {
            if (column < 0) throw new ArgumentOutOfRangeException(nameof(column), "Column must be zero or greater.");
            if (row < 0) throw new ArgumentOutOfRangeException(nameof(row), "Row must be zero or greater.");
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
