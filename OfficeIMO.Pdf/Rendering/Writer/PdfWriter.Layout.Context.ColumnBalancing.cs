namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void QueueColumnBalanceRemainder(IPdfBlock remainder, IList<IPdfBlock> blocks, int blockIndex) {
            if (!CanQueueColumnBalanceRemainder(blocks)) return;
            var remaining = new List<IPdfBlock> { remainder };
            for (int index = blockIndex + 1; index < blocks.Count; index++) remaining.Add(blocks[index]);
            activeColumnFlow!.PendingBalanceContent = IncludeColumnBalanceAncestors(new ColumnBalanceContent(remaining), blocks);
        }

        private bool CanQueueColumnBalanceRemainder(IList<IPdfBlock> blocks) =>
            activeColumnFlow is { } scope && scope.Options.BalanceLastPage &&
            scope.ColumnIndex + 1 == scope.Widths.Length &&
            HasColumnBalanceSequencePath(blocks, scope) && scope.Widths.All(value => Math.Abs(value - scope.Widths[0]) <= 0.001D);

        private void QueueColumnBalanceRemainder(List<ColumnBalanceUnit>? units, IPdfBlock continuingBlock,
            IList<IPdfBlock> blocks, int blockIndex) {
            if (units == null || !CanQueueColumnBalanceRemainder(blocks)) return;
            var remaining = new List<IPdfBlock>();
            for (int index = blockIndex + 1; index < blocks.Count; index++) remaining.Add(blocks[index]);
            activeColumnFlow!.PendingBalanceContent = IncludeColumnBalanceAncestors(
                new ColumnBalanceContent(remaining, units, KeepsWithNext(continuingBlock)), blocks);
        }

        /// <summary>Finds a height that packs measured breakable units without losing a line to fractional balancing.</summary>
        private double ResolveBalancedColumnHeight(IReadOnlyList<IPdfBlock> blocks, ColumnFlowScope scope) =>
            ResolveBalancedColumnHeight(new ColumnBalanceContent(blocks), scope);

        private double ResolveBalancedColumnHeight(ColumnBalanceContent content, ColumnFlowScope scope) {
            ColumnFlowScope? previous = columnBalanceMeasurementScope;
            columnBalanceMeasurementScope = scope;
            try { return ResolveBalancedColumnHeightCore(content, scope); }
            finally { columnBalanceMeasurementScope = previous; }
        }

        private double ResolveBalancedColumnHeightCore(ColumnBalanceContent content, ColumnFlowScope scope) {
            double available = scope.Top - scope.ParentOptions.MarginBottom;
            if (!scope.Options.BalanceLastPage || content.Blocks.Any(block => block is ColumnBreakBlock or PageBreakBlock) ||
                scope.Widths.Any(value => Math.Abs(value - scope.Widths[0]) > 0.001D)) return available;
            List<ColumnBalanceUnit>? units = MeasureColumnBalanceContent(content, scope, scope.Widths[0]);
            if (units == null) return available;
            if (units.Count == 0) return available;
            bool Fits(double height) {
                int columns = 1;
                double used = 0D;
                return PackColumnBalanceUnits(units, height, scope.Widths.Length, 0D, ref columns, ref used);
            }
            if (!Fits(available)) return available;
            double low = 0D;
            double high = available;
            for (int iteration = 0; iteration < 24; iteration++) {
                double middle = (low + high) / 2D;
                if (Fits(middle)) high = middle;
                else low = middle;
            }
            return Math.Min(available, Math.Ceiling((high + 0.001D) * 1000D) / 1000D);
        }

        private List<ColumnBalanceUnit>? MeasureColumnBalanceUnits(IPdfBlock block, ColumnFlowScope scope, double frameWidth) {
            if (block is SemanticBlock or FlowBlock or ContainerBlock)
                return MeasureNestedColumnBalanceUnits(block, null, scope, frameWidth);
            if (block is PdfListBlock list)
                return MeasureListColumnBalanceUnits(PrepareListLayout(list, frameWidth, currentOpts.DefaultFontSize, topLevelSpacing: true));
            if (block is TableBlock table) return MeasureTableColumnBalanceUnits(table, scope, frameWidth);
            if (block is RichParagraphBlock paragraph) {
                PdfParagraphStyle? style = EffectiveParagraphStyle(paragraph);
                double size = style?.FontSize ?? currentOpts.DefaultFontSize;
                double leading = GetParagraphLeading(style, size);
                var frame = GetParagraphTextFrame(style, scope.ParentOptions.MarginLeft, frameWidth);
                var wrapped = WrapRichRunsCoreWithFirstLineOrigin(paragraph.Runs, frame.Width, size, ChooseNormal(currentOpts.DefaultFont),
                    leading, frame.FirstLineWidth, frame.FirstLineX - frame.X, GetParagraphTabStopWidth(style), currentOpts,
                    style?.TabStops.ToArray(), lineSpacing: style?.LineSpacing);
                var heights = wrapped.LineHeights.ToList();
                if (heights.Count == 0) heights.Add(leading);
                double spacingBefore = GetParagraphSpacingBefore(style);
                heights[heights.Count - 1] += GetParagraphSpacingAfter(style, leading);
                if (style?.KeepTogether == true && !scope.Options.BalanceKeptParagraphLines || !scope.Options.BalanceParagraphLines)
                    return new() { new(heights.Sum(), spacingBefore: spacingBefore) };
                int orphans = style == null ? 1 : Math.Max(1, ResolveMinimumOrphanLines(style));
                int widows = style == null ? 1 : Math.Max(1, ResolveMinimumWidowLines(style));
                if (orphans + widows > heights.Count) return new() { new(heights.Sum(), spacingBefore: spacingBefore) };
                var units = new List<ColumnBalanceUnit> { new(heights.Take(orphans).Sum(), spacingBefore: spacingBefore) };
                units.AddRange(heights.Skip(orphans).Take(heights.Count - orphans - widows).Select(height => new ColumnBalanceUnit(height)));
                units.Add(new ColumnBalanceUnit(heights.Skip(heights.Count - widows).Sum()));
                if (style?.KeepTogether == true && scope.Options.BalanceKeptParagraphLines)
                    return new() { new(new ColumnBalanceParagraph(units)) };
                return units;
            }
            double? height = MeasureWholeBlockHeight(block, scope.ParentOptions.MarginLeft, frameWidth, currentOpts.DefaultFontSize);
            return height.HasValue ? new List<ColumnBalanceUnit> { new(height.Value) } : null;
        }

        private readonly struct ColumnBalanceUnit {
            public ColumnBalanceUnit(double height, double continuationHeight = 0D, double spacingBefore = 0D) {
                Height = height; ContinuationHeight = continuationHeight; SpacingBefore = spacingBefore;
                RowFragment = null; Container = null; Paragraph = null;
            }
            public ColumnBalanceUnit(ColumnBalanceRowFragment fragment, double spacingBefore = 0D) {
                Height = fragment.Height; ContinuationHeight = fragment.MovedBefore; SpacingBefore = spacingBefore;
                RowFragment = fragment; Container = null; Paragraph = null;
            }
            public ColumnBalanceUnit(ColumnBalanceContainer container) {
                Height = container.Units.Sum(unit => unit.Height) + container.Style.PaddingY * 2D + container.Style.SpacingAfter;
                ContinuationHeight = 0D; SpacingBefore = 0D; RowFragment = null; Container = container; Paragraph = null;
            }
            public ColumnBalanceUnit(ColumnBalanceParagraph paragraph) {
                Height = paragraph.Units.Sum(unit => unit.Height); ContinuationHeight = 0D;
                SpacingBefore = paragraph.Units.Count == 0 ? 0D : paragraph.Units[0].SpacingBefore;
                RowFragment = null; Container = null; Paragraph = paragraph;
            }
            public double Height { get; }
            /// <summary>Repeated content or spacing required when this unit starts a new column.</summary>
            public double ContinuationHeight { get; }
            /// <summary>Paragraph spacing applied only when this unit follows content in the same frame.</summary>
            public double SpacingBefore { get; }
            public ColumnBalanceRowFragment? RowFragment { get; }
            public ColumnBalanceContainer? Container { get; }
            public ColumnBalanceParagraph? Paragraph { get; }
        }

        private sealed class ColumnBalanceContent {
            public ColumnBalanceContent(IReadOnlyList<IPdfBlock> blocks, List<ColumnBalanceUnit>? prefixUnits = null,
                bool prefixKeepsNext = false, ColumnBalanceNestedContent? nestedPrefix = null) {
                Blocks = blocks; PrefixUnits = prefixUnits; PrefixKeepsNext = prefixKeepsNext; NestedPrefix = nestedPrefix;
            }
            public IReadOnlyList<IPdfBlock> Blocks { get; }
            public List<ColumnBalanceUnit>? PrefixUnits { get; }
            public bool PrefixKeepsNext { get; }
            public ColumnBalanceNestedContent? NestedPrefix { get; }
        }
    }
}
