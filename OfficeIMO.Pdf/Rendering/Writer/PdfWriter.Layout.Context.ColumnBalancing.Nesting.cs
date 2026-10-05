namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly List<BlockSequenceScope> activeBlockSequences = new();

        /// <summary>Retains the actual sequence path so a leaf continuation includes unfinished wrappers and their following siblings.</summary>
        private bool HasColumnBalanceSequencePath(IList<IPdfBlock> blocks, ColumnFlowScope scope) {
            bool foundLeaf = false;
            for (int index = activeBlockSequences.Count - 1; index >= 0; index--) {
                BlockSequenceScope sequence = activeBlockSequences[index];
                if (!foundLeaf) {
                    if (!ReferenceEquals(sequence.Blocks, blocks)) continue;
                    foundLeaf = true;
                }
                if (ReferenceEquals(sequence.Blocks, scope.Blocks)) return true;
                if (sequence.Owner is not (ContainerBlock or SemanticBlock or FlowBlock)) return false;
            }
            return false;
        }

        private ColumnBalanceContent IncludeColumnBalanceAncestors(ColumnBalanceContent content, IList<IPdfBlock> leaf) {
            int index = activeBlockSequences.FindLastIndex(sequence => ReferenceEquals(sequence.Blocks, leaf));
            while (index >= 0 && !ReferenceEquals(activeBlockSequences[index].Blocks, activeColumnFlow!.Blocks)) {
                IPdfBlock owner = activeBlockSequences[index].Owner!;
                BlockSequenceScope parent = activeBlockSequences[--index];
                var siblings = parent.Blocks.Skip(parent.Index + 1).ToArray();
                content = new ColumnBalanceContent(siblings, prefixKeepsNext: KeepsWithNext(owner),
                    nestedPrefix: new ColumnBalanceNestedContent(owner, content));
            }
            return content;
        }

        private List<ColumnBalanceUnit>? MeasureColumnBalanceContent(ColumnBalanceContent content, ColumnFlowScope scope, double frameWidth) {
            var units = content.NestedPrefix is { } nested
                ? MeasureNestedColumnBalanceUnits(nested.Owner, nested.Content, scope, frameWidth)
                : content.PrefixUnits == null ? new List<ColumnBalanceUnit>() : new List<ColumnBalanceUnit>(content.PrefixUnits);
            if (units == null) return null;
            bool previousKeepsNext = content.PrefixKeepsNext;
            foreach (IPdfBlock block in content.Blocks) {
                if (IsNonVisualFlowMarker(block)) continue;
                List<ColumnBalanceUnit>? blockUnits = MeasureColumnBalanceUnits(block, scope, frameWidth);
                if (blockUnits == null) return null;
                if (scope.Options.HonorKeepWithNextWhenBalancing && previousKeepsNext && units.Count > 0 && blockUnits.Count > 0) {
                    units[units.Count - 1] = JoinColumnBalanceUnits(units[units.Count - 1], blockUnits[0]);
                    blockUnits.RemoveAt(0);
                }
                units.AddRange(blockUnits);
                previousKeepsNext = KeepsWithNext(block);
            }
            return units;
        }

        /// <summary>Measures children at their real container width; semantic and unconstrained static flow remain transparent.</summary>
        private List<ColumnBalanceUnit>? MeasureNestedColumnBalanceUnits(IPdfBlock wrapper, ColumnBalanceContent? remainder,
            ColumnFlowScope scope, double frameWidth) {
            if (wrapper is SemanticBlock semantic)
                return MeasureColumnBalanceContent(remainder ?? new ColumnBalanceContent(semantic.Blocks), scope, frameWidth);
            if (wrapper is FlowBlock flow) {
                if (!PdfFlowNestingRules.IsColumnFlowSupported(flow)) return null;
                if (flow.Options.KeepTogether) {
                    double? keptHeight = MeasureWholeBlockHeight(flow, scope.ParentOptions.MarginLeft, frameWidth, currentOpts.DefaultFontSize);
                    return keptHeight.HasValue ? new() { new(keptHeight.Value) } : null;
                }
                List<ColumnBalanceUnit>? children = MeasureColumnBalanceContent(remainder ?? new ColumnBalanceContent(flow.StaticBlocks!), scope, frameWidth);
                return children;
            }
            var container = (ContainerBlock)wrapper;
            PdfPanelStyle style = ResolveContainerStyle(container);
            var frame = ResolveContainerFrame(style, scope.ParentOptions.MarginLeft, frameWidth);
            if (style.KeepTogether) {
                double? keptHeight = MeasureWholeBlockHeight(container, frame.X, frameWidth, currentOpts.DefaultFontSize);
                return keptHeight.HasValue ? new() { new(keptHeight.Value) } : null;
            }
            List<ColumnBalanceUnit>? units = MeasureWithContainerPaddingReservation(style.PaddingY, () =>
                MeasureColumnBalanceContent(remainder ?? new ColumnBalanceContent(container.Blocks), scope, frame.ContentWidth));
            if (units == null) return null;
            double firstHeight = remainder != null || container.Blocks.Count == 0 ? 0D :
                MeasureWithContainerPaddingReservation(style.PaddingY, () => MeasureNextBlockFirstVisualHeight(
                    container.Blocks[0], frame.X + style.PaddingX, frame.ContentWidth, currentOpts.DefaultFontSize, allowTableFragments: true));
            return new() { new(new ColumnBalanceContainer(style, units, firstHeight, remainder != null)) };
        }

        private static bool PackColumnBalanceUnits(IReadOnlyList<ColumnBalanceUnit> units, double height, int columnCount,
            double continuationPadding, ref int columns, ref double used) {
            foreach (ColumnBalanceUnit unit in units) {
                if (unit.Container is { } container) {
                    if (!PackColumnBalanceContainer(container, height, columnCount, continuationPadding, ref columns, ref used)) return false;
                } else if (unit.Paragraph is { } paragraph) {
                    if (!PackColumnBalanceParagraph(paragraph, height, columnCount, continuationPadding, ref columns, ref used)) return false;
                } else if (unit.RowFragment is { } row) {
                    if (!PackColumnBalanceRowFragment(row, height, columnCount, ref columns, ref used, continuationPadding)) return false;
                } else {
                    if (used + unit.Height > height + .001D) { columns++; used = continuationPadding + unit.ContinuationHeight; }
                    if (columns > columnCount || used + unit.Height > height + .001D) return false;
                    used += unit.Height;
                }
            }
            return true;
        }

        /// <summary>Matches container fragment starts, repeated top padding and bounded bottom padding during column packing.</summary>
        private static bool PackColumnBalanceContainer(ColumnBalanceContainer container, double height, int columnCount,
            double parentPadding, ref int columns, ref double used) {
            PdfPanelStyle style = container.Style;
            double before = container.IsContinuation || used <= parentPadding + .001D ? 0D : style.SpacingBefore;
            if (!container.IsContinuation) {
                double minimumStart = style.PaddingY * 2D + container.FirstVisualHeight;
                if (minimumStart > height - parentPadding + .001D) return false;
                if (used > parentPadding + .001D && used + before + minimumStart > height + .001D) {
                    if (++columns > columnCount) return false;
                    used = parentPadding; before = 0D;
                }
            }
            used += before + Math.Min(style.PaddingY, Math.Max(0D, height - used));
            if (!PackColumnBalanceUnits(container.Units, height, columnCount, parentPadding + style.PaddingY, ref columns, ref used)) return false;
            used += Math.Min(style.PaddingY, Math.Max(0D, height - used));
            double after = style.SpacingAfter;
            while (after > .001D) {
                double take = Math.Min(after, Math.Max(0D, height - used));
                used += take; after -= take;
                if (after <= .001D) break;
                if (++columns > columnCount) return false;
                used = parentPadding;
                if (height <= parentPadding + .001D) return false;
            }
            return true;
        }

        private sealed class ColumnBalanceContainer {
            public ColumnBalanceContainer(PdfPanelStyle style, List<ColumnBalanceUnit> units, double firstVisualHeight, bool isContinuation) {
                Style = style; Units = units; FirstVisualHeight = firstVisualHeight; IsContinuation = isContinuation;
            }
            public PdfPanelStyle Style { get; }
            public List<ColumnBalanceUnit> Units { get; }
            public double FirstVisualHeight { get; }
            public bool IsContinuation { get; }
        }

        private sealed class ColumnBalanceNestedContent {
            public ColumnBalanceNestedContent(IPdfBlock owner, ColumnBalanceContent content) { Owner = owner; Content = content; }
            public IPdfBlock Owner { get; }
            public ColumnBalanceContent Content { get; }
        }

        private sealed class BlockSequenceScope {
            public BlockSequenceScope(IList<IPdfBlock> blocks, IPdfBlock? owner) { Blocks = blocks; Owner = owner; }
            public IList<IPdfBlock> Blocks { get; }
            public IPdfBlock? Owner { get; }
            public int Index { get; set; }
        }
    }
}
