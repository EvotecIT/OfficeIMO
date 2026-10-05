namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private ColumnFlowScope? activeColumnFlow;

        /// <summary>Uses ordinary block flow in each column frame so all supported blocks share their existing continuation and keep rules.</summary>
        private void RenderMultiColumnBlock(MultiColumnBlock columns) {
            if (activeColumnFlow != null) throw new NotSupportedException("Columns blocks cannot be nested inside another Columns block.");
            ValidateSequentialColumnBlocks(columns.Blocks);
            if (HasFloatingTables) AvoidFloatingBlock(Math.Max(1D, y - currentOpts.MarginBottom));
            if (y - currentOpts.MarginBottom <= 0.5D) NewPage();
            PdfMultiColumnOptions options = columns.Options;
            double[] gaps = Enumerable.Range(0, options.ColumnCount - 1)
                .Select(index => options.ColumnDefinitions.Count > 0 ? options.ColumnDefinitions[index].GapAfter ?? options.Gap : options.Gap).ToArray();
            double columnArea = width - gaps.Sum();
            if (columnArea <= 0D) throw new ArgumentException("Multi-column gaps must leave positive column widths.");
            var sizing = new RowBlock();
            for (int index = 0; index < options.ColumnCount; index++)
                sizing.AddColumn(new RowColumn(options.ColumnDefinitions.Count > 0 ? options.ColumnDefinitions[index].Width : PdfColumnWidth.Relative()));
            double[] widths = ResolveRowColumnWidths(sizing, columnArea);
            var scope = new ColumnFlowScope(options, columns.Blocks, currentOpts, width, yStart, y, widths, gaps, activeContainerScopes.Count);
            scope.TargetHeight = ResolveBalancedColumnHeight(columns.Blocks, scope);
            activeColumnFlow = scope;
            EnterColumnFrame(scope);
            try {
                ProcessBlocks(columns.Blocks, columns);
                // Terminal spacing follows the last fragment without changing the content's balanced breakpoints.
                if (y >= scope.Top - 0.001D) y = scope.LastOccupiedColumnBottom;
                y -= Math.Min(options.FinalColumnSpacingAfter, Math.Max(0D, y - scope.ParentOptions.MarginBottom));
                FinishColumnFrame(scope);
                DrawColumnFlowSeparators(scope);
            } finally {
                scope.ParentOptions.MergeFontProgramUsageFrom(scope.FrameOptions);
                currentOpts = scope.ParentOptions;
                width = scope.ParentWidth;
                yStart = scope.ParentYStart;
                y = scope.LowestY;
                activeColumnFlow = null;
            }
        }

        private static void ValidateSequentialColumnBlocks(IReadOnlyList<IPdfBlock> blocks) {
            foreach (IPdfBlock block in blocks) {
                if (block is ColumnBreakBlock or PageBreakBlock || PdfFlowNestingRules.IsColumnFlowSupported(block)) continue;
                throw new NotSupportedException("Sequential column flow does not support block type " + block.GetType().Name + ".");
            }
        }

        private void EnterColumnFrame(ColumnFlowScope scope) {
            double left = scope.ParentOptions.MarginLeft;
            for (int index = 0; index < scope.ColumnIndex; index++) left += scope.Widths[index] + scope.Gaps[index];
            // Reuse one options object. Asset and font maps must not be copied for every continuation.
            scope.FrameOptions.MarginLeft = left;
            scope.FrameOptions.MarginRight = scope.FrameOptions.PageWidth - left - scope.Widths[scope.ColumnIndex];
            scope.FrameOptions.MarginTop = scope.FrameOptions.PageHeight - scope.Top;
            scope.FrameOptions.MarginBottom = Math.Max(scope.ParentOptions.MarginBottom, scope.Top - scope.TargetHeight);
            currentOpts = scope.FrameOptions;
            width = scope.Widths[scope.ColumnIndex];
            yStart = scope.Top;
            y = scope.Top;
        }

        private void FinishColumnFrame(ColumnFlowScope scope) {
            if (y < scope.Top - 0.001D) scope.LastOccupiedColumnBottom = y;
            scope.LowestY = Math.Min(scope.LowestY, y);
            scope.ParentOptions.MergeFontProgramUsageFrom(scope.FrameOptions);
        }

        /// <summary>Overflow advances in reading order; an explicit page break skips the remaining columns.</summary>
        private void AdvanceColumnFrame(bool forcePhysicalPage, bool preserveEmptyPage = false) {
            cancellationToken.ThrowIfCancellationRequested();
            ColumnFlowScope scope = activeColumnFlow ?? throw new InvalidOperationException("There is no active column frame.");
            PrepareActiveContainerScopesForPageBreak(scope.ContainerDepth);
            FinishColumnFrame(scope);
            currentOpts = scope.ParentOptions;
            width = scope.ParentWidth;
            yStart = scope.ParentYStart;
            if (!forcePhysicalPage && scope.ColumnIndex + 1 < scope.Widths.Length) {
                scope.ColumnIndex++;
                EnterColumnFrame(scope);
                ResumeActiveContainerScopesOnNewPage(scope.ContainerDepth);
                return;
            }
            DrawColumnFlowSeparators(scope);
            y = scope.LowestY;
            PrepareActiveContainerScopesForPageBreak(0, scope.ContainerDepth);
            FlushPage(preserveEmptyPage || pageDirty || HasCurrentPageNonContentObjects());
            StartPage(currentPageBaseOptions);
            ResumeActiveContainerScopesOnNewPage(0, scope.ContainerDepth);
            scope.ParentOptions = currentOpts;
            scope.ParentWidth = width;
            scope.ParentYStart = yStart;
            scope.Top = y;
            scope.LowestY = y;
            scope.LastOccupiedColumnBottom = y;
            scope.TargetHeight = scope.PendingBalanceContent == null ? y - currentOpts.MarginBottom :
                ResolveBalancedColumnHeight(scope.PendingBalanceContent, scope);
            scope.PendingBalanceContent = null;
            scope.ColumnIndex = 0;
            EnterColumnFrame(scope);
            ResumeActiveContainerScopesOnNewPage(scope.ContainerDepth);
        }

        private void DrawColumnFlowSeparators(ColumnFlowScope scope) {
            if (scope.Options.SeparatorColor is not { } color || scope.Options.SeparatorWidth <= 0D || scope.Top - scope.LowestY <= 0.001D) return;
            double x = scope.ParentOptions.MarginLeft;
            for (int index = 0; index < scope.Widths.Length - 1; index++) {
                x += scope.Widths[index];
                DrawVLine(sb, color, scope.Options.SeparatorWidth, x + scope.Gaps[index] / 2D, scope.Top, scope.LowestY, emitGeneratedStructure);
                x += scope.Gaps[index];
            }
            pageDirty = true;
        }

        private sealed class ColumnFlowScope {
            public ColumnFlowScope(PdfMultiColumnOptions options, IReadOnlyList<IPdfBlock> blocks, PdfOptions parent, double parentWidth, double parentYStart,
                double top, double[] widths, double[] gaps, int containerDepth) {
                Options = options; Blocks = blocks; ParentOptions = parent; FrameOptions = parent.Clone();
                ParentWidth = parentWidth; ParentYStart = parentYStart; Top = top; LowestY = top; LastOccupiedColumnBottom = top;
                Widths = widths; Gaps = gaps; ContainerDepth = containerDepth;
            }
            public PdfMultiColumnOptions Options { get; }
            public IReadOnlyList<IPdfBlock> Blocks { get; }
            public ColumnBalanceContent? PendingBalanceContent { get; set; }
            public PdfOptions ParentOptions { get; set; }
            public PdfOptions FrameOptions { get; }
            public double ParentWidth { get; set; }
            public double ParentYStart { get; set; }
            public double Top { get; set; }
            public double LowestY { get; set; }
            public double LastOccupiedColumnBottom { get; set; }
            public double TargetHeight { get; set; }
            public double[] Widths { get; }
            public double[] Gaps { get; }
            public int ContainerDepth { get; }
            public int ColumnIndex { get; set; }
        }
    }
}
