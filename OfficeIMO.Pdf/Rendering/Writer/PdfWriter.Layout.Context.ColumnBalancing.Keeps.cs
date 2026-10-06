namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Uses the balancer's leading group in a shortened frame and ordinary whole-chain reservation on physical frames.</summary>
        private double MeasureCurrentFrameKeepNextHeight(IList<IPdfBlock> blocks, int startIndex, double frameX,
            double frameWidth, double fontSize, double precedingHeight) {
            if (ReservesWholeKeepGroup())
                return MeasureKeepWithNextChainHeight(blocks, startIndex, frameX, frameWidth, fontSize, precedingHeight);
            ColumnFlowScope scope = activeColumnFlow!;
            if (!scope.Options.HonorKeepWithNextWhenBalancing) return 0D;
            double height = 0D;
            int inspected = 0;
            for (int index = startIndex; index < blocks.Count; index++) {
                if (IsNonVisualFlowMarker(blocks[index])) continue;
                if (++inspected > MaxKeepWithNextChainBlocks)
                    throw new NotSupportedException("KeepWithNext chains cannot contain more than " + MaxKeepWithNextChainBlocks + " visual blocks.");
                List<ColumnBalanceUnit>? units = MeasureColumnBalanceUnits(blocks[index], scope, frameWidth);
                if (units == null || units.Count == 0) return height;
                height += units[0].Height + (precedingHeight + height > .001D ? units[0].SpacingBefore : 0D);
                if (units.Count > 1 || !KeepsWithNext(blocks[index])) break;
            }
            return height;
        }

        /// <summary>Moves a paragraph/list's terminal group with the next leading unit without reserving its already splittable prefix.</summary>
        private void ReserveBalancedKeepNextTail(IList<double> lineHeights, int lineIndex, int lineCount,
            double spacingAfter, double available, IList<IPdfBlock> blocks, int blockIndex, bool keepsNext,
            int minimumTailLines, ref int take, ref double heightSum) {
            if (!keepsNext || ReservesWholeKeepGroup() || take == 0 || lineIndex + take != lineCount) return;
            double nextHeight = MeasureCurrentFrameKeepNextHeight(blocks, blockIndex + 1, currentOpts.MarginLeft,
                width, currentOpts.DefaultFontSize, heightSum + spacingAfter);
            int tail = Math.Min(take, Math.Max(1, minimumTailLines));
            double tailHeight = lineHeights.Skip(lineCount - tail).Sum();
            double groupHeight = tailHeight + spacingAfter + nextHeight;
            if (nextHeight <= .001D || heightSum + spacingAfter + nextHeight <= available + .001D ||
                groupHeight > GetCurrentFramePageStartY() - currentOpts.MarginBottom + .001D) return;
            take -= tail;
            heightSum -= tailHeight;
        }
    }
}
