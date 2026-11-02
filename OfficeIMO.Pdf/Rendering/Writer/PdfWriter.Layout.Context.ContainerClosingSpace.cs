namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        // Only the last visual content needs closing space. Intermediate fragments
        // keep their full capacity; an inset may already reserve some or all of it.
        private static double GetUnreservedClosingPadding(PdfPanelStyle? style) =>
            style is { RepeatFragmentDecoration: false, AllowTrailingPaddingOverflow: false }
                ? Math.Max(0D, style.BottomPadding - style.FragmentBottomInset) : 0D;

        private double GetClosingContainerPadding(int? pendingSequenceDepth = null) {
            if (!pendingSequenceDepth.HasValue && activeBlockSequences.Count > 0) {
                BlockSequenceScope leaf = activeBlockSequences[activeBlockSequences.Count - 1];
                // A section heading precedes its body. Its enclosing panel is
                // not closing until the section's final child is placed.
                if (leaf.Blocks[leaf.Index] is SectionBlock section && section.Blocks.Count > 0) return 0D;
            }
            double padding = 0D;
            double trailingSpacing = 0D;
            for (int index = pendingSequenceDepth ?? activeBlockSequences.Count - 1; index >= 0; index--) {
                BlockSequenceScope sequence = activeBlockSequences[index];
                // Pending admission checks the selected leaf's remaining
                // siblings itself; ancestor indices still name active children.
                if (index != pendingSequenceDepth) {
                    for (int next = sequence.Index + 1; next < sequence.Blocks.Count; next++) {
                        if (!IsNonVisualFlowMarker(sequence.Blocks[next])) return padding;
                    }
                }
                if (sequence.Owner is ContainerBlock container) {
                    PdfPanelStyle style = ResolveContainerStyle(container);
                    double closing = GetUnreservedClosingPadding(style);
                    if (closing > 0D) { padding += closing + trailingSpacing; trailingSpacing = 0D; }
                    trailingSpacing += style.SpacingAfter;
                }
                else if (sequence.Owner is not (SemanticBlock or FlowBlock or LayerBlock or SectionBlock or TableOfContentsBlock)) break;
            }
            return padding;
        }

        private double[] BuildClosingColumnGroupPaddings(List<ColItem> items) {
            var paddings = new double[items.Count + 1];
            double outerClosing = GetClosingContainerPadding();
            paddings[items.Count] = outerClosing;
            bool absorbsTrailingSpacing = outerClosing > 0D;
            for (int index = items.Count - 1; index >= 0; index--) {
                if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (items[index] is ColBookmark) {
                    paddings[index] = paddings[index + 1];
                } else if (items[index] is ColGroupEnd end) {
                    double closing = GetUnreservedClosingPadding(end.Group.Style);
                    double spacing = end.Group.Style?.SpacingAfter ?? 0D;
                    paddings[index] = closing + paddings[index + 1] +
                        (absorbsTrailingSpacing ? spacing : 0D);
                    absorbsTrailingSpacing |= closing > 0D;
                } else {
                    paddings[index] = 0D;
                    absorbsTrailingSpacing = false;
                }
            }
            return paddings;
        }

        private double GetClosingTextPadding(double spacingAfter) {
            double padding = GetClosingContainerPadding();
            return padding > 0D ? padding + spacingAfter : 0D;
        }
    }
}
