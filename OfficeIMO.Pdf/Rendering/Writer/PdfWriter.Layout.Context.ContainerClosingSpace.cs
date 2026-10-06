namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        // Only the last visual content needs closing space. Intermediate fragments
        // keep their full capacity; an inset may already reserve some or all of it.
        private static double GetUnreservedClosingPadding(PdfPanelStyle? style) =>
            style is { RepeatFragmentDecoration: false } ? Math.Max(0D, style.BottomPadding - style.FragmentBottomInset) : 0D;

        private double GetClosingContainerPadding() {
            if (activeBlockSequences.Count > 0) {
                BlockSequenceScope leaf = activeBlockSequences[activeBlockSequences.Count - 1];
                // A section heading precedes its body. Its enclosing panel is
                // not closing until the section's final child is placed.
                if (leaf.Blocks[leaf.Index] is SectionBlock section && section.Blocks.Count > 0) return 0D;
            }
            double padding = 0D;
            double trailingSpacing = 0D;
            for (int index = activeBlockSequences.Count - 1; index >= 0; index--) {
                BlockSequenceScope sequence = activeBlockSequences[index];
                for (int next = sequence.Index + 1; next < sequence.Blocks.Count; next++) {
                    if (!IsNonVisualFlowMarker(sequence.Blocks[next])) return padding;
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

        private double GetClosingColumnGroupPadding(List<ColItem> items, int itemIndex) {
            double padding = 0D;
            double trailingSpacing = 0D;
            for (int next = itemIndex + 1; next < items.Count; next++) {
                if (items[next] is ColGroupEnd end) {
                    double closing = GetUnreservedClosingPadding(end.Group.Style);
                    if (closing > 0D) { padding += closing + trailingSpacing; trailingSpacing = 0D; }
                    trailingSpacing += end.Group.Style?.SpacingAfter ?? 0D;
                }
                else if (items[next] is not ColBookmark) return padding;
            }
            double outerClosing = GetClosingContainerPadding();
            return padding + outerClosing + (outerClosing > 0D ? trailingSpacing : 0D);
        }

        private double GetClosingTextPadding(double spacingAfter) {
            double padding = GetClosingContainerPadding();
            return padding > 0D ? padding + spacingAfter : 0D;
        }
    }
}
