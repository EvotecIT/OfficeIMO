namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        // A strict final decoration closes after the consumed text advance, even
        // when padding is zero or already reserved by the fragment's bottom inset.
        private readonly struct ContainerClosingSpace {
            public ContainerClosingSpace(double padding, bool requiresFullAdvance) {
                Padding = padding;
                RequiresFullAdvance = requiresFullAdvance;
            }

            public double Padding { get; }
            public bool RequiresFullAdvance { get; }

            public ContainerClosingSpace Include(PdfPanelStyle style) => new(
                Padding + GetUnreservedClosingPadding(style) + (RequiresFullAdvance ? style.SpacingAfter : 0D),
                RequiresFullAdvance || RequiresFullContainerClosingAdvance(style));

            public double GetTextHeight(IReadOnlyList<RichSeg> line, double advance, double spacingAfter, double topGap = 0D) =>
                (RequiresFullAdvance ? advance : GetRichLineFrameFitHeight(line, advance, topGap)) + Padding +
                (RequiresFullAdvance ? spacingAfter : 0D);
        }

        private static bool RequiresFullContainerClosingAdvance(PdfPanelStyle? style) =>
            style is { RepeatFragmentDecoration: false, AllowTrailingPaddingOverflow: false };

        // Intermediate fragments keep their full capacity; an inset may already
        // reserve some or all of the final padding without changing its advance.
        private static double GetUnreservedClosingPadding(PdfPanelStyle? style) =>
            style is { RepeatFragmentDecoration: false, AllowTrailingPaddingOverflow: false }
                ? Math.Max(0D, style.BottomPadding - style.FragmentBottomInset) : 0D;

        private double GetClosingContainerPadding(int? pendingSequenceDepth = null) =>
            GetClosingContainerSpace(pendingSequenceDepth).Padding;

        private ContainerClosingSpace GetClosingContainerSpace(int? pendingSequenceDepth = null) {
            if (!pendingSequenceDepth.HasValue && activeBlockSequences.Count > 0) {
                BlockSequenceScope leaf = activeBlockSequences[activeBlockSequences.Count - 1];
                // A section heading precedes its body. Its enclosing panel is
                // not closing until the section's final child is placed.
                if (leaf.Blocks[leaf.Index] is SectionBlock section && section.Blocks.Count > 0) return default;
            }
            double padding = 0D;
            double trailingSpacing = 0D;
            bool requiresFullAdvance = false;
            for (int index = pendingSequenceDepth ?? activeBlockSequences.Count - 1; index >= 0; index--) {
                BlockSequenceScope sequence = activeBlockSequences[index];
                // Pending admission checks the selected leaf's remaining
                // siblings itself; ancestor indices still name active children.
                if (index != pendingSequenceDepth) {
                    for (int next = sequence.Index + 1; next < sequence.Blocks.Count; next++) {
                        if (!IsNonVisualFlowMarker(sequence.Blocks[next])) return new(padding, requiresFullAdvance);
                    }
                }
                if (sequence.Owner is ContainerBlock container) {
                    PdfPanelStyle style = ResolveContainerStyle(container);
                    double closing = GetUnreservedClosingPadding(style);
                    if (RequiresFullContainerClosingAdvance(style)) {
                        requiresFullAdvance = true;
                        padding += closing + trailingSpacing;
                        trailingSpacing = 0D;
                    }
                    trailingSpacing += style.SpacingAfter;
                }
                else if (sequence.Owner is not (SemanticBlock or FlowBlock or LayerBlock or SectionBlock or TableOfContentsBlock)) break;
            }
            return new(padding, requiresFullAdvance);
        }

        private double GetTextLineAdmissionHeight(IReadOnlyList<RichSeg> line, double advance,
            bool closesText, double spacingAfter, double topGap = 0D) =>
            (closesText ? GetClosingContainerSpace() : default).GetTextHeight(line, advance, spacingAfter, topGap);

        private double MeasureTextBlockAdmissionHeight(IReadOnlyList<List<RichSeg>> lines,
            IReadOnlyList<double> heights, double leading, double spacingAfter, List<double>? topGaps = null) {
            double height = MeasureRichLinesHeight(heights, lines.Count, leading);
            if (lines.Count == 0) return height;
            int last = lines.Count - 1;
            double advance = GetRichLineHeight(heights, last, leading);
            return height - advance + GetTextLineAdmissionHeight(lines[last], advance, closesText: true,
                spacingAfter, topGaps?[last] ?? 0D);
        }

        private double GetClosingColumnGroupPadding(List<ColItem> items, int itemIndex, double rowSpacingAfter = 0D) {
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
            ContainerClosingSpace outerClosing = GetClosingContainerSpace();
            return padding + outerClosing.Padding + (outerClosing.Padding > 0D ? trailingSpacing : 0D) +
                (outerClosing.RequiresFullAdvance ? rowSpacingAfter : 0D);
        }

        private double GetClosingTextPadding(double spacingAfter) {
            ContainerClosingSpace closing = GetClosingContainerSpace();
            return closing.Padding + (closing.RequiresFullAdvance ? spacingAfter : 0D);
        }
    }
}
