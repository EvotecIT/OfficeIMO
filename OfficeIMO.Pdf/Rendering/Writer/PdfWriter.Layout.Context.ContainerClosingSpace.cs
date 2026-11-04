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

            /// <summary>Returns the space a child of the given container must leave, where this value is the enclosing space.</summary>
            public ContainerClosingSpace Include(PdfPanelStyle? style) => new(
                Padding + GetUnreservedClosingPadding(style) + (RequiresFullAdvance ? style?.SpacingAfter.ToPoints() ?? 0D : 0D),
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
            RequiresFullContainerClosingAdvance(style)
                ? Math.Max(0D, style!.BottomPadding - style!.FragmentBottomInset.ToPoints()) : 0D;

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
                    trailingSpacing += style.SpacingAfter.ToPoints();
                }
                else if (sequence.Owner is not (SemanticBlock or FlowBlock or LayerBlock or SectionBlock or TableOfContentsBlock)) break;
            }
            return new(padding, requiresFullAdvance);
        }

        // A row's after-space lies inside an enclosing strict container. Each entry
        // is the space owed once the items before that index have been placed.
        private ContainerClosingSpace[] BuildClosingColumnGroupSpaces(List<ColItem> items, double rowSpacingAfter) {
            var spaces = new ContainerClosingSpace[items.Count + 1];
            ContainerClosingSpace outer = GetClosingContainerSpace();
            spaces[items.Count] = outer.RequiresFullAdvance ? new(outer.Padding + rowSpacingAfter, true) : outer;
            for (int index = items.Count - 1; index >= 0; index--) {
                if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (items[index] is ColBookmark) spaces[index] = spaces[index + 1];
                else if (items[index] is ColGroupEnd end) spaces[index] = spaces[index + 1].Include(end.Group.Style);
            }
            return spaces;
        }

        private double GetTextLineFitHeight(IReadOnlyList<RichSeg> line, double advance, bool closesText, double topGap = 0D) =>
            closesText && GetClosingContainerSpace().RequiresFullAdvance ? advance : GetRichLineFrameFitHeight(line, advance, topGap);

        private double MeasureTextBlockAdmissionHeight(IReadOnlyList<List<RichSeg>> lines,
            IReadOnlyList<double> heights, double leading, double spacingAfter, List<double>? topGaps = null) {
            double height = MeasureRichLinesHeight(heights, lines.Count, leading);
            if (lines.Count == 0) return height;
            int last = lines.Count - 1;
            double advance = GetRichLineHeight(heights, last, leading);
            return height - advance + GetClosingContainerSpace().GetTextHeight(lines[last], advance, spacingAfter, topGaps?[last] ?? 0D);
        }

        private double GetClosingTextPadding(double spacingAfter) {
            ContainerClosingSpace closing = GetClosingContainerSpace();
            return closing.Padding + (closing.RequiresFullAdvance ? spacingAfter : 0D);
        }
    }
}
