namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocParagraphFormat : IEquatable<LegacyDocParagraphFormat> {
        internal LegacyDocParagraphFormat(
            LegacyDocParagraphAlignment? alignment,
            ushort? styleIndex = null,
            int? spacingBeforeTwips = null,
            int? spacingAfterTwips = null,
            int? lineSpacingTwips = null,
            int? leftIndentTwips = null,
            int? rightIndentTwips = null,
            int? firstLineIndentTwips = null,
            bool? keepLinesTogether = null,
            bool? keepWithNext = null,
            bool? pageBreakBefore = null,
            bool? avoidWidowAndOrphan = null,
            bool? isInTable = null,
            bool? isTableTerminatingParagraph = null,
            IReadOnlyList<LegacyDocTabStop>? tabStops = null,
            IReadOnlyList<int>? tableCellWidthsTwips = null,
            int? tableRowHeightTwips = null,
            bool tableRowHeightIsExact = false,
            bool hasMergedTableCells = false) {
            Alignment = alignment;
            StyleIndex = styleIndex;
            SpacingBeforeTwips = spacingBeforeTwips;
            SpacingAfterTwips = spacingAfterTwips;
            LineSpacingTwips = lineSpacingTwips;
            LeftIndentTwips = leftIndentTwips;
            RightIndentTwips = rightIndentTwips;
            FirstLineIndentTwips = firstLineIndentTwips;
            KeepLinesTogether = keepLinesTogether;
            KeepWithNext = keepWithNext;
            PageBreakBefore = pageBreakBefore;
            AvoidWidowAndOrphan = avoidWidowAndOrphan;
            IsInTable = isInTable;
            IsTableTerminatingParagraph = isTableTerminatingParagraph;
            TabStops = tabStops == null || tabStops.Count == 0
                ? Array.Empty<LegacyDocTabStop>()
                : tabStops.ToArray();
            TableCellWidthsTwips = tableCellWidthsTwips == null || tableCellWidthsTwips.Count == 0
                ? Array.Empty<int>()
                : tableCellWidthsTwips.ToArray();
            TableRowHeightTwips = tableRowHeightTwips;
            TableRowHeightIsExact = tableRowHeightIsExact;
            HasMergedTableCells = hasMergedTableCells;
        }

        internal LegacyDocParagraphAlignment? Alignment { get; }

        internal ushort? StyleIndex { get; }

        internal int? SpacingBeforeTwips { get; }

        internal int? SpacingAfterTwips { get; }

        internal int? LineSpacingTwips { get; }

        internal int? LeftIndentTwips { get; }

        internal int? RightIndentTwips { get; }

        internal int? FirstLineIndentTwips { get; }

        internal bool? KeepLinesTogether { get; }

        internal bool? KeepWithNext { get; }

        internal bool? PageBreakBefore { get; }

        internal bool? AvoidWidowAndOrphan { get; }

        internal bool? IsInTable { get; }

        internal bool? IsTableTerminatingParagraph { get; }

        internal IReadOnlyList<LegacyDocTabStop> TabStops { get; }

        internal IReadOnlyList<int> TableCellWidthsTwips { get; }

        internal int? TableRowHeightTwips { get; }

        internal bool TableRowHeightIsExact { get; }

        internal bool HasMergedTableCells { get; }

        internal bool HasFormatting => Alignment != null
            || StyleIndex != null
            || SpacingBeforeTwips != null
            || SpacingAfterTwips != null
            || LineSpacingTwips != null
            || LeftIndentTwips != null
            || RightIndentTwips != null
            || FirstLineIndentTwips != null
            || KeepLinesTogether != null
            || KeepWithNext != null
            || PageBreakBefore != null
            || AvoidWidowAndOrphan != null
            || IsInTable != null
            || IsTableTerminatingParagraph != null
            || TabStops.Count > 0
            || TableCellWidthsTwips.Count > 0
            || TableRowHeightTwips != null
            || HasMergedTableCells;

        internal static LegacyDocParagraphFormat Default { get; } = new LegacyDocParagraphFormat(null);

        public bool Equals(LegacyDocParagraphFormat other) {
            return Alignment == other.Alignment
                && StyleIndex == other.StyleIndex
                && SpacingBeforeTwips == other.SpacingBeforeTwips
                && SpacingAfterTwips == other.SpacingAfterTwips
                && LineSpacingTwips == other.LineSpacingTwips
                && LeftIndentTwips == other.LeftIndentTwips
                && RightIndentTwips == other.RightIndentTwips
                && FirstLineIndentTwips == other.FirstLineIndentTwips
                && KeepLinesTogether == other.KeepLinesTogether
                && KeepWithNext == other.KeepWithNext
                && PageBreakBefore == other.PageBreakBefore
                && AvoidWidowAndOrphan == other.AvoidWidowAndOrphan
                && IsInTable == other.IsInTable
                && IsTableTerminatingParagraph == other.IsTableTerminatingParagraph
                && TabStopsEqual(TabStops, other.TabStops)
                && TableCellWidthsEqual(TableCellWidthsTwips, other.TableCellWidthsTwips)
                && TableRowHeightTwips == other.TableRowHeightTwips
                && TableRowHeightIsExact == other.TableRowHeightIsExact
                && HasMergedTableCells == other.HasMergedTableCells;
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocParagraphFormat other && Equals(other);
        }

        public override int GetHashCode() {
            int hash = 17;
            hash = (hash * 31) + Alignment.GetHashCode();
            hash = (hash * 31) + StyleIndex.GetHashCode();
            hash = (hash * 31) + SpacingBeforeTwips.GetHashCode();
            hash = (hash * 31) + SpacingAfterTwips.GetHashCode();
            hash = (hash * 31) + LineSpacingTwips.GetHashCode();
            hash = (hash * 31) + LeftIndentTwips.GetHashCode();
            hash = (hash * 31) + RightIndentTwips.GetHashCode();
            hash = (hash * 31) + FirstLineIndentTwips.GetHashCode();
            hash = (hash * 31) + KeepLinesTogether.GetHashCode();
            hash = (hash * 31) + KeepWithNext.GetHashCode();
            hash = (hash * 31) + PageBreakBefore.GetHashCode();
            hash = (hash * 31) + AvoidWidowAndOrphan.GetHashCode();
            hash = (hash * 31) + IsInTable.GetHashCode();
            hash = (hash * 31) + IsTableTerminatingParagraph.GetHashCode();
            hash = (hash * 31) + TableRowHeightTwips.GetHashCode();
            hash = (hash * 31) + TableRowHeightIsExact.GetHashCode();
            hash = (hash * 31) + HasMergedTableCells.GetHashCode();
            foreach (LegacyDocTabStop tabStop in TabStops) {
                hash = (hash * 31) + tabStop.GetHashCode();
            }

            foreach (int width in TableCellWidthsTwips) {
                hash = (hash * 31) + width.GetHashCode();
            }

            return hash;
        }

        private static bool TableCellWidthsEqual(IReadOnlyList<int> first, IReadOnlyList<int> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (first[index] != second[index]) {
                    return false;
                }
            }

            return true;
        }

        private static bool TabStopsEqual(IReadOnlyList<LegacyDocTabStop> first, IReadOnlyList<LegacyDocTabStop> second) {
            if (first.Count != second.Count) {
                return false;
            }

            for (int index = 0; index < first.Count; index++) {
                if (!first[index].Equals(second[index])) {
                    return false;
                }
            }

            return true;
        }
    }
}
