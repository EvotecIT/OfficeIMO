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
            bool? isTableTerminatingParagraph = null) {
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
            || IsTableTerminatingParagraph != null;

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
                && IsTableTerminatingParagraph == other.IsTableTerminatingParagraph;
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
            return hash;
        }
    }
}
