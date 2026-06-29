namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocParagraphFormat : IEquatable<LegacyDocParagraphFormat> {
        internal LegacyDocParagraphFormat(
            LegacyDocParagraphAlignment? alignment,
            int? spacingBeforeTwips = null,
            int? spacingAfterTwips = null,
            int? lineSpacingTwips = null,
            int? leftIndentTwips = null,
            int? rightIndentTwips = null,
            int? firstLineIndentTwips = null) {
            Alignment = alignment;
            SpacingBeforeTwips = spacingBeforeTwips;
            SpacingAfterTwips = spacingAfterTwips;
            LineSpacingTwips = lineSpacingTwips;
            LeftIndentTwips = leftIndentTwips;
            RightIndentTwips = rightIndentTwips;
            FirstLineIndentTwips = firstLineIndentTwips;
        }

        internal LegacyDocParagraphAlignment? Alignment { get; }

        internal int? SpacingBeforeTwips { get; }

        internal int? SpacingAfterTwips { get; }

        internal int? LineSpacingTwips { get; }

        internal int? LeftIndentTwips { get; }

        internal int? RightIndentTwips { get; }

        internal int? FirstLineIndentTwips { get; }

        internal bool HasFormatting => Alignment != null
            || SpacingBeforeTwips != null
            || SpacingAfterTwips != null
            || LineSpacingTwips != null
            || LeftIndentTwips != null
            || RightIndentTwips != null
            || FirstLineIndentTwips != null;

        internal static LegacyDocParagraphFormat Default { get; } = new LegacyDocParagraphFormat(null);

        public bool Equals(LegacyDocParagraphFormat other) {
            return Alignment == other.Alignment
                && SpacingBeforeTwips == other.SpacingBeforeTwips
                && SpacingAfterTwips == other.SpacingAfterTwips
                && LineSpacingTwips == other.LineSpacingTwips
                && LeftIndentTwips == other.LeftIndentTwips
                && RightIndentTwips == other.RightIndentTwips
                && FirstLineIndentTwips == other.FirstLineIndentTwips;
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocParagraphFormat other && Equals(other);
        }

        public override int GetHashCode() {
            int hash = 17;
            hash = (hash * 31) + Alignment.GetHashCode();
            hash = (hash * 31) + SpacingBeforeTwips.GetHashCode();
            hash = (hash * 31) + SpacingAfterTwips.GetHashCode();
            hash = (hash * 31) + LineSpacingTwips.GetHashCode();
            hash = (hash * 31) + LeftIndentTwips.GetHashCode();
            hash = (hash * 31) + RightIndentTwips.GetHashCode();
            hash = (hash * 31) + FirstLineIndentTwips.GetHashCode();
            return hash;
        }
    }
}
