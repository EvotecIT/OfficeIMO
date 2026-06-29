namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocCharacterFormat : IEquatable<LegacyDocCharacterFormat> {
        internal LegacyDocCharacterFormat(
            bool bold,
            bool italic,
            bool strike,
            bool doubleStrike,
            LegacyDocCapsKind? caps,
            LegacyDocVerticalPositionKind? verticalPosition,
            LegacyDocUnderlineKind? underline,
            LegacyDocHighlightColorKind? highlight,
            int? fontSizeHalfPoints,
            string? colorHex,
            string? fontFamily) {
            Bold = bold;
            Italic = italic;
            Strike = strike;
            DoubleStrike = doubleStrike;
            Caps = caps;
            VerticalPosition = verticalPosition;
            Underline = underline;
            Highlight = highlight;
            FontSizeHalfPoints = fontSizeHalfPoints;
            ColorHex = colorHex;
            FontFamily = fontFamily;
        }

        internal bool Bold { get; }

        internal bool Italic { get; }

        internal bool Strike { get; }

        internal bool DoubleStrike { get; }

        internal LegacyDocCapsKind? Caps { get; }

        internal LegacyDocVerticalPositionKind? VerticalPosition { get; }

        internal LegacyDocUnderlineKind? Underline { get; }

        internal LegacyDocHighlightColorKind? Highlight { get; }

        internal int? FontSizeHalfPoints { get; }

        internal string? ColorHex { get; }

        internal string? FontFamily { get; }

        internal bool HasFormatting =>
            Bold
            || Italic
            || Strike
            || DoubleStrike
            || Caps != null
            || VerticalPosition != null
            || Underline != null
            || Highlight != null
            || FontSizeHalfPoints != null
            || ColorHex != null
            || FontFamily != null;

        internal static LegacyDocCharacterFormat Default { get; } = new LegacyDocCharacterFormat(false, false, false, false, null, null, null, null, null, null, null);

        public bool Equals(LegacyDocCharacterFormat other) {
            return Bold == other.Bold
                && Italic == other.Italic
                && Strike == other.Strike
                && DoubleStrike == other.DoubleStrike
                && Caps == other.Caps
                && VerticalPosition == other.VerticalPosition
                && Underline == other.Underline
                && Highlight == other.Highlight
                && FontSizeHalfPoints == other.FontSizeHalfPoints
                && string.Equals(ColorHex, other.ColorHex, StringComparison.OrdinalIgnoreCase)
                && string.Equals(FontFamily, other.FontFamily, StringComparison.OrdinalIgnoreCase);
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocCharacterFormat other && Equals(other);
        }

        public override int GetHashCode() {
            int hash = 17;
            hash = (hash * 31) + Bold.GetHashCode();
            hash = (hash * 31) + Italic.GetHashCode();
            hash = (hash * 31) + Strike.GetHashCode();
            hash = (hash * 31) + DoubleStrike.GetHashCode();
            hash = (hash * 31) + Caps.GetHashCode();
            hash = (hash * 31) + VerticalPosition.GetHashCode();
            hash = (hash * 31) + Underline.GetHashCode();
            hash = (hash * 31) + Highlight.GetHashCode();
            hash = (hash * 31) + FontSizeHalfPoints.GetHashCode();
            hash = (hash * 31) + StringComparer.OrdinalIgnoreCase.GetHashCode(ColorHex ?? string.Empty);
            hash = (hash * 31) + StringComparer.OrdinalIgnoreCase.GetHashCode(FontFamily ?? string.Empty);
            return hash;
        }
    }
}
