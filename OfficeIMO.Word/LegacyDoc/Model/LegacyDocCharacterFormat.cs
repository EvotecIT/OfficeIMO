namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocCharacterFormat : IEquatable<LegacyDocCharacterFormat> {
        internal LegacyDocCharacterFormat(
            bool bold,
            bool italic,
            bool strike,
            LegacyDocUnderlineKind? underline,
            int? fontSizeHalfPoints,
            string? colorHex,
            string? fontFamily) {
            Bold = bold;
            Italic = italic;
            Strike = strike;
            Underline = underline;
            FontSizeHalfPoints = fontSizeHalfPoints;
            ColorHex = colorHex;
            FontFamily = fontFamily;
        }

        internal bool Bold { get; }

        internal bool Italic { get; }

        internal bool Strike { get; }

        internal LegacyDocUnderlineKind? Underline { get; }

        internal int? FontSizeHalfPoints { get; }

        internal string? ColorHex { get; }

        internal string? FontFamily { get; }

        internal bool HasFormatting =>
            Bold
            || Italic
            || Strike
            || Underline != null
            || FontSizeHalfPoints != null
            || ColorHex != null
            || FontFamily != null;

        internal static LegacyDocCharacterFormat Default { get; } = new LegacyDocCharacterFormat(false, false, false, null, null, null, null);

        public bool Equals(LegacyDocCharacterFormat other) {
            return Bold == other.Bold
                && Italic == other.Italic
                && Strike == other.Strike
                && Underline == other.Underline
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
            hash = (hash * 31) + Underline.GetHashCode();
            hash = (hash * 31) + FontSizeHalfPoints.GetHashCode();
            hash = (hash * 31) + StringComparer.OrdinalIgnoreCase.GetHashCode(ColorHex ?? string.Empty);
            hash = (hash * 31) + StringComparer.OrdinalIgnoreCase.GetHashCode(FontFamily ?? string.Empty);
            return hash;
        }
    }
}
