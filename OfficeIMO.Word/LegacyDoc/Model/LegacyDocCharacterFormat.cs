namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocCharacterFormat : IEquatable<LegacyDocCharacterFormat> {
        internal LegacyDocCharacterFormat(
            bool bold,
            bool italic,
            LegacyDocUnderlineKind? underline,
            int? fontSizeHalfPoints,
            string? colorHex) {
            Bold = bold;
            Italic = italic;
            Underline = underline;
            FontSizeHalfPoints = fontSizeHalfPoints;
            ColorHex = colorHex;
        }

        internal bool Bold { get; }

        internal bool Italic { get; }

        internal LegacyDocUnderlineKind? Underline { get; }

        internal int? FontSizeHalfPoints { get; }

        internal string? ColorHex { get; }

        internal bool HasFormatting =>
            Bold
            || Italic
            || Underline != null
            || FontSizeHalfPoints != null
            || ColorHex != null;

        internal static LegacyDocCharacterFormat Default { get; } = new LegacyDocCharacterFormat(false, false, null, null, null);

        public bool Equals(LegacyDocCharacterFormat other) {
            return Bold == other.Bold
                && Italic == other.Italic
                && Underline == other.Underline
                && FontSizeHalfPoints == other.FontSizeHalfPoints
                && string.Equals(ColorHex, other.ColorHex, StringComparison.OrdinalIgnoreCase);
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocCharacterFormat other && Equals(other);
        }

        public override int GetHashCode() {
            int hash = 17;
            hash = (hash * 31) + Bold.GetHashCode();
            hash = (hash * 31) + Italic.GetHashCode();
            hash = (hash * 31) + Underline.GetHashCode();
            hash = (hash * 31) + FontSizeHalfPoints.GetHashCode();
            hash = (hash * 31) + StringComparer.OrdinalIgnoreCase.GetHashCode(ColorHex ?? string.Empty);
            return hash;
        }
    }
}
