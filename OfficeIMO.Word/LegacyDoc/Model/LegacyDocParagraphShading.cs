namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocParagraphShading : IEquatable<LegacyDocParagraphShading> {
        internal LegacyDocParagraphShading(string? fillColorHex) {
            FillColorHex = string.IsNullOrWhiteSpace(fillColorHex)
                ? null
                : fillColorHex!.Replace("#", string.Empty).ToUpperInvariant();
        }

        internal string? FillColorHex { get; }

        internal bool HasAny => !string.IsNullOrEmpty(FillColorHex);

        public bool Equals(LegacyDocParagraphShading other) {
            return string.Equals(FillColorHex, other.FillColorHex, StringComparison.Ordinal);
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocParagraphShading other && Equals(other);
        }

        public override int GetHashCode() {
            return FillColorHex == null ? 0 : FillColorHex.GetHashCode();
        }
    }
}
