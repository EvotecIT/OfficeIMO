namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocHyperlinkTarget : IEquatable<LegacyDocHyperlinkTarget> {
        private LegacyDocHyperlinkTarget(string? uri, string? anchor) {
            Uri = string.IsNullOrWhiteSpace(uri) ? null : uri;
            Anchor = string.IsNullOrWhiteSpace(anchor) ? null : anchor;
        }

        internal static LegacyDocHyperlinkTarget ForUri(string uri) {
            return new LegacyDocHyperlinkTarget(uri, null);
        }

        internal static LegacyDocHyperlinkTarget ForAnchor(string anchor) {
            return new LegacyDocHyperlinkTarget(null, anchor);
        }

        internal string? Uri { get; }

        internal string? Anchor { get; }

        internal bool HasValue => Uri != null || Anchor != null;

        public bool Equals(LegacyDocHyperlinkTarget other) {
            return string.Equals(Uri, other.Uri, StringComparison.Ordinal)
                && string.Equals(Anchor, other.Anchor, StringComparison.Ordinal);
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocHyperlinkTarget other && Equals(other);
        }

        public override int GetHashCode() {
            unchecked {
                int hash = 17;
                hash = (hash * 31) + StringComparer.Ordinal.GetHashCode(Uri ?? string.Empty);
                hash = (hash * 31) + StringComparer.Ordinal.GetHashCode(Anchor ?? string.Empty);
                return hash;
            }
        }

        public static bool operator ==(LegacyDocHyperlinkTarget left, LegacyDocHyperlinkTarget right) {
            return left.Equals(right);
        }

        public static bool operator !=(LegacyDocHyperlinkTarget left, LegacyDocHyperlinkTarget right) {
            return !left.Equals(right);
        }
    }
}
