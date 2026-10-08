namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocHyperlinkTarget : IEquatable<LegacyDocHyperlinkTarget> {
        private LegacyDocHyperlinkTarget(string? uri, string? anchor, string? tooltip, string? targetFrame) {
            Uri = string.IsNullOrWhiteSpace(uri) ? null : uri;
            Anchor = string.IsNullOrWhiteSpace(anchor) ? null : anchor;
            Tooltip = tooltip;
            TargetFrame = string.IsNullOrEmpty(targetFrame) ? null : targetFrame;
        }

        internal static LegacyDocHyperlinkTarget ForUri(string uri, string? tooltip = null, string? targetFrame = null) {
            return new LegacyDocHyperlinkTarget(uri, null, tooltip, targetFrame);
        }

        internal static LegacyDocHyperlinkTarget ForAnchor(string anchor, string? tooltip = null, string? targetFrame = null) {
            return new LegacyDocHyperlinkTarget(null, anchor, tooltip, targetFrame);
        }

        internal string? Uri { get; }

        internal string? Anchor { get; }
        internal string? Tooltip { get; }
        internal string? TargetFrame { get; }

        internal bool HasValue => Uri != null || Anchor != null;

        public bool Equals(LegacyDocHyperlinkTarget other) {
            return string.Equals(Uri, other.Uri, StringComparison.Ordinal)
                && string.Equals(Anchor, other.Anchor, StringComparison.Ordinal)
                && string.Equals(Tooltip, other.Tooltip, StringComparison.Ordinal)
                && string.Equals(TargetFrame, other.TargetFrame, StringComparison.Ordinal);
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocHyperlinkTarget other && Equals(other);
        }

        public override int GetHashCode() {
            unchecked {
                int hash = 17;
                hash = (hash * 31) + StringComparer.Ordinal.GetHashCode(Uri ?? string.Empty);
                hash = (hash * 31) + StringComparer.Ordinal.GetHashCode(Anchor ?? string.Empty);
                hash = (hash * 31) + StringComparer.Ordinal.GetHashCode(Tooltip ?? string.Empty);
                hash = (hash * 31) + StringComparer.Ordinal.GetHashCode(TargetFrame ?? string.Empty);
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
