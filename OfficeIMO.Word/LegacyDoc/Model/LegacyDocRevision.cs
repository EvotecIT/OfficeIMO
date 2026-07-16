namespace OfficeIMO.Word.LegacyDoc.Model {
    internal enum LegacyDocRevisionKind {
        None,
        Inserted,
        Deleted
    }

    internal readonly struct LegacyDocRevision : IEquatable<LegacyDocRevision> {
        internal LegacyDocRevision(LegacyDocRevisionKind kind, string author, DateTime? date) {
            Kind = kind;
            Author = string.IsNullOrWhiteSpace(author) ? "Unknown" : author;
            Date = date;
        }

        internal LegacyDocRevisionKind Kind { get; }

        internal string? Author { get; }

        internal DateTime? Date { get; }

        internal bool HasValue => Kind != LegacyDocRevisionKind.None;

        internal static LegacyDocRevision None { get; } = default;

        public bool Equals(LegacyDocRevision other) {
            return Kind == other.Kind
                && string.Equals(Author, other.Author, StringComparison.Ordinal)
                && Date == other.Date;
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocRevision other && Equals(other);
        }

        public override int GetHashCode() {
            int hash = 17;
            hash = (hash * 31) + Kind.GetHashCode();
            hash = (hash * 31) + StringComparer.Ordinal.GetHashCode(Author ?? string.Empty);
            hash = (hash * 31) + Date.GetHashCode();
            return hash;
        }
    }
}
