namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents a PowerPoint 2000-2003 presentation comment.</summary>
    public sealed class LegacyPptComment {
        internal LegacyPptComment(int index, string author, string initials, string text,
            DateTime? createdAtUtc, int x, int y) {
            Index = index;
            Author = author ?? string.Empty;
            Initials = initials ?? string.Empty;
            Text = text ?? string.Empty;
            CreatedAtUtc = createdAtUtc;
            X = x;
            Y = y;
        }

        /// <summary>Gets the non-negative comment label index.</summary>
        public int Index { get; }

        /// <summary>Gets the embedded comment author name.</summary>
        public string Author { get; }

        /// <summary>Gets the embedded author initials.</summary>
        public string Initials { get; }

        /// <summary>Gets the comment text.</summary>
        public string Text { get; }

        /// <summary>Gets the UTC creation time, or null when the source contains an empty SYSTEMTIME.</summary>
        public DateTime? CreatedAtUtc { get; }

        /// <summary>Gets the comment-label x-coordinate in PowerPoint master units.</summary>
        public int X { get; }

        /// <summary>Gets the comment-label y-coordinate in PowerPoint master units.</summary>
        public int Y { get; }
    }
}
