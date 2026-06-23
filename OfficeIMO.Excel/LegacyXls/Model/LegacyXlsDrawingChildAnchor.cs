namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes shallow metadata decoded from an OfficeArtChildAnchor record.
    /// </summary>
    public sealed class LegacyXlsDrawingChildAnchor {
        /// <summary>
        /// Creates preserve-only metadata for an OfficeArt child anchor.
        /// </summary>
        public LegacyXlsDrawingChildAnchor(int left, int top, int right, int bottom) {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        /// <summary>Gets the left coordinate in the parent drawing coordinate space.</summary>
        public int Left { get; }

        /// <summary>Gets the top coordinate in the parent drawing coordinate space.</summary>
        public int Top { get; }

        /// <summary>Gets the right coordinate in the parent drawing coordinate space.</summary>
        public int Right { get; }

        /// <summary>Gets the bottom coordinate in the parent drawing coordinate space.</summary>
        public int Bottom { get; }

        /// <summary>Gets the decoded width.</summary>
        public int Width => Right - Left;

        /// <summary>Gets the decoded height.</summary>
        public int Height => Bottom - Top;
    }
}
