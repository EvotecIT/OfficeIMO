namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents a legacy PowerPoint shape rectangle in master units.</summary>
    public readonly struct LegacyPptBounds {
        /// <summary>Creates a shape rectangle.</summary>
        public LegacyPptBounds(int left, int top, int width, int height) {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }

        /// <summary>Gets the left coordinate.</summary>
        public int Left { get; }

        /// <summary>Gets the top coordinate.</summary>
        public int Top { get; }

        /// <summary>Gets the width.</summary>
        public int Width { get; }

        /// <summary>Gets the height.</summary>
        public int Height { get; }
    }
}
