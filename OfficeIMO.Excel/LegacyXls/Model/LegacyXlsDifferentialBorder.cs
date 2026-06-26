namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents border formatting decoded from a legacy XLS differential-format record.
    /// </summary>
    public sealed class LegacyXlsDifferentialBorder {
        /// <summary>
        /// Creates a decoded differential border.
        /// </summary>
        public LegacyXlsDifferentialBorder(
            LegacyXlsDifferentialBorderSide? top,
            LegacyXlsDifferentialBorderSide? bottom,
            LegacyXlsDifferentialBorderSide? left,
            LegacyXlsDifferentialBorderSide? right) {
            Top = top;
            Bottom = bottom;
            Left = left;
            Right = right;
        }

        /// <summary>Gets the top border formatting, when present.</summary>
        public LegacyXlsDifferentialBorderSide? Top { get; }

        /// <summary>Gets the bottom border formatting, when present.</summary>
        public LegacyXlsDifferentialBorderSide? Bottom { get; }

        /// <summary>Gets the left border formatting, when present.</summary>
        public LegacyXlsDifferentialBorderSide? Left { get; }

        /// <summary>Gets the right border formatting, when present.</summary>
        public LegacyXlsDifferentialBorderSide? Right { get; }

        /// <summary>Gets whether at least one side contains decoded border formatting.</summary>
        public bool HasAnySide => Top != null || Bottom != null || Left != null || Right != null;
    }

    /// <summary>
    /// Represents one side of a legacy XLS differential border.
    /// </summary>
    public sealed class LegacyXlsDifferentialBorderSide {
        /// <summary>
        /// Creates a decoded differential border side.
        /// </summary>
        public LegacyXlsDifferentialBorderSide(ushort style, string? color) {
            Style = style;
            Color = color;
        }

        /// <summary>Gets the legacy border style value.</summary>
        public ushort Style { get; }

        /// <summary>Gets the border color as ARGB hex, when present.</summary>
        public string? Color { get; }

        /// <summary>Gets whether the side carries a visible border style.</summary>
        public bool HasStyle => Style != 0;
    }
}
