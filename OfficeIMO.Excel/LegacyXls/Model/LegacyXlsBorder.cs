namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents parsed legacy XLS border formatting from an XF record.
    /// </summary>
    public sealed class LegacyXlsBorder {
        /// <summary>
        /// Creates parsed legacy XLS border formatting.
        /// </summary>
        public LegacyXlsBorder(
            byte leftStyle,
            byte rightStyle,
            byte topStyle,
            byte bottomStyle,
            ushort leftColorIndex,
            ushort rightColorIndex,
            ushort topColorIndex,
            ushort bottomColorIndex,
            byte diagonalStyle,
            ushort diagonalColorIndex,
            bool diagonalUp,
            bool diagonalDown) {
            LeftStyle = leftStyle;
            RightStyle = rightStyle;
            TopStyle = topStyle;
            BottomStyle = bottomStyle;
            LeftColorIndex = leftColorIndex;
            RightColorIndex = rightColorIndex;
            TopColorIndex = topColorIndex;
            BottomColorIndex = bottomColorIndex;
            DiagonalStyle = diagonalStyle;
            DiagonalColorIndex = diagonalColorIndex;
            DiagonalUp = diagonalUp;
            DiagonalDown = diagonalDown;
        }

        /// <summary>Gets the legacy left border style.</summary>
        public byte LeftStyle { get; }

        /// <summary>Gets the legacy right border style.</summary>
        public byte RightStyle { get; }

        /// <summary>Gets the legacy top border style.</summary>
        public byte TopStyle { get; }

        /// <summary>Gets the legacy bottom border style.</summary>
        public byte BottomStyle { get; }

        /// <summary>Gets the legacy left border color index.</summary>
        public ushort LeftColorIndex { get; }

        /// <summary>Gets the legacy right border color index.</summary>
        public ushort RightColorIndex { get; }

        /// <summary>Gets the legacy top border color index.</summary>
        public ushort TopColorIndex { get; }

        /// <summary>Gets the legacy bottom border color index.</summary>
        public ushort BottomColorIndex { get; }

        /// <summary>Gets the legacy diagonal border style.</summary>
        public byte DiagonalStyle { get; }

        /// <summary>Gets the legacy diagonal border color index.</summary>
        public ushort DiagonalColorIndex { get; }

        /// <summary>Gets whether the bottom-left to top-right diagonal border is present.</summary>
        public bool DiagonalUp { get; }

        /// <summary>Gets whether the top-left to bottom-right diagonal border is present.</summary>
        public bool DiagonalDown { get; }
    }
}
