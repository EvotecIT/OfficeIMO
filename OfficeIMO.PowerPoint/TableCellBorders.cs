using System;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Flags describing which table cell borders to apply.
    /// </summary>
    [Flags]
    public enum TableCellBorders {
        /// <summary>
        /// No borders.
        /// </summary>
        None = 0,
        /// <summary>
        /// Left border.
        /// </summary>
        Left = 1,
        /// <summary>
        /// Top border.
        /// </summary>
        Top = 2,
        /// <summary>
        /// Right border.
        /// </summary>
        Right = 4,
        /// <summary>
        /// Bottom border.
        /// </summary>
        Bottom = 8,
        /// <summary>
        /// All borders.
        /// </summary>
        All = Left | Top | Right | Bottom
    }
}
