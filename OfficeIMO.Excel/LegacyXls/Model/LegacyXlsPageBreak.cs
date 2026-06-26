namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents an explicit legacy XLS manual page break and its affected row or column span.
    /// </summary>
    public sealed class LegacyXlsPageBreak {
        /// <summary>
        /// Creates a parsed manual page break.
        /// </summary>
        /// <param name="position">One-based row or column after which the page break is inserted.</param>
        /// <param name="start">One-based start row or column affected by the break.</param>
        /// <param name="end">One-based end row or column affected by the break.</param>
        public LegacyXlsPageBreak(int position, int start, int end) {
            Position = position;
            Start = start;
            End = end;
        }

        /// <summary>Gets the one-based row or column after which the page break is inserted.</summary>
        public int Position { get; }

        /// <summary>Gets the one-based start row or column affected by the break.</summary>
        public int Start { get; }

        /// <summary>Gets the one-based end row or column affected by the break.</summary>
        public int End { get; }
    }
}
