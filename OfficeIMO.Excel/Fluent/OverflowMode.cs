namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Behavior when a rendered table exceeds the fixed grid column width in SheetComposer.Columns(...).
    /// </summary>
    public enum OverflowMode {
        /// <summary>Throw an exception (default): fail fast to avoid silent data loss.</summary>
        Throw = 0,
        /// <summary>Render only the first N columns that fit; omit the rest.</summary>
        Shrink = 1,
        /// <summary>Render the first N-1 columns and a final "More" column that summarizes omitted columns.</summary>
        Summarize = 2
    }
}

