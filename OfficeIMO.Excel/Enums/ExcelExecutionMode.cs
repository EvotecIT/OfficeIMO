namespace OfficeIMO.Excel {
    /// <summary>
    /// Determines how operations are executed.
    /// </summary>
    public enum ExcelExecutionMode {
        /// <summary>
        /// Automatically choose between sequential and parallel execution based on thresholds.
        /// </summary>
        Automatic,

        /// <summary>
        /// Prefer single-threaded compute. Specialized single-pass readers may be used when they are faster.
        /// </summary>
        Sequential,

        /// <summary>
        /// Permit eligible compute work to run in parallel and apply results in source order. A specialized
        /// single-pass reader may still be selected when parallel staging would only add overhead.
        /// </summary>
        Parallel
    }
}

