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
        /// Force single-threaded execution with no locking.
        /// </summary>
        Sequential,

        /// <summary>
        /// Compute work in parallel and apply results sequentially.
        /// </summary>
        Parallel
    }
}

