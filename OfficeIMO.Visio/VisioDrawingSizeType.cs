namespace OfficeIMO.Visio {
    /// <summary>
    /// Specifies how Visio determines the drawing page size.
    /// </summary>
    public enum VisioDrawingSizeType {
        /// <summary>
        /// Match the printer page size.
        /// </summary>
        SameAsPrinter = 0,

        /// <summary>
        /// Fit the page to the drawing contents.
        /// </summary>
        FitToDrawingContents = 1,

        /// <summary>
        /// Use a standard page size.
        /// </summary>
        Standard = 2,

        /// <summary>
        /// Use a custom page size.
        /// </summary>
        Custom = 3,

        /// <summary>
        /// Use a custom scaled drawing size.
        /// </summary>
        CustomScaled = 4,

        /// <summary>
        /// Use a metric ISO page size.
        /// </summary>
        Metric = 5,

        /// <summary>
        /// Use an ANSI engineering page size.
        /// </summary>
        AnsiEngineering = 6,

        /// <summary>
        /// Use an ANSI architectural page size.
        /// </summary>
        AnsiArchitectural = 7
    }
}
