namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Built-in deterministic layouts for high-level flowchart authoring.
    /// </summary>
    public enum VisioFlowchartLayout {
        /// <summary>Places nodes in one vertical column.</summary>
        Vertical,

        /// <summary>
        /// Places nodes before the first continuation marker in the left column
        /// and nodes after it in the right column.
        /// </summary>
        TwoColumnContinuation
    }
}
