namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Semantic flowchart node kinds used by <see cref="VisioFlowchartBuilder"/>.
    /// </summary>
    public enum VisioFlowchartNodeKind {
        /// <summary>Start terminator.</summary>
        Start,

        /// <summary>Standard process step.</summary>
        Process,

        /// <summary>Decision/branch node.</summary>
        Decision,

        /// <summary>Input/output data node.</summary>
        Data,

        /// <summary>Off-page reference marker.</summary>
        OffPageReference,

        /// <summary>Continuation marker for another column or page region.</summary>
        Continuation,

        /// <summary>End terminator.</summary>
        End
    }
}
