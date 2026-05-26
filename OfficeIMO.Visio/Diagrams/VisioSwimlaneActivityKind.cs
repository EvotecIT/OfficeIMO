namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Semantic activity kinds used by the swimlane diagram builder.
    /// </summary>
    public enum VisioSwimlaneActivityKind {
        /// <summary>A process step.</summary>
        Step,

        /// <summary>A branching decision.</summary>
        Decision,

        /// <summary>A data/input/output activity.</summary>
        Data,

        /// <summary>A start marker.</summary>
        Start,

        /// <summary>An end marker.</summary>
        End
    }
}
