namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Built-in visual category for generic graph nodes.
    /// </summary>
    public enum VisioGraphNodeKind {
        /// <summary>Standard component or process node.</summary>
        Process,

        /// <summary>Decision, gateway, or policy node.</summary>
        Decision,

        /// <summary>Data store or persisted data node.</summary>
        Data,

        /// <summary>External actor or external system node.</summary>
        External,

        /// <summary>Emphasized infrastructure or boundary node.</summary>
        Emphasis
    }
}
