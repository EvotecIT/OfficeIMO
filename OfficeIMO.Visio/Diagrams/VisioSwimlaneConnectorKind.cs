namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Semantic flow kinds used by the swimlane diagram builder.
    /// </summary>
    public enum VisioSwimlaneConnectorKind {
        /// <summary>A normal process flow.</summary>
        Flow,

        /// <summary>A handoff between roles or lanes.</summary>
        Handoff,

        /// <summary>An exception, retry, or alternate path.</summary>
        Exception
    }
}
