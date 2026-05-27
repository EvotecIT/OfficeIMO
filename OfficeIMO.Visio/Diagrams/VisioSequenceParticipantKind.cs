namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Semantic participant kinds used by <see cref="VisioSequenceDiagramBuilder"/>.
    /// </summary>
    public enum VisioSequenceParticipantKind {
        /// <summary>Generic system, service, or component participant.</summary>
        Participant,

        /// <summary>External person, user, or actor.</summary>
        Actor,

        /// <summary>Boundary or edge/interface component.</summary>
        Boundary,

        /// <summary>Control or coordinator component.</summary>
        Control,

        /// <summary>Domain entity or stateful object.</summary>
        Entity,

        /// <summary>Database or durable data store.</summary>
        Database
    }
}
