namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Semantic message kinds used by <see cref="VisioSequenceDiagramBuilder"/>.
    /// </summary>
    public enum VisioSequenceMessageKind {
        /// <summary>Synchronous request/call message.</summary>
        Call,

        /// <summary>Asynchronous message or event.</summary>
        Async,

        /// <summary>Return or response message.</summary>
        Return,

        /// <summary>Notification/event message.</summary>
        Event
    }
}
