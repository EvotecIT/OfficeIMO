namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Dependency diagram node kind.
    /// </summary>
    public enum VisioDependencyNodeKind {
        /// <summary>Standard process, service, or component.</summary>
        Component,

        /// <summary>Data store or stateful dependency.</summary>
        Data,

        /// <summary>External actor, client, or system.</summary>
        External,

        /// <summary>Decision, gateway, or policy dependency.</summary>
        Decision
    }
}
