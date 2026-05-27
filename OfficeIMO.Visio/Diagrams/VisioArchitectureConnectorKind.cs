namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Semantic connector kinds used by the architecture diagram builder.
    /// </summary>
    public enum VisioArchitectureConnectorKind {
        /// <summary>Primary data or request flow.</summary>
        Data,

        /// <summary>Control, management, or orchestration flow.</summary>
        Control,

        /// <summary>Dependency or supporting relationship.</summary>
        Dependency
    }
}
