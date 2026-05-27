namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Built-in visual category for generic graph edges.
    /// </summary>
    public enum VisioGraphConnectorKind {
        /// <summary>Standard relationship edge.</summary>
        Standard,

        /// <summary>Data-flow or data-dependency edge.</summary>
        Data,

        /// <summary>Control-flow, policy, or management edge.</summary>
        Control,

        /// <summary>Highlighted or critical-path edge.</summary>
        Emphasis
    }
}
