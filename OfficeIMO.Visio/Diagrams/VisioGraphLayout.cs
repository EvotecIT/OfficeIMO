namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Automatic layout strategy for generic graph diagrams.
    /// </summary>
    public enum VisioGraphLayout {
        /// <summary>Place nodes in breadth-first layers from roots or inferred roots. Cycles are tolerated.</summary>
        Layered,

        /// <summary>Place nodes in a compact grid by insertion order.</summary>
        Grid,

        /// <summary>Place nodes on rings around root nodes.</summary>
        Radial
    }
}
