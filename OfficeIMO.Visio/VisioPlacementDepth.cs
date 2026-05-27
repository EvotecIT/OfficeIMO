namespace OfficeIMO.Visio {
    /// <summary>
    /// Specifies how deeply Visio analyzes connected shapes before creating a page layout.
    /// </summary>
    public enum VisioPlacementDepth {
        /// <summary>Use the page or template default.</summary>
        Default = 0,

        /// <summary>Use medium analysis depth.</summary>
        Medium = 1,

        /// <summary>Use deep analysis.</summary>
        Deep = 2,

        /// <summary>Use shallow analysis.</summary>
        Shallow = 3
    }
}
