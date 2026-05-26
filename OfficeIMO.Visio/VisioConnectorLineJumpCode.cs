namespace OfficeIMO.Visio {
    /// <summary>
    /// Specifies when a dynamic connector receives line jumps.
    /// </summary>
    public enum VisioConnectorLineJumpCode {
        /// <summary>Use the page-level line jump setting.</summary>
        PageDefault = 0,

        /// <summary>Never display line jumps on this connector.</summary>
        Never = 1,

        /// <summary>Always display line jumps on this connector.</summary>
        Always = 2,

        /// <summary>The other connector receives the line jump.</summary>
        OtherConnectorJumps = 3,

        /// <summary>Neither connector receives the line jump.</summary>
        NeitherConnectorJumps = 4
    }
}
