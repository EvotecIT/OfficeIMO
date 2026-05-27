namespace OfficeIMO.Visio {
    /// <summary>
    /// Specifies which connectors receive line jumps on a Visio page.
    /// </summary>
    public enum VisioLineJumpCode {
        /// <summary>No connectors receive line jumps.</summary>
        None = 0,

        /// <summary>Horizontal connectors receive line jumps.</summary>
        Horizontal = 1,

        /// <summary>Vertical connectors receive line jumps.</summary>
        Vertical = 2,

        /// <summary>The last routed connector receives the line jump.</summary>
        LastRouted = 3,

        /// <summary>The last displayed connector receives the line jump.</summary>
        DisplayOrder = 4,

        /// <summary>The first displayed connector receives the line jump.</summary>
        ReverseDisplayOrder = 5
    }
}
