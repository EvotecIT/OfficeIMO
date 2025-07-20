namespace OfficeIMO.Word {
    /// <summary>
    /// Specifies text matching behavior when applying conditional formatting.
    /// </summary>
    public enum TextMatchType {
        /// <summary>
        /// Text must exactly equal the provided value.
        /// </summary>
        Equals,

        /// <summary>
        /// Text must contain the provided value.
        /// </summary>
        Contains,

        /// <summary>
        /// Text must start with the provided value.
        /// </summary>
        StartsWith,

        /// <summary>
        /// Text must end with the provided value.
        /// </summary>
        EndsWith
    }
}
