namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Describes how null values are represented when flattening objects.
    /// </summary>
    public enum NullPolicy {
        /// <summary>
        /// Replace nulls with an empty string.
        /// </summary>
        EmptyString,
        /// <summary>
        /// Keep null values as null references.
        /// </summary>
        NullLiteral,
        /// <summary>
        /// Use a default value supplied for the property.
        /// </summary>
        DefaultValue
    }
}
