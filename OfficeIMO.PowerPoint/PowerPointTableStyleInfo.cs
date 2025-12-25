namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a table style entry available in the presentation.
    /// </summary>
    public readonly struct PowerPointTableStyleInfo {
        /// <summary>
        ///     Creates a table style info entry.
        /// </summary>
        /// <param name="styleId">The Open XML style identifier (GUID).</param>
        /// <param name="name">The display name of the style.</param>
        public PowerPointTableStyleInfo(string styleId, string name) {
            StyleId = styleId;
            Name = name;
        }

        /// <summary>
        ///     Open XML style identifier (GUID).
        /// </summary>
        public string StyleId { get; }

        /// <summary>
        ///     Display name of the style.
        /// </summary>
        public string Name { get; }

        /// <summary>
        ///     Returns the style name or ID.
        /// </summary>
        public override string ToString() {
            return string.IsNullOrWhiteSpace(Name) ? StyleId : Name;
        }
    }
}
