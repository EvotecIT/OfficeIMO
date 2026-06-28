namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a BIFF8 protected range entry from a worksheet shared-feature record.
    /// </summary>
    public sealed class LegacyXlsProtectedRange {
        /// <summary>
        /// Creates protected range metadata.
        /// </summary>
        /// <param name="name">Protected range title.</param>
        /// <param name="references">A1 references covered by the protected range.</param>
        /// <param name="legacyPasswordHash">Optional 16-bit legacy password verifier formatted as four uppercase hexadecimal digits.</param>
        /// <param name="hasSecurityDescriptor">Whether the source record carried Windows security descriptor metadata.</param>
        public LegacyXlsProtectedRange(
            string name,
            IReadOnlyList<string> references,
            string? legacyPasswordHash = null,
            bool hasSecurityDescriptor = false) {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            References = references ?? throw new ArgumentNullException(nameof(references));
            LegacyPasswordHash = legacyPasswordHash;
            HasSecurityDescriptor = hasSecurityDescriptor;
        }

        /// <summary>
        /// Gets the protected range title.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets A1 references covered by the protected range.
        /// </summary>
        public IReadOnlyList<string> References { get; }

        /// <summary>
        /// Gets the optional legacy password verifier.
        /// </summary>
        public string? LegacyPasswordHash { get; }

        /// <summary>
        /// Gets whether the source record carried Windows security descriptor metadata.
        /// </summary>
        public bool HasSecurityDescriptor { get; }
    }
}
