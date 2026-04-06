namespace OfficeIMO.Excel {
    /// <summary>
    /// Declares explicit spreadsheet header aliases for typed read mapping.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public sealed class ExcelColumnAttribute : Attribute {
        /// <summary>
        /// Primary header name to map to the decorated property.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Additional header aliases to map to the decorated property.
        /// </summary>
        public IReadOnlyList<string> Aliases { get; }

        /// <summary>
        /// Creates a new explicit header mapping definition.
        /// </summary>
        public ExcelColumnAttribute(string name, params string[] aliases) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Header name cannot be null or whitespace.", nameof(name));
            }

            Name = name;
            Aliases = aliases?.Where(alias => !string.IsNullOrWhiteSpace(alias)).ToArray() ?? Array.Empty<string>();
        }
    }
}
