namespace OfficeIMO.Excel {
    /// <summary>
    /// Controls how object values are projected into cells when building a <see cref="System.Data.DataTable"/>.
    /// </summary>
    public sealed class ObjectDataTableBuilderOptions {
        /// <summary>
        /// Gets or sets whether non-string enumerable property values are joined into one display string.
        /// </summary>
        public bool NormalizeCollectionValues { get; set; } = true;

        /// <summary>
        /// Gets or sets the separator used when <see cref="NormalizeCollectionValues"/> joins collection values.
        /// </summary>
        public string CollectionSeparator { get; set; } = ", ";
    }
}
