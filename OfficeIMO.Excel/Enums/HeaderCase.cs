namespace OfficeIMO.Excel
{
    /// <summary>
    /// Specifies the casing applied to generated headers.
    /// </summary>
    public enum HeaderCase
    {
        /// <summary>
        /// Use the original casing of the property names.
        /// </summary>
        Raw,

        /// <summary>
        /// Convert headers to PascalCase.
        /// </summary>
        Pascal,

        /// <summary>
        /// Convert headers to title case with spaces.
        /// </summary>
        Title
    }
}

