namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        /// <summary>Returns independent field-update options using the native PDF layout for bookmark and page-count fields.</summary>
        /// <remarks>The Word package continues to work without PDF. This explicit selection retains a static, AOT-safe reference to the rendering owner.</remarks>
        /// <param name="options">Field refresh settings to copy; the caller's instance is unchanged.</param>
        /// <returns>Options that paginate a private document copy after any requested table-of-contents regeneration.</returns>
        public static WordFieldUpdateOptions WithPdfLayoutPagination(this WordFieldUpdateOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            return new WordFieldUpdateOptions {
                CurrentDateTime = options.CurrentDateTime,
                DateTimeFieldsOnly = options.DateTimeFieldsOnly,
                UpdateTablesOfContents = options.UpdateTablesOfContents,
                UseLayoutPagination = true,
                LayoutPaginationProvider = WordPdfLayoutPagination.Create
            };
        }
    }
}
