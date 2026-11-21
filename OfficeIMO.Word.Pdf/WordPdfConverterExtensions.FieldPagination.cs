namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        /// <summary>Returns independent field-update options using the native PDF layout for bookmark and page-count fields.</summary>
        /// <remarks>
        /// The layout is the one <see cref="SaveAsPdf(WordDocument, string, WordPdfExportOptions?, CancellationToken)"/> produces
        /// with default options. Use <see cref="WithPdfLayoutPagination(WordFieldUpdateOptions, WordPdfExportOptions)"/> to lay
        /// the document out with the fonts and resource policy the application exports with.
        /// The Word package continues to work without PDF. This explicit selection retains a static, AOT-safe reference to the rendering owner.
        /// </remarks>
        /// <param name="options">Field refresh settings to copy; the caller's instance is unchanged.</param>
        /// <returns>Options that paginate a private document copy after any requested table-of-contents regeneration.</returns>
        public static WordFieldUpdateOptions WithPdfLayoutPagination(this WordFieldUpdateOptions options) =>
            WithPdfLayoutPagination(options, new WordPdfExportOptions());

        /// <summary>
        /// Returns independent field-update options whose page numbers come from the same layout as PDF export with
        /// <paramref name="layoutOptions"/>.
        /// </summary>
        /// <remarks>
        /// Field pagination lays the document out exactly as
        /// <see cref="SaveAsPdf(WordDocument, string, WordPdfExportOptions?, CancellationToken)"/> does with the same options:
        /// the same font registrations and substitutions in <see cref="WordPdfExportOptions.PdfOptions"/>, the same
        /// <see cref="WordPdfExportOptions.FontFamily"/>, <see cref="WordPdfExportOptions.TextFallbacks"/>, page settings and
        /// <see cref="WordPdfExportOptions.ResourcePolicy"/>. A policy that forbids system fonts, such as
        /// <see cref="OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic"/>, therefore gives the same page numbers on every host.
        /// The options are copied when this method is called; later changes to <paramref name="layoutOptions"/> do not affect the result.
        /// Loss reporting does not apply: <see cref="WordPdfExportOptions.FailOnLoss"/> never stops a field update, and
        /// <see cref="WordPdfExportOptions.DateTimeFieldUpdateOptions"/> is ignored because the field update owns DATE and TIME fields.
        /// </remarks>
        /// <param name="options">Field refresh settings to copy; the caller's instance is unchanged.</param>
        /// <param name="layoutOptions">PDF export options whose layout supplies the page numbers.</param>
        /// <returns>Options that paginate a private document copy after any requested table-of-contents regeneration.</returns>
        public static WordFieldUpdateOptions WithPdfLayoutPagination(this WordFieldUpdateOptions options, WordPdfExportOptions layoutOptions) {
            ArgumentNullException.ThrowIfNull(options);
            ArgumentNullException.ThrowIfNull(layoutOptions);
            WordPdfExportOptions layout = layoutOptions.Clone();
            return new WordFieldUpdateOptions {
                CurrentDateTime = options.CurrentDateTime,
                DateTimeFieldsOnly = options.DateTimeFieldsOnly,
                UpdateTablesOfContents = options.UpdateTablesOfContents,
                TablesOfContentsOnly = options.TablesOfContentsOnly,
                UseLayoutPagination = true,
                LayoutPaginationProvider = document => WordPdfLayoutPagination.Create(document, layout)
            };
        }
    }
}
