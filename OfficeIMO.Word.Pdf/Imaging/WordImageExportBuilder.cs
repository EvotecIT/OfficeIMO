using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Fluent image-export builder for a Word document page.
    /// </summary>
    public sealed class WordDocumentImageExportBuilder : OfficeImageExportBuilder<WordDocumentImageExportBuilder, WordImageExportOptions> {
        internal WordDocumentImageExportBuilder(WordDocument document, WordImageExportOptions? options = null)
            : base(
                options?.Clone() ?? new WordImageExportOptions(),
                (format, effective, cancellationToken) => document.ExportImage(format, effective, cancellationToken)) {
        }

        /// <summary>Exports the first page.</summary>
        public WordDocumentImageExportBuilder FirstPage() => Page(0);

        /// <summary>Sets the zero-based page index to export.</summary>
        public WordDocumentImageExportBuilder Page(int pageIndex) {
            if (pageIndex < 0) {
                throw new System.ArgumentOutOfRangeException(nameof(pageIndex), "Page index cannot be negative.");
            }

            Options.PageIndex = pageIndex;
            return this;
        }
    }

    /// <summary>Fluent batch image export for Word document pages.</summary>
    public sealed class WordDocumentPageImageExportBuilder : OfficeImageExportBatchBuilder<WordDocumentPageImageExportBuilder, WordImageExportOptions> {
        internal WordDocumentPageImageExportBuilder(WordDocument document, WordImageExportOptions? options = null)
            : base(
                options?.Clone() ?? new WordImageExportOptions(),
                document.ExportImages,
                (format, effective, consumer, cancellationToken) =>
                    document.ExportImages(format, consumer, effective, cancellationToken)) {
        }

        /// <summary>Exports from the specified zero-based page index.</summary>
        public WordDocumentPageImageExportBuilder FromPage(int pageIndex) {
            if (pageIndex < 0) throw new System.ArgumentOutOfRangeException(nameof(pageIndex));
            Options.PageIndex = pageIndex;
            return this;
        }

        /// <summary>Limits batch output to the requested number of pages.</summary>
        public WordDocumentPageImageExportBuilder TakePages(int pageCount) {
            if (pageCount < 1) throw new System.ArgumentOutOfRangeException(nameof(pageCount));
            Options.PageCount = pageCount;
            return this;
        }

        /// <summary>Exports every page from the beginning of the document.</summary>
        public WordDocumentPageImageExportBuilder AllPages() {
            Options.PageIndex = 0;
            Options.PageCount = null;
            return this;
        }
    }

    public static partial class WordImageExportExtensions {
        /// <summary>
        /// Starts a fluent image export for this document.
        /// </summary>
        public static WordDocumentImageExportBuilder ToImage(this WordDocument document) => new WordDocumentImageExportBuilder(document);

        /// <summary>Starts a fluent image export using a cloned options snapshot.</summary>
        public static WordDocumentImageExportBuilder ToImage(this WordDocument document, WordImageExportOptions options) =>
            new WordDocumentImageExportBuilder(document, options ?? throw new ArgumentNullException(nameof(options)));

        /// <summary>Starts a fluent batch image export for document pages.</summary>
        public static WordDocumentPageImageExportBuilder ToImages(this WordDocument document) => new WordDocumentPageImageExportBuilder(document);

        /// <summary>Starts a fluent page batch export using a cloned options snapshot.</summary>
        public static WordDocumentPageImageExportBuilder ToImages(this WordDocument document, WordImageExportOptions options) =>
            new WordDocumentPageImageExportBuilder(document, options ?? throw new ArgumentNullException(nameof(options)));
    }
}
