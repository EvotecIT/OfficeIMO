using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Fluent image-export builder for a Word document page preview.
    /// </summary>
    public sealed class WordDocumentImageExportBuilder : OfficeImageExportBuilder<WordDocumentImageExportBuilder, WordImageExportOptions> {
        internal WordDocumentImageExportBuilder(WordDocument document, WordImageExportOptions? options = null)
            : base(options?.Clone() ?? new WordImageExportOptions(), document.ExportImage) {
        }

        /// <summary>Exports the first page preview.</summary>
        public WordDocumentImageExportBuilder FirstPage() => Page(0);

        /// <summary>Sets the zero-based page index to export.</summary>
        public WordDocumentImageExportBuilder Page(int pageIndex) {
            if (pageIndex < 0) {
                throw new System.ArgumentOutOfRangeException(nameof(pageIndex), "Page index cannot be negative.");
            }

            Options.PageIndex = pageIndex;
            return this;
        }

        /// <summary>Includes or excludes document body content.</summary>
        public WordDocumentImageExportBuilder IncludeContent(bool include = true) {
            Options.IncludeDocumentContent = include;
            return this;
        }

        /// <summary>Excludes document body content.</summary>
        public WordDocumentImageExportBuilder WithoutContent() => IncludeContent(false);
    }

    /// <summary>Fluent batch image export for Word document pages.</summary>
    public sealed class WordDocumentPageImageExportBuilder : OfficeImageExportBatchBuilder<WordDocumentPageImageExportBuilder, WordImageExportOptions> {
        internal WordDocumentPageImageExportBuilder(WordDocument document, WordImageExportOptions? options = null)
            : base(options?.Clone() ?? new WordImageExportOptions(), document.ExportImages) {
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

        /// <summary>Exports every estimated page from the beginning of the document.</summary>
        public WordDocumentPageImageExportBuilder AllPages() {
            Options.PageIndex = 0;
            Options.PageCount = null;
            return this;
        }

        /// <summary>Includes or excludes document body content.</summary>
        public WordDocumentPageImageExportBuilder IncludeContent(bool include = true) {
            Options.IncludeDocumentContent = include;
            return this;
        }
    }

    public partial class WordDocument {
        /// <summary>
        /// Starts a fluent image export for this document.
        /// </summary>
        public WordDocumentImageExportBuilder ToImage() => new WordDocumentImageExportBuilder(this);

        /// <summary>Starts a fluent image export using a cloned options snapshot.</summary>
        public WordDocumentImageExportBuilder ToImage(WordImageExportOptions options) =>
            new WordDocumentImageExportBuilder(this, options ?? throw new ArgumentNullException(nameof(options)));

        /// <summary>Starts a fluent batch image export for document pages.</summary>
        public WordDocumentPageImageExportBuilder ToImages() => new WordDocumentPageImageExportBuilder(this);

        /// <summary>Starts a fluent page batch export using a cloned options snapshot.</summary>
        public WordDocumentPageImageExportBuilder ToImages(WordImageExportOptions options) =>
            new WordDocumentPageImageExportBuilder(this, options ?? throw new ArgumentNullException(nameof(options)));
    }
}
