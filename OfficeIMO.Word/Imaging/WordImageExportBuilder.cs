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

    public partial class WordDocument {
        /// <summary>
        /// Starts a fluent image export for this document.
        /// </summary>
        public WordDocumentImageExportBuilder ToImage() => new WordDocumentImageExportBuilder(this);
    }
}
