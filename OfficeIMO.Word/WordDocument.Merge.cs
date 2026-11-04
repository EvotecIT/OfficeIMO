using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    // Supports merging multiple documents.
    public partial class WordDocument {
        /// <summary>
        /// Appends the content of another <see cref="WordDocument"/> to this
        /// document, keeping destination styles and the source page setup.
        /// </summary>
        /// <param name="source">The document to append.</param>
        public void AppendDocument(WordDocument source) {
            AppendDocument(source, null);
        }

        /// <summary>
        /// Appends the content of another <see cref="WordDocument"/> to this document.
        /// </summary>
        /// <remarks>
        /// Styles, numbering, images, hyperlinks, headers, footers, footnotes, endnotes and comments used by the
        /// appended content are imported. Numbering, note, comment, bookmark, drawing and content-control ids are
        /// renumbered so they do not collide with the destination, and bookmark names that already exist in the
        /// destination are renamed together with the hyperlinks and fields that target them.
        /// </remarks>
        /// <param name="source">The document to append.</param>
        /// <param name="options">Controls style and section handling. When null, defaults are used.</param>
        public void AppendDocument(WordDocument source, WordDocumentMergeOptions? options) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (ReferenceEquals(source, this)) throw new ArgumentException("A document cannot be appended to itself.", nameof(source));

            var sourceMain = source._wordprocessingDocument.MainDocumentPart;
            var destinationMain = _wordprocessingDocument.MainDocumentPart;
            Body? sourceBody = sourceMain?.Document?.Body;
            Body? destinationBody = destinationMain?.Document?.Body;
            if (sourceMain == null || destinationMain == null || sourceBody == null || destinationBody == null) return;

            new WordDocumentMerge(this, sourceMain, destinationMain, options ?? new WordDocumentMergeOptions())
                .Append(sourceBody, destinationBody);

            LoadSections();
            WordListStyles.InitializeAbstractNumberId(_wordprocessingDocument);
        }
    }
}
