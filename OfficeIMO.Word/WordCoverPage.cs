using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Built-in cover page templates available for Word documents.
    /// </summary>
    public enum WordCoverPageTemplate {
        /// <summary>
        /// The "Austin" built-in template.
        /// </summary>
        Austin,
        /// <summary>
        /// The "Banded" built-in template.
        /// </summary>
        Banded,
        /// <summary>
        /// The "Facet" built-in template.
        /// </summary>
        Facet,
        /// <summary>
        /// The "Grid" built-in template.
        /// </summary>
        Grid,
        /// <summary>
        /// The "Ion (Dark)" built-in template.
        /// </summary>
        IonDark,
        /// <summary>
        /// The "Ion (Light)" built-in template.
        /// </summary>
        IonLight,
        /// <summary>
        /// The "Element" built-in template.
        /// </summary>
        Element,
        /// <summary>
        /// The "Wisp" built-in template.
        /// </summary>
        Wisp,
        /// <summary>
        /// The "View Master" built-in template.
        /// </summary>
        ViewMaster,
        /// <summary>
        /// The "Slice (Light)" built-in template.
        /// </summary>
        SliceLight,
        /// <summary>
        /// The "Slice (Dark)" built-in template.
        /// </summary>
        SliceDark,
        /// <summary>
        /// The "Sideline" built-in template.
        /// </summary>
        SideLine,
        /// <summary>
        /// The "Semaphore" built-in template.
        /// </summary>
        Semaphore,
        /// <summary>
        /// The "Retrospect" built-in template.
        /// </summary>
        Retrospect
    }

    /// <summary>
    /// Represents a cover page within a Word document.
    /// </summary>
    public partial class WordCoverPage : WordElement {
        private readonly WordDocument _document;
        private readonly SdtBlock _sdtBlock;

        /// <summary>
        /// Gets the underlying structured document tag block for internal converters.
        /// </summary>
        internal SdtBlock SdtBlock => _sdtBlock;

        /// <summary>
        /// Gets the parent document for internal converters.
        /// </summary>
        internal WordDocument Document => _document;

        /// <summary>
        /// Initializes a new instance from an existing structured document tag block.
        /// </summary>
        /// <param name="wordDocument">Parent document.</param>
        /// <param name="sdtBlock">Structured document tag to wrap.</param>
        internal WordCoverPage(WordDocument wordDocument, SdtBlock sdtBlock) {
            _document = wordDocument;
            _sdtBlock = sdtBlock;
        }

        /// <summary>
        /// Initializes a new instance using one of the predefined templates.
        /// </summary>
        /// <param name="wordDocument">Parent document.</param>
        /// <param name="coverPageTemplate">Template to insert.</param>
        public WordCoverPage(WordDocument wordDocument, WordCoverPageTemplate coverPageTemplate) {
            _document = wordDocument;
            _sdtBlock = GetStyle(coverPageTemplate);
            _document.AssignNewSdtIds(_sdtBlock);
            var body = _document._wordprocessingDocument?.MainDocumentPart?.Document?.Body
                ?? throw new InvalidOperationException("Document body is missing.");
            body.Append(_sdtBlock);
        }

        private SdtBlock GetStyle(WordCoverPageTemplate template) {
            switch (template) {
                case WordCoverPageTemplate.Austin: return CoverPageAustin;
                case WordCoverPageTemplate.Banded: return CoverPageBanded;
                case WordCoverPageTemplate.Facet: return CoverPageFacet;
                case WordCoverPageTemplate.Grid: return CoverPageGrid;
                case WordCoverPageTemplate.IonDark: return CoverPageIonDark;
                case WordCoverPageTemplate.IonLight: return CoverPageIonLight;
                case WordCoverPageTemplate.Element: return CoverPageElement;
                case WordCoverPageTemplate.Wisp: return CoverPageWisp;
                case WordCoverPageTemplate.ViewMaster: return CoverPageViewMaster;
                case WordCoverPageTemplate.SliceLight: return CoverPageSliceLight;
                case WordCoverPageTemplate.SliceDark: return CoverPageSliceDark;
                case WordCoverPageTemplate.SideLine: return CoverPageSideLine;
                case WordCoverPageTemplate.Semaphore: return CoverPageSemaphore;
                case WordCoverPageTemplate.Retrospect: return CoverPageRetrospect;
            }
            throw new ArgumentOutOfRangeException(nameof(template));
        }
    }
}
