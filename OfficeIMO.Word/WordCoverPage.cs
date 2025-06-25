using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Built-in cover page templates available for Word documents.
    /// </summary>
    public enum CoverPageTemplate {
        /// <summary>Template named Austin.</summary>
        Austin,
        /// <summary>Template named Banded.</summary>
        Banded,
        /// <summary>Template named Facet.</summary>
        Facet,
        /// <summary>Template named Grid.</summary>
        Grid,
        /// <summary>Template named IonDark.</summary>
        IonDark,
        /// <summary>Template named IonLight.</summary>
        IonLight,
        /// <summary>Template named Element.</summary>
        Element,
        /// <summary>Template named Wisp.</summary>
        Wisp,
        /// <summary>Template named ViewMaster.</summary>
        ViewMaster,
        /// <summary>Template named SliceLight.</summary>
        SliceLight,
        /// <summary>Template named SliceDark.</summary>
        SliceDark,
        /// <summary>Template named SideLine.</summary>
        SideLine,
        /// <summary>Template named Semaphore.</summary>
        Semaphore,
        /// <summary>Template named Retrospect.</summary>
        Retrospect
    }

    /// <summary>
    /// Represents a cover page within a Word document.
    /// </summary>
    public partial class WordCoverPage : WordElement {
        private readonly WordDocument _document;
        private readonly SdtBlock _sdtBlock;

        /// <summary>
        /// Initializes a new instance from an existing structured document tag block.
        /// </summary>
        /// <param name="wordDocument">Parent document.</param>
        /// <param name="sdtBlock">Structured document tag to wrap.</param>
        public WordCoverPage(WordDocument wordDocument, SdtBlock sdtBlock) {
            _document = wordDocument;
            _sdtBlock = sdtBlock;
        }

        /// <summary>
        /// Initializes a new instance using one of the predefined templates.
        /// </summary>
        /// <param name="wordDocument">Parent document.</param>
        /// <param name="coverPageTemplate">Template to insert.</param>
        public WordCoverPage(WordDocument wordDocument, CoverPageTemplate coverPageTemplate) {
            _document = wordDocument;
            _sdtBlock = GetStyle(coverPageTemplate);
            this._document._wordprocessingDocument.MainDocumentPart.Document.Body.Append(_sdtBlock);
        }

        private SdtBlock GetStyle(CoverPageTemplate template) {
            switch (template) {
                case CoverPageTemplate.Austin: return CoverPageAustin;
                case CoverPageTemplate.Banded: return CoverPageBanded;
                case CoverPageTemplate.Facet: return CoverPageFacet;
                case CoverPageTemplate.Grid: return CoverPageGrid;
                case CoverPageTemplate.IonDark: return CoverPageIonDark;
                case CoverPageTemplate.IonLight: return CoverPageIonLight;
                case CoverPageTemplate.Element: return CoverPageElement;
                case CoverPageTemplate.Wisp: return CoverPageWisp;
                case CoverPageTemplate.ViewMaster: return CoverPageViewMaster;
                case CoverPageTemplate.SliceLight: return CoverPageSliceLight;
                case CoverPageTemplate.SliceDark: return CoverPageSliceDark;
                case CoverPageTemplate.SideLine: return CoverPageSideLine;
                case CoverPageTemplate.Semaphore: return CoverPageSemaphore;
                case CoverPageTemplate.Retrospect: return CoverPageRetrospect;
            }
            throw new ArgumentOutOfRangeException(nameof(template));
        }
    }
}
