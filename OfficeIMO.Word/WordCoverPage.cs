using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public enum CoverPageTemplate {
        Austin,
        Banded,
        Facet,
        FiliGree
    }

    public partial class WordCoverPage {
        private readonly WordDocument _document;
        private readonly SdtBlock _sdtBlock;

        public WordCoverPage(WordDocument wordDocument, SdtBlock sdtBlock) {
            _document = wordDocument;
            _sdtBlock = sdtBlock;
        }

        public WordCoverPage(WordDocument wordDocument, CoverPageTemplate coverPageTemplate) {
            _document = wordDocument;
            _sdtBlock = GetStyle(coverPageTemplate);
            this._document._wordprocessingDocument.MainDocumentPart.Document.Body.InsertAt(_sdtBlock, 0);
        }

        private SdtBlock GetStyle(CoverPageTemplate template) {
            switch (template) {
                case CoverPageTemplate.Austin: return CoverPageAustin;
                case CoverPageTemplate.Banded: return CoverPageBanded;
                case CoverPageTemplate.Facet: return CoverPageFacet;
                case CoverPageTemplate.FiliGree: return CoverPageFiliGree;
            }
            throw new ArgumentOutOfRangeException(nameof(template));
        }
    }
}
