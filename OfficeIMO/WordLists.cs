using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO {
    public class WordLists {
        private WordprocessingDocument _wordprocessingDocument;
        private WordDocument _document;
        internal NumberingDefinitionsPart _numberingDefinitionsPart;

        public WordLists(WordDocument document) {
            _document = document;
            _wordprocessingDocument = document._wordprocessingDocument;

            //NumberingDefinitionsPart numberingDefinitionsPart = document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart;
            //if (numberingDefinitionsPart == null) {
            //    numberingDefinitionsPart = _wordprocessingDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
            //}
            //_numberingDefinitionsPart = numberingDefinitionsPart;

        }
    }
}
