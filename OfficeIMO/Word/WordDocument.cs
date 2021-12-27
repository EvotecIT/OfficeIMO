using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        public static WordprocessingDocument Create(string filePath, bool autoSave = false, WordprocessingDocumentType documentType = WordprocessingDocumentType.Document) {
            WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, documentType, autoSave);
            wordDocument.AddMainDocumentPart();
            wordDocument.MainDocumentPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
            wordDocument.MainDocumentPart.Document.Body = new DocumentFormat.OpenXml.Wordprocessing.Body();
            AddDefaultStyleDefinitions(wordDocument, null);
            return wordDocument;
        }
    }
}