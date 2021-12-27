using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    public class HeadersAndFooters {
        public static void RemoveHeadersAndFooters(string filename) {
            // Given a document name, remove all of the headers and footers
            // from the document.
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filename, true)) {
                Headers.RemoveHeaders(doc);
                Footers.RemoveFooters(doc);
                // save document
                doc.MainDocumentPart.Document.Save();
            }
        }
    }
}