using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    public partial class WordHelpers {
        /// <summary>
        /// Given a document name, remove all of the headers and footers from the document.
        /// </summary>
        /// <param name="filename"></param>
        public static void RemoveHeadersAndFooters(string filename) {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filename, true)) {
                WordHeader.RemoveHeaders(doc);
                WordFooter.RemoveFooters(doc);
                // save document
                doc.MainDocumentPart.Document.Save();
            }
        }
    }
}
