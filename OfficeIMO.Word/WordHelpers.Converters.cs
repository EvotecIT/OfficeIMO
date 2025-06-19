using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordHelpers {
        /// <summary>
        /// Given a document name, remove specified headers and footers from the document.
        /// When no types are provided all headers and footers are removed.
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="types">Header or footer types to remove</param>
        public static void RemoveHeadersAndFooters(string filename, params HeaderFooterValues[] types) {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filename, true)) {
                WordHeader.RemoveHeaders(doc, types);
                WordFooter.RemoveFooters(doc, types);
                doc.MainDocumentPart.Document.Save();
            }
        }
    }
}
