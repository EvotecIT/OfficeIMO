using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides helper methods for Word document manipulation.
    /// </summary>
    public partial class WordHelpers {
        /// <summary>
        /// Removes headers and footers from the file at <paramref name="filename"/>.
        /// When no <paramref name="types"/> are provided all headers and footers are removed.
        /// </summary>
        /// <param name="filename">Path to the document.</param>
        /// <param name="types">Header or footer types to remove.</param>
        public static void RemoveHeadersAndFooters(string filename, params HeaderFooterValues[] types) {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filename, true)) {
                WordHeader.RemoveHeaders(doc, types);
                WordFooter.RemoveFooters(doc, types);
                doc.MainDocumentPart.Document.Save();
            }
        }
    }
}
