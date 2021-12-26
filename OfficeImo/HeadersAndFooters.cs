using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeImo
{
    public class HeadersAndFooters
    {
        public static void RemoveHeadersAndFooters(string filename)
        {
            // Given a document name, remove all of the headers and footers
            // from the document.
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filename, true))
            {
                Headers.RemoveHeaders(doc);
                Footers.RemoveFooters(doc);
                // save document
                doc.MainDocumentPart.Document.Save();
            }
        }
    }
}
