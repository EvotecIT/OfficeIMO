using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static class Footers {
        public static void RemoveFooters(WordprocessingDocument wordprocessingDocument) {
            var docPart = wordprocessingDocument.MainDocumentPart;
            DocumentFormat.OpenXml.Wordprocessing.Document document = docPart.Document;
            if (docPart.FooterParts.Count() > 0) {
                // Remove the header
                docPart.DeleteParts(docPart.FooterParts);

                // First, create a list of all descendants of type
                // HeaderReference. Then, navigate the list and call
                // Remove on each item to delete the reference.
                var footers = document.Descendants<FooterReference>().ToList();
                foreach (var footer in footers) {
                    footer.Remove();
                }
            }
        }
    }
}