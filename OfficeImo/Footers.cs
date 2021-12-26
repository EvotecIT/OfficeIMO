using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeImo
{
    public class Footers
    {
        public static void RemoveFooters(WordprocessingDocument wordprocessingDocument)
        {
            var docPart = wordprocessingDocument.MainDocumentPart;
            Document document = docPart.Document;
            if (docPart.FooterParts.Count() > 0)
            {
                // Remove the header
                docPart.DeleteParts(docPart.FooterParts);

                // First, create a list of all descendants of type
                // HeaderReference. Then, navigate the list and call
                // Remove on each item to delete the reference.
                var footers = document.Descendants<FooterReference>().ToList();
                foreach (var footer in footers)
                {
                    footer.Remove();
                }
            }
        }
    }    
}
