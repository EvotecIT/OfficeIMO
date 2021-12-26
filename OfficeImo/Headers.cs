using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeImo
{
    public class Headers
    {
        public static void RemoveHeaders(WordprocessingDocument wordprocessingDocument)
        {
            var docPart = wordprocessingDocument.MainDocumentPart;
            Document document = docPart.Document;
            if (docPart.HeaderParts.Count() > 0)
            {
                // Remove the header
                docPart.DeleteParts(docPart.HeaderParts);

                // First, create a list of all descendants of type
                // HeaderReference. Then, navigate the list and call
                // Remove on each item to delete the reference.
                var headers = document.Descendants<HeaderReference>().ToList();
                foreach (var header in headers)
                {
                    header.Remove();
                }
            }
        }
    }
}
