using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    internal class ImportTableStyle {
        private static XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        internal static void ImportTableStyles(string sourcefilepath, string destinationfilepath) {
            using (var repeaterSourceDocument = WordprocessingDocument.Open(sourcefilepath, true)) {
                XDocument source_style_doc;

                var repeaterSourceDocumentPackagePart = repeaterSourceDocument.MainDocumentPart.OpenXmlPackage.Package.GetPart(new Uri("/word/styles.xml", UriKind.Relative));

                //Get styles.xml
                using (TextReader tr = new StreamReader(repeaterSourceDocumentPackagePart.GetStream())) {
                    source_style_doc = XDocument.Load(tr);
                }

                var tableStylesFromRepeaterSource = source_style_doc.Descendants(w + "style").Where(x => x.Attribute(w + "type").Value == "table").Select(x => x).ToList();

                using (var targetFileToImportTableStyles = WordprocessingDocument.Open(destinationfilepath, true)) {
                    XDocument dest_style_doc;

                    var destpart = targetFileToImportTableStyles.MainDocumentPart.OpenXmlPackage.Package.GetPart(new Uri("/word/styles.xml", UriKind.Relative));

                    //Get styles.xml
                    using (TextReader trd = new StreamReader(destpart.GetStream())) {
                        dest_style_doc = XDocument.Load(trd);
                    }

                    //Add all the style elements from source document styles.xml 
                    foreach (var styleelement in tableStylesFromRepeaterSource) {
                        if (!dest_style_doc.Elements(XName.Get("styles", w.NamespaceName)).Any(x => (string)x.Attribute("styleId") == (string)styleelement.Attribute("styleId"))) {
                            dest_style_doc.Element(XName.Get("styles", w.NamespaceName)).Add(styleelement);
                        }
                    }

                    //Save the style.xml of targetFile
                    using (TextWriter tw = new StreamWriter(destpart.GetStream(FileMode.Create))) {
                        dest_style_doc.Save(tw, SaveOptions.None);
                    }

                }
            }
        }
    }
}
