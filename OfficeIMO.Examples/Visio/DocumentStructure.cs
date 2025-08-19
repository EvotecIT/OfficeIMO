using System;
using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    public static class DocumentStructure {
        public static void Example_DocumentStructure(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Document structure");
            string filePath = Path.Combine(folderPath, "Document Structure.vsdx");

            VisioDocument document = new();
            document.Save(filePath);

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            PackagePart documentPart = package.GetPart(new Uri("/visio/document.xml", UriKind.Relative));
            XDocument xml = XDocument.Load(documentPart.GetStream());
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";

            Console.WriteLine(xml.Root?.Element(ns + "DocumentSettings") != null);
            Console.WriteLine(xml.Root?.Element(ns + "Colors") != null);
            Console.WriteLine(xml.Root?.Element(ns + "FaceNames") != null);
            Console.WriteLine(xml.Root?.Element(ns + "StyleSheets") != null);

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

