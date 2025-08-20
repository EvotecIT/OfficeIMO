using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;

namespace OfficeIMO.Examples.Visio {
    /// <summary>
    /// Demonstrates style sheet definitions in a Visio document.
    /// </summary>
    public static class StyleSheets {
        public static void Example_StyleSheets(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - StyleSheets");
            string filePath = Path.Combine(folderPath, "StyleSheets.vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            VisioShape start = new("1", 1, 1, 2, 1, "Start");
            VisioShape end = new("2", 4, 1, 2, 1, "End");
            page.Shapes.Add(start);
            page.Shapes.Add(end);
            page.Connectors.Add(new VisioConnector(start, end));
            document.Save(filePath);

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            PackagePart docPart = package.GetPart(new Uri("/visio/document.xml", UriKind.Relative));
            XDocument docXml = XDocument.Load(docPart.GetStream());
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            var styles = docXml.Root!.Element(ns + "StyleSheets")!
                .Elements(ns + "StyleSheet")
                .Select(e => $"{e.Attribute("ID")?.Value}: {e.Attribute("NameU")?.Value}");
            foreach (string style in styles) {
                Console.WriteLine($"Style {style}");
            }

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
