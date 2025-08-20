using System;
using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioDocumentSettings {
        [Fact]
        public void DocumentHasDefaultStyles() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = new();
            document.AddPage("Page-1");
            document.Save(filePath);

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            PackagePart docPart = package.GetPart(new Uri("/visio/document.xml", UriKind.Relative));
            XDocument docXml = XDocument.Load(docPart.GetStream());
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement settings = docXml.Root!.Element(ns + "DocumentSettings")!;
            Assert.Equal("3", settings.Attribute("DefaultLineStyle")?.Value);
            Assert.Equal("3", settings.Attribute("DefaultFillStyle")?.Value);
        }
    }
}
