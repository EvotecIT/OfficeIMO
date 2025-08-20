using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioLineWeight {
        [Fact]
        public void ShapeContainsLineWeight() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Rect") { NameU = "Rectangle" });
            document.Save(filePath);

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            PackagePart pagePart = package.GetPart(new Uri("/visio/pages/page1.xml", UriKind.Relative));
            XDocument pageDoc = XDocument.Load(pagePart.GetStream());
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement shape = pageDoc.Root!.Element(ns + "Shapes")!.Element(ns + "Shape")!;
            XElement? lineWeight = shape.Elements(ns + "Cell")
                .FirstOrDefault(e => e.Attribute("N")?.Value == "LineWeight");
            Assert.NotNull(lineWeight);
        }
    }
}
