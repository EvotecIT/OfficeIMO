using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioRectangleGeometry {
        [Fact]
        public void RectangleShapeHasProperGeometry() {
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
            XElement geom = shape.Element(ns + "Geom")!;
            var lines = geom.Elements(ns + "LineTo").ToList();
            Assert.Equal(4, lines.Count);
            Assert.Equal("2", lines[0].Attribute("X")!.Value);
            Assert.Equal("0", lines[0].Attribute("Y")!.Value);
            Assert.Equal("2", lines[1].Attribute("X")!.Value);
            Assert.Equal("1", lines[1].Attribute("Y")!.Value);
            Assert.Equal("0", lines[2].Attribute("X")!.Value);
            Assert.Equal("1", lines[2].Attribute("Y")!.Value);
            Assert.Equal("0", lines[3].Attribute("X")!.Value);
            Assert.Equal("0", lines[3].Attribute("Y")!.Value);

            var noFill = geom.Elements(ns + "Cell").FirstOrDefault(e => e.Attribute("N")?.Value == "NoFill");
            Assert.NotNull(noFill);
            Assert.Equal("0", noFill!.Attribute("V")!.Value);
            var noLine = geom.Elements(ns + "Cell").FirstOrDefault(e => e.Attribute("N")?.Value == "NoLine");
            Assert.NotNull(noLine);
            Assert.Equal("0", noLine!.Attribute("V")!.Value);
        }

    }
}
