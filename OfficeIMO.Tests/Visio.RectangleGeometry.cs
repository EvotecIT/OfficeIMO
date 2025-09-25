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
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Rect") { NameU = "Rectangle" });
            document.Save();

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            PackagePart pagePart = package.GetPart(new Uri("/visio/pages/page1.xml", UriKind.Relative));
            XDocument pageDoc = XDocument.Load(pagePart.GetStream());
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement shape = pageDoc.Root!.Element(ns + "Shapes")!.Element(ns + "Shape")!;
            // In master-delta mode these may be omitted; when present they should be 0
            var lineStyle = shape.Attribute("LineStyle")?.Value;
            var fillStyle = shape.Attribute("FillStyle")?.Value;
            var textStyle = shape.Attribute("TextStyle")?.Value;
            if (lineStyle != null) Assert.Equal("0", lineStyle);
            if (fillStyle != null) Assert.Equal("0", fillStyle);
            if (textStyle != null) Assert.Equal("0", textStyle);
            XElement? geom = shape.Element(ns + "Geom");
            Assert.Null(geom);
        }

    }
}
