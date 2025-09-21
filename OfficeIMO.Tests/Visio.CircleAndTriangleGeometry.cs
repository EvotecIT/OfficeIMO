using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioCircleAndTriangleGeometryTests {
        [Fact]
        public void CircleAndTriangleEmitExpectedGeometry() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            var doc = VisioDocument.Create(filePath);
            doc.AsFluent().Page("Page-1", p => p
                .Circle("C1", x: 2, y: 2, diameter: 1)
                .Triangle("T1", x: 4, y: 4, width: 2, height: 1));
            doc.Save();

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            var pagePart = package.GetPart(new Uri("/visio/pages/page1.xml", UriKind.Relative));
            XDocument pageDoc = XDocument.Load(pagePart.GetStream());
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            var shapes = pageDoc.Root!.Element(ns + "Shapes")!.Elements(ns + "Shape").ToArray();
            Assert.Equal(2, shapes.Length);
            var circleGeom = shapes.First(s => (string?)s.Attribute("ID") == "C1").Elements(ns + "Section").FirstOrDefault(e => (string?)e.Attribute("N") == "Geometry");
            Assert.NotNull(circleGeom);
            Assert.Contains(circleGeom!.Elements(ns + "Row"), r => (string?)r.Attribute("T") == "EllipticalArcTo");
            var triGeom = shapes.First(s => (string?)s.Attribute("ID") == "T1").Elements(ns + "Section").FirstOrDefault(e => (string?)e.Attribute("N") == "Geometry");
            Assert.NotNull(triGeom);
            Assert.True(triGeom!.Elements(ns + "Row").Count(r => (string?)r.Attribute("T") == "LineTo") >= 3);
        }
    }
}

