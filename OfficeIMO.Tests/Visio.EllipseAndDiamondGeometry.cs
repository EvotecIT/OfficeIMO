using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioEllipseAndDiamondGeometryTests {
        [Fact]
        public void EllipseShapeEmitsEllipticalArcGeometry() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            var doc = VisioDocument.Create(filePath);
            doc.AsFluent().Page("Page-1", p => p.Ellipse("E1", x: 2, y: 2, width: 2, height: 1));
            doc.Save();

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            var pagePart = package.GetPart(new Uri("/visio/pages/page1.xml", UriKind.Relative));
            XDocument pageDoc = XDocument.Load(pagePart.GetStream());
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement shape = pageDoc.Root!.Element(ns + "Shapes")!.Elements(ns + "Shape")
                .First(e => e.Attribute("ID")?.Value == "E1");
            var geom = shape.Elements(ns + "Section").FirstOrDefault(s => s.Attribute("N")?.Value == "Geometry");
            Assert.NotNull(geom);
            var rows = geom!.Elements(ns + "Row").ToArray();
            Assert.Contains(rows, r => r.Attribute("T")?.Value == "EllipticalArcTo");
        }

        [Fact]
        public void DiamondShapeEmitsLineGeometry() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            var doc = VisioDocument.Create(filePath);
            doc.AsFluent().Page("Page-1", p => p.Diamond("D1", x: 4, y: 4, width: 2, height: 2));
            doc.Save();

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            var pagePart = package.GetPart(new Uri("/visio/pages/page1.xml", UriKind.Relative));
            XDocument pageDoc = XDocument.Load(pagePart.GetStream());
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement shape = pageDoc.Root!.Element(ns + "Shapes")!.Elements(ns + "Shape")
                .First(e => e.Attribute("ID")?.Value == "D1");
            var geom = shape.Elements(ns + "Section").FirstOrDefault(s => s.Attribute("N")?.Value == "Geometry");
            Assert.NotNull(geom);
            var rows = geom!.Elements(ns + "Row").ToArray();
            Assert.True(rows.Count(r => (string?)r.Attribute("T") == "LineTo") >= 4);
        }
    }
}

