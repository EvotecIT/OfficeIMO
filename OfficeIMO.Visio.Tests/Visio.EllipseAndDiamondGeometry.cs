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
        public void EllipseShapeEmitsClosedLineGeometry() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            var doc = VisioDocument.Create(filePath);
            doc.AsFluent().AddPage("Page-1", p => p.AddEllipse("E1", x:OfficeLength.Inches( 2), y:OfficeLength.Inches( 2), width:OfficeLength.Inches( 2), height:OfficeLength.Inches( 1)));
            doc.Save();

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            var pagePart = package.GetPart(new Uri("/visio/pages/page1.xml", UriKind.Relative));
            XDocument pageDoc = XDocument.Load(pagePart.GetStream());
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement shape = pageDoc.Root!.Element(ns + "Shapes")!.Elements(ns + "Shape")
                .First(e => e.Attribute("NameU")?.Value == "Ellipse");
            var geom = shape.Elements(ns + "Section").FirstOrDefault(s => s.Attribute("N")?.Value == "Geometry");
            Assert.NotNull(geom);
            var rows = geom!.Elements(ns + "Row").ToArray();
            Assert.True(rows.Count(r => r.Attribute("T")?.Value == "LineTo") >= 20);
        }

        [Fact]
        public void DiamondShapeEmitsLineGeometry() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            var doc = VisioDocument.Create(filePath);
            doc.AsFluent().AddPage("Page-1", p => p.AddDiamond("D1", x:OfficeLength.Inches( 4), y:OfficeLength.Inches( 4), width:OfficeLength.Inches( 2), height:OfficeLength.Inches( 2)));
            doc.Save();

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            var pagePart = package.GetPart(new Uri("/visio/pages/page1.xml", UriKind.Relative));
            XDocument pageDoc = XDocument.Load(pagePart.GetStream());
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement shape = pageDoc.Root!.Element(ns + "Shapes")!.Elements(ns + "Shape")
                .First(e => e.Attribute("NameU")?.Value == "Diamond");
            var geom = shape.Elements(ns + "Section").FirstOrDefault(s => s.Attribute("N")?.Value == "Geometry");
            Assert.NotNull(geom);
            var rows = geom!.Elements(ns + "Row").ToArray();
            Assert.True(rows.Count(r => (string?)r.Attribute("T") == "LineTo") >= 4);
        }
    }
}

