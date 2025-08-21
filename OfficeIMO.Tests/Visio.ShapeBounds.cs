using System;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioShapeBounds {
        [Fact]
        public void ComputesBoundsForRotatedShapesAndConnectors() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");
            VisioShape start = new("1", 2, 2, 2, 1, "Start") {
                Angle = Math.PI / 2,
            };
            VisioShape end = new("2", 5, 2, 2, 1, "End");
            page.Shapes.Add(start);
            page.Shapes.Add(end);
            VisioConnector connector = new(start, end);
            page.Connectors.Add(connector);

            (double left, double bottom, double right, double top) = start.GetBounds();
            Assert.Equal(1.5, left, 5);
            Assert.Equal(1, bottom, 5);
            Assert.Equal(2.5, right, 5);
            Assert.Equal(3, top, 5);

            document.Save();

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            PackagePart pagePart = package.GetPart(new Uri("/visio/pages/page1.xml", UriKind.Relative));
            XDocument pageXml = XDocument.Load(pagePart.GetStream());
            XElement connectorShape = pageXml.Root?
                .Element(ns + "Shapes")?
                .Elements(ns + "Shape")
                .First(e => e.Attribute("ID")?.Value == connector.Id);

            XElement[] segments = connectorShape.Element(ns + "Geom")?.Elements().ToArray() ?? Array.Empty<XElement>();
            Assert.Empty(segments);
        }
    }
}

