using System;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioConnectors {
        [Fact]
        public void ConnectorBetweenRectanglesEmitsGeometryAndRecalcFlag() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = new();
            document.RequestRecalcOnOpen();
            VisioPage page = document.AddPage("Page-1");
            VisioShape start = new("1", 1, 1, 1, 1, "Start");
            VisioShape end = new("2", 3, 2, 1, 1, "End");
            page.Shapes.Add(start);
            page.Shapes.Add(end);
            VisioConnector connector = new(start, end);
            page.Connectors.Add(connector);
            document.Save(filePath);

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";

            PackagePart documentPart = package.GetPart(new Uri("/visio/document.xml", UriKind.Relative));
            XDocument docXml = XDocument.Load(documentPart.GetStream());
            Assert.Equal("1", docXml.Root?.Element(ns + "DocumentSettings")?.Element(ns + "RelayoutAndRerouteUponOpen")?.Value);

            PackagePart pagePart = package.GetPart(new Uri("/visio/pages/page1.xml", UriKind.Relative));
            XDocument pageXml = XDocument.Load(pagePart.GetStream());
            XElement connectorShape = pageXml.Root?
                .Element(ns + "Shapes")?
                .Elements(ns + "Shape")
                .First(e => e.Attribute("ID")?.Value == connector.Id);

            XElement[] segments = connectorShape.Element(ns + "Geom")?.Elements().ToArray() ?? Array.Empty<XElement>();
            Assert.Equal("MoveTo", segments[0].Name.LocalName);
            Assert.Equal(1.5, double.Parse(segments[0].Attribute("X")!.Value, CultureInfo.InvariantCulture));
            Assert.Equal(1, double.Parse(segments[0].Attribute("Y")!.Value, CultureInfo.InvariantCulture));
            Assert.Equal("LineTo", segments[1].Name.LocalName);
            Assert.Equal(1.5, double.Parse(segments[1].Attribute("X")!.Value, CultureInfo.InvariantCulture));
            Assert.Equal(2, double.Parse(segments[1].Attribute("Y")!.Value, CultureInfo.InvariantCulture));
            Assert.Equal("LineTo", segments[2].Name.LocalName);
            Assert.Equal(2.5, double.Parse(segments[2].Attribute("X")!.Value, CultureInfo.InvariantCulture));
            Assert.Equal(2, double.Parse(segments[2].Attribute("Y")!.Value, CultureInfo.InvariantCulture));

            XElement connects = pageXml.Root?.Element(ns + "Connects") ?? new XElement("none");
            XElement[] entries = connects.Elements(ns + "Connect").ToArray();
            Assert.Equal(2, entries.Length);
            Assert.Contains(entries, e => e.Attribute("FromSheet")?.Value == connector.Id && e.Attribute("FromCell")?.Value == "BeginX" && e.Attribute("ToSheet")?.Value == start.Id);
            Assert.Contains(entries, e => e.Attribute("FromSheet")?.Value == connector.Id && e.Attribute("FromCell")?.Value == "EndX" && e.Attribute("ToSheet")?.Value == end.Id);
        }
    }
}
