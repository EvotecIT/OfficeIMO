using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioStyleSheets {
        [Fact]
        public void DocumentDefinesAndReferencesStyles() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

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
            XElement styleSheets = docXml.Root!.Element(ns + "StyleSheets")!;

            XElement normal = styleSheets.Elements(ns + "StyleSheet").First(e => e.Attribute("ID")?.Value == "1");
            Assert.Equal("Normal", normal.Attribute("NameU")?.Value);
            Assert.Equal("1", normal.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "LinePattern").Attribute("V")?.Value);
            Assert.Equal("RGB(0,0,0)", normal.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "LineColor").Attribute("V")?.Value);
            Assert.Equal("1", normal.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "FillPattern").Attribute("V")?.Value);
            Assert.Equal("RGB(255,255,255)", normal.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "FillForegnd").Attribute("V")?.Value);

            XElement connectorStyle = styleSheets.Elements(ns + "StyleSheet").First(e => e.Attribute("ID")?.Value == "2");
            Assert.Equal("1", connectorStyle.Attribute("BasedOn")?.Value);
            Assert.Equal("0", connectorStyle.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "EndArrow").Attribute("V")?.Value);

            PackagePart pagePart = package.GetPart(new Uri("/visio/pages/page1.xml", UriKind.Relative));
            XDocument pageXml = XDocument.Load(pagePart.GetStream());
            XElement shapesRoot = pageXml.Root!.Element(ns + "Shapes")!;
            XElement shapeXml = shapesRoot.Elements(ns + "Shape").First(e => e.Attribute("ID")?.Value == "1");
            Assert.Equal("1", shapeXml.Attribute("LineStyle")?.Value);
            Assert.Equal("1", shapeXml.Attribute("FillStyle")?.Value);
            Assert.Equal("1", shapeXml.Attribute("TextStyle")?.Value);
            XElement connectorXml = shapesRoot.Elements(ns + "Shape").First(e => e.Attribute("NameU")?.Value == "Connector");
            Assert.Equal("2", connectorXml.Attribute("LineStyle")?.Value);
            Assert.Equal("2", connectorXml.Attribute("FillStyle")?.Value);
            Assert.Equal("2", connectorXml.Attribute("TextStyle")?.Value);
        }
    }
}
