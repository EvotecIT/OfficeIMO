using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioShapeAndConnectorStyles {
        [Fact]
        public void ShapesAndConnectorsHaveDefaultStyles() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");
            VisioShape start = new("1", 1, 1, 2, 1, "Start");
            VisioShape end = new("2", 4, 1, 2, 1, "End");
            page.Shapes.Add(start);
            page.Shapes.Add(end);
            VisioConnector connector = new(start, end) { EndArrow = 13 };
            page.Connectors.Add(connector);
            document.Save();

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            PackagePart pagePart = package.GetPart(new Uri("/visio/pages/page1.xml", UriKind.Relative));
            XDocument pageXml = XDocument.Load(pagePart.GetStream());
            XElement shapesRoot = pageXml.Root!.Element(ns + "Shapes")!;

            XElement shapeXml = shapesRoot.Elements(ns + "Shape").First(e => e.Attribute("ID")?.Value == "1");
            Assert.Equal("1", shapeXml.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "LinePattern").Attribute("V")?.Value);
            Assert.Equal("#000000", shapeXml.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "LineColor").Attribute("V")?.Value);
            Assert.Equal("1", shapeXml.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "FillPattern").Attribute("V")?.Value);
            Assert.Equal("#FFFFFF", shapeXml.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "FillForegnd").Attribute("V")?.Value);

            XElement connectorXml = shapesRoot.Elements(ns + "Shape").First(e => e.Attribute("ID")?.Value == connector.Id);
            Assert.Equal("1", connectorXml.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "LinePattern").Attribute("V")?.Value);
            Assert.Equal("#000000", connectorXml.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "LineColor").Attribute("V")?.Value);
            Assert.Equal("0", connectorXml.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "FillPattern").Attribute("V")?.Value);
            Assert.Equal("#000000", connectorXml.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "FillForegnd").Attribute("V")?.Value);
            Assert.Equal("1", connectorXml.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "OneD").Attribute("V")?.Value);
            Assert.Equal("13", connectorXml.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "EndArrow").Attribute("V")?.Value);
            Assert.Equal("2", connectorXml.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "BeginX").Attribute("V")?.Value);
            Assert.Equal("1", connectorXml.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "BeginY").Attribute("V")?.Value);
            Assert.Equal("3", connectorXml.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "EndX").Attribute("V")?.Value);
            Assert.Equal("1", connectorXml.Elements(ns + "Cell").First(c => c.Attribute("N")?.Value == "EndY").Attribute("V")?.Value);

            var badCells = pageXml.Descendants(ns + "Cell")
                .Where(c => (c.Attribute("N")?.Value == "NoLine" || c.Attribute("N")?.Value == "NoFill") && c.Attribute("V")?.Value == "1");
            Assert.Empty(badCells);
        }
    }
}
