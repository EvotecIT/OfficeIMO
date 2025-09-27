using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using SixLabors.ImageSharp;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioConnectorStylesRoundTrip {
        [Fact]
        public void StyledConnectorPreservesAppearanceAfterReload() {
            string initialPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string roundTripPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            try {
                VisioDocument document = VisioDocument.Create(initialPath);
                VisioPage page = document.AddPage("Page-1");
                VisioShape start = new("1", 1, 4, 2, 1, "Start");
                VisioShape end = new("2", 4, 4, 2, 1, "End");
                page.Shapes.Add(start);
                page.Shapes.Add(end);

                VisioConnector connector = new(start, end) {
                    BeginArrow = EndArrow.Arrow,
                    EndArrow = EndArrow.Triangle
                };
                connector.LineWeight = 0.08;
                connector.LinePattern = 4;
                connector.LineColor = Color.LimeGreen;
                page.Connectors.Add(connector);
                document.Save();

                using (Package package = Package.Open(initialPath, FileMode.Open, FileAccess.ReadWrite)) {
                    PackagePart pagePart = package.GetPart(new Uri("/visio/pages/page1.xml", UriKind.Relative));
                    XDocument pageXml;
                    using (Stream readStream = pagePart.GetStream(FileMode.Open, FileAccess.Read)) {
                        pageXml = XDocument.Load(readStream);
                    }
                    XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                    XElement connectorElement = pageXml.Root!
                        .Element(ns + "Shapes")!
                        .Elements(ns + "Shape")
                        .First(e => e.Attribute("ID")?.Value == connector.Id);

                    static void SetCell(XElement shape, XNamespace ns, string name, string value) {
                        XElement? cell = shape.Elements(ns + "Cell").FirstOrDefault(e => e.Attribute("N")?.Value == name);
                        if (cell == null) {
                            cell = new XElement(ns + "Cell", new XAttribute("N", name));
                            shape.Add(cell);
                        }
                        cell.SetAttributeValue("V", value);
                        cell.SetAttributeValue("Result", value);
                    }

                    SetCell(connectorElement, ns, "LineColor", "THEMEGUARD(RGB(10,20,30))");
                    SetCell(connectorElement, ns, "LinePattern", "6");
                    SetCell(connectorElement, ns, "LineWeight", "0.123");

                    using Stream writeStream = pagePart.GetStream(FileMode.Create, FileAccess.Write);
                    pageXml.Save(writeStream);
                }

                VisioDocument loaded = VisioDocument.Load(initialPath);
                VisioConnector loadedConnector = loaded.Pages[0].Connectors.Single();
                Color expectedColor = Color.FromRgb(10, 20, 30);
                Assert.Equal(0.123, loadedConnector.LineWeight, 5);
                Assert.Equal(6, loadedConnector.LinePattern);
                Assert.Equal(expectedColor, loadedConnector.LineColor);

                loaded.Save(roundTripPath);

                VisioDocument roundTripped = VisioDocument.Load(roundTripPath);
                VisioConnector finalConnector = roundTripped.Pages[0].Connectors.Single();
                Assert.Equal(0.123, finalConnector.LineWeight, 5);
                Assert.Equal(6, finalConnector.LinePattern);
                Assert.Equal(expectedColor, finalConnector.LineColor);
            } finally {
                if (File.Exists(initialPath)) {
                    File.Delete(initialPath);
                }
                if (File.Exists(roundTripPath)) {
                    File.Delete(roundTripPath);
                }
            }
        }
    }
}
