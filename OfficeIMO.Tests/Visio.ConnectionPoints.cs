using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.IO.Compression;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioConnectionPointsTests {
        [Fact]
        public void ConnectionPointsAreSavedAndLoaded() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");

            VisioShape from = new("1", 2, 2, 2, 2, "From");
            from.ConnectionPoints.Add(new VisioConnectionPoint(2, 1, 1, 0));
            page.Shapes.Add(from);

            VisioShape to = new("2", 6, 2, 2, 2, "To");
            to.ConnectionPoints.Add(new VisioConnectionPoint(0, 1, -1, 0));
            page.Shapes.Add(to);

            VisioConnector connector = new(from, to) {
                FromConnectionPoint = from.ConnectionPoints[0],
                ToConnectionPoint = to.ConnectionPoints[0]
            };
            page.Connectors.Add(connector);

            document.Save();

            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument pageXml = ReadPage(filePath);
            XElement shapeXml = pageXml.Root!
                .Element(ns + "Shapes")!
                .Elements(ns + "Shape")
                .First(e => e.Attribute("ID")?.Value == from.Id);
            XElement? connectionSection = shapeXml.Elements(ns + "Section")
                .FirstOrDefault(e => e.Attribute("N")?.Value == "Connection");
            Assert.NotNull(connectionSection);
            XElement[] rows = connectionSection!.Elements(ns + "Row").ToArray();
            Assert.Single(rows);
            Assert.Equal("0", rows[0].Attribute("IX")?.Value);

            XElement connects = pageXml.Root!.Element(ns + "Connects")!;
            XElement[] connectRows = connects.Elements(ns + "Connect").ToArray();
            Assert.Equal("Connections.X1", connectRows[0].Attribute("ToCell")?.Value);
            Assert.Equal("Connections.X1", connectRows[1].Attribute("ToCell")?.Value);
        }

        [Fact]
        public void ConnectionPointGlueRoundTripsThroughLoad() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");

            VisioShape from = new("1", 2, 2, 2, 2, "From");
            from.ConnectionPoints.Add(new VisioConnectionPoint(2, 1, 1, 0));
            page.Shapes.Add(from);

            VisioShape to = new("2", 6, 2, 2, 2, "To");
            to.ConnectionPoints.Add(new VisioConnectionPoint(0, 1, -1, 0));
            page.Shapes.Add(to);

            VisioConnector connector = new(from, to) {
                FromConnectionPoint = from.ConnectionPoints[0],
                ToConnectionPoint = to.ConnectionPoints[0]
            };
            page.Connectors.Add(connector);
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector loadedConnector = Assert.Single(loaded.Pages[0].Connectors);
            Assert.NotNull(loadedConnector.FromConnectionPoint);
            Assert.NotNull(loadedConnector.ToConnectionPoint);
            Assert.Same(loaded.Pages[0].Shapes[0].ConnectionPoints[0], loadedConnector.FromConnectionPoint);
            Assert.Same(loaded.Pages[0].Shapes[1].ConnectionPoints[0], loadedConnector.ToConnectionPoint);
        }

        [Fact]
        public void SparseConnectionPointIndicesRoundTripThroughLoadAndSave() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");

            VisioShape from = new("1", 2, 2, 2, 2, "From");
            from.ConnectionPoints.Add(new VisioConnectionPoint(2, 1, 1, 0));
            page.Shapes.Add(from);

            VisioShape to = new("2", 6, 2, 2, 2, "To");
            to.ConnectionPoints.Add(new VisioConnectionPoint(0, 1, -1, 0));
            page.Shapes.Add(to);

            VisioConnector connector = new(from, to) {
                FromConnectionPoint = from.ConnectionPoints[0],
                ToConnectionPoint = to.ConnectionPoints[0]
            };
            page.Connectors.Add(connector);
            document.Save();

            RewritePage(filePath, pageXml => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";

                XElement shapes = pageXml.Root!.Element(ns + "Shapes")!;
                SetConnectionRowIndex(
                    shapes.Elements(ns + "Shape").First(e => e.Attribute("ID")?.Value == from.Id),
                    ns,
                    2);
                SetConnectionRowIndex(
                    shapes.Elements(ns + "Shape").First(e => e.Attribute("ID")?.Value == to.Id),
                    ns,
                    5);

                foreach (XElement connect in pageXml.Root!.Element(ns + "Connects")!.Elements(ns + "Connect")) {
                    if ((string?)connect.Attribute("ToSheet") == from.Id) {
                        connect.SetAttributeValue("ToCell", "Connections.X3");
                    } else if ((string?)connect.Attribute("ToSheet") == to.Id) {
                        connect.SetAttributeValue("ToCell", "Connections.X6");
                    }
                }
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector loadedConnector = Assert.Single(loaded.Pages[0].Connectors);
            Assert.Same(loaded.Pages[0].Shapes[0].ConnectionPoints[0], loadedConnector.FromConnectionPoint);
            Assert.Same(loaded.Pages[0].Shapes[1].ConnectionPoints[0], loadedConnector.ToConnectionPoint);

            loaded.Save();

            XNamespace verifyNs = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument savedPageXml = ReadPage(filePath);
            XElement savedShapes = savedPageXml.Root!.Element(verifyNs + "Shapes")!;
            Assert.Equal(
                "2",
                GetConnectionRowIndex(savedShapes.Elements(verifyNs + "Shape").First(e => e.Attribute("ID")?.Value == from.Id), verifyNs));
            Assert.Equal(
                "5",
                GetConnectionRowIndex(savedShapes.Elements(verifyNs + "Shape").First(e => e.Attribute("ID")?.Value == to.Id), verifyNs));

            XElement[] connectRows = savedPageXml.Root!.Element(verifyNs + "Connects")!.Elements(verifyNs + "Connect").ToArray();
            Assert.Equal("Connections.X3", connectRows.First(e => (string?)e.Attribute("ToSheet") == from.Id).Attribute("ToCell")?.Value);
            Assert.Equal("Connections.X6", connectRows.First(e => (string?)e.Attribute("ToSheet") == to.Id).Attribute("ToCell")?.Value);
        }

        [Fact]
        public void SparseConnectionPointIndicesDoNotFallbackToSequentialGlueWhenReferenceIsMissing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");

            VisioShape from = new("1", 2, 2, 2, 2, "From");
            from.ConnectionPoints.Add(new VisioConnectionPoint(2, 1, 1, 0));
            page.Shapes.Add(from);

            VisioShape to = new("2", 6, 2, 2, 2, "To");
            to.ConnectionPoints.Add(new VisioConnectionPoint(0, 1, -1, 0));
            page.Shapes.Add(to);

            VisioConnector connector = new(from, to) {
                FromConnectionPoint = from.ConnectionPoints[0],
                ToConnectionPoint = to.ConnectionPoints[0]
            };
            page.Connectors.Add(connector);
            document.Save();

            RewritePage(filePath, pageXml => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement shapes = pageXml.Root!.Element(ns + "Shapes")!;
                SetConnectionRowIndex(
                    shapes.Elements(ns + "Shape").First(e => e.Attribute("ID")?.Value == from.Id),
                    ns,
                    2);
                SetConnectionRowIndex(
                    shapes.Elements(ns + "Shape").First(e => e.Attribute("ID")?.Value == to.Id),
                    ns,
                    5);

                foreach (XElement connect in pageXml.Root!.Element(ns + "Connects")!.Elements(ns + "Connect")) {
                    if ((string?)connect.Attribute("ToSheet") == from.Id) {
                        connect.SetAttributeValue("ToCell", "Connections.X1");
                    } else if ((string?)connect.Attribute("ToSheet") == to.Id) {
                        connect.SetAttributeValue("ToCell", "Connections.X6");
                    }
                }
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector loadedConnector = Assert.Single(loaded.Pages[0].Connectors);

            Assert.Null(loadedConnector.FromConnectionPoint);
            Assert.Same(loaded.Pages[0].Shapes[1].ConnectionPoints[0], loadedConnector.ToConnectionPoint);
        }

        private static XDocument ReadPage(string vsdxPath) {
            using FileStream stream = File.Open(vsdxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using ZipArchive archive = new(stream, ZipArchiveMode.Read);
            using Stream pageStream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            return XDocument.Load(pageStream);
        }

        private static void RewritePage(string vsdxPath, Action<XDocument> transform) {
            using FileStream stream = File.Open(vsdxPath, FileMode.Open, FileAccess.ReadWrite);
            using ZipArchive archive = new(stream, ZipArchiveMode.Update);
            ZipArchiveEntry pageEntry = archive.GetEntry("visio/pages/page1.xml")!;
            XDocument pageXml;
            using (Stream pageStream = pageEntry.Open()) {
                pageXml = XDocument.Load(pageStream);
            }

            transform(pageXml);
            pageEntry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry("visio/pages/page1.xml");
            using Stream replacementStream = replacement.Open();
            pageXml.Save(replacementStream);
        }

        private static void SetConnectionRowIndex(XElement shapeElement, XNamespace ns, int rowIndex) {
            XElement row = shapeElement.Elements(ns + "Section")
                .First(e => e.Attribute("N")?.Value == "Connection")
                .Elements(ns + "Row")
                .Single();
            row.SetAttributeValue("IX", rowIndex.ToString());
        }

        private static string? GetConnectionRowIndex(XElement shapeElement, XNamespace ns) {
            return shapeElement.Elements(ns + "Section")
                .First(e => e.Attribute("N")?.Value == "Connection")
                .Elements(ns + "Row")
                .Single()
                .Attribute("IX")
                ?.Value;
        }
    }
}
