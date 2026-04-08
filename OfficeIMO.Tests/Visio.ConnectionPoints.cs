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

        [Fact]
        public void UnresolvedConnectionCellReferencesArePreservedOnRoundTrip() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");

            VisioShape from = new("1", 2, 2, 2, 2, "From");
            page.Shapes.Add(from);

            VisioShape to = new("2", 6, 2, 2, 2, "To");
            page.Shapes.Add(to);

            page.Connectors.Add(new VisioConnector(from, to));
            document.Save();

            RewritePage(filePath, pageXml => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";

                foreach (XElement connect in pageXml.Root!.Element(ns + "Connects")!.Elements(ns + "Connect")) {
                    if ((string?)connect.Attribute("FromCell") == "BeginX") {
                        connect.SetAttributeValue("ToCell", "LocPinX");
                    } else if ((string?)connect.Attribute("FromCell") == "EndX") {
                        connect.SetAttributeValue("ToCell", "Width");
                    }
                }
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector loadedConnector = Assert.Single(loaded.Pages[0].Connectors);
            Assert.Null(loadedConnector.FromConnectionPoint);
            Assert.Null(loadedConnector.ToConnectionPoint);

            loaded.Save();

            XNamespace verifyNs = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement[] connectRows = ReadPage(filePath).Root!.Element(verifyNs + "Connects")!.Elements(verifyNs + "Connect").ToArray();
            Assert.Equal("LocPinX", connectRows.First(e => (string?)e.Attribute("FromCell") == "BeginX").Attribute("ToCell")?.Value);
            Assert.Equal("Width", connectRows.First(e => (string?)e.Attribute("FromCell") == "EndX").Attribute("ToCell")?.Value);
        }

        [Fact]
        public void ExtraConnectAttributesArePreservedOnRoundTrip() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");

            VisioShape from = new("1", 2, 2, 2, 2, "From");
            VisioShape to = new("2", 6, 2, 2, 2, "To");
            page.Shapes.Add(from);
            page.Shapes.Add(to);
            page.Connectors.Add(new VisioConnector(from, to));
            document.Save();

            RewritePage(filePath, pageXml => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";

                foreach (XElement connect in pageXml.Root!.Element(ns + "Connects")!.Elements(ns + "Connect")) {
                    if ((string?)connect.Attribute("FromCell") == "BeginX") {
                        connect.SetAttributeValue("ToPart", "9");
                        connect.SetAttributeValue("Del", "1");
                    } else if ((string?)connect.Attribute("FromCell") == "EndX") {
                        connect.SetAttributeValue("ToPart", "12");
                    }
                }
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            loaded.Save();

            XNamespace verifyNs = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement[] connectRows = ReadPage(filePath).Root!.Element(verifyNs + "Connects")!.Elements(verifyNs + "Connect").ToArray();
            XElement begin = connectRows.First(e => (string?)e.Attribute("FromCell") == "BeginX");
            XElement end = connectRows.First(e => (string?)e.Attribute("FromCell") == "EndX");
            Assert.Equal("9", begin.Attribute("ToPart")?.Value);
            Assert.Equal("1", begin.Attribute("Del")?.Value);
            Assert.Equal("12", end.Attribute("ToPart")?.Value);
        }

        [Fact]
        public void ReconnectClearsPreservedAttributesOnlyForUpdatedEndpoint() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");

            VisioShape from = new("1", 2, 2, 2, 2, "From");
            VisioShape replacement = new("2", 4, 2, 2, 2, "Replacement");
            VisioShape to = new("3", 7, 2, 2, 2, "To");
            page.Shapes.Add(from);
            page.Shapes.Add(replacement);
            page.Shapes.Add(to);
            page.Connectors.Add(new VisioConnector(from, to));
            document.Save();

            RewritePage(filePath, pageXml => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";

                foreach (XElement connect in pageXml.Root!.Element(ns + "Connects")!.Elements(ns + "Connect")) {
                    if ((string?)connect.Attribute("FromCell") == "BeginX") {
                        connect.SetAttributeValue("ToPart", "9");
                    } else if ((string?)connect.Attribute("FromCell") == "EndX") {
                        connect.SetAttributeValue("ToPart", "12");
                    }
                }
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages[0];
            VisioConnector connector = Assert.Single(loadedPage.Connectors);

            loadedPage.ReconnectConnectorStart(connector, loadedPage.Shapes[1], VisioSide.Left);
            loaded.Save();

            XNamespace verifyNs = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement[] connectRows = ReadPage(filePath).Root!.Element(verifyNs + "Connects")!.Elements(verifyNs + "Connect").ToArray();
            XElement begin = connectRows.First(e => (string?)e.Attribute("FromCell") == "BeginX");
            XElement end = connectRows.First(e => (string?)e.Attribute("FromCell") == "EndX");
            Assert.Null(begin.Attribute("ToPart"));
            Assert.Equal("Connections.X1", begin.Attribute("ToCell")?.Value);
            Assert.Equal("12", end.Attribute("ToPart")?.Value);
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
