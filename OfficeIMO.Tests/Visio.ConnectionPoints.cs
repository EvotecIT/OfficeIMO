using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.IO.Compression;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioConnectionPointsTests {
        [Fact(Skip = "File locking behavior on CI causes this test to be unreliable")]
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

            byte[] data = File.ReadAllBytes(filePath);
            using MemoryStream ms = new(data);
            using ZipArchive archive = new(ms, ZipArchiveMode.Read);
            ZipArchiveEntry pageEntry = archive.GetEntry("visio/pages/page1.xml")!;
            using Stream pageStream = pageEntry.Open();
            XDocument pageXml = XDocument.Load(pageStream);

            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
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
    }
}
