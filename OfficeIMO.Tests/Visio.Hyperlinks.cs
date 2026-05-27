using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Tests {
    public class VisioHyperlinkTests {
        [Fact]
        public void ShapeAndConnectorHyperlinksSaveLoadAndExposeQueries() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string roundTripPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Links", 8.5, 6);

            VisioShape start = page.AddRectangle(2, 4, 2, 1, "Start");
            VisioShape finish = page.AddRectangle(6, 4, 2, 1, "Finish");
            VisioConnector connector = page.AddConnector(start, finish, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            connector.Label = "details";

            VisioHyperlink shapeLink = start.AddHyperlink("https://github.com/EvotecIT/OfficeIMO", "OfficeIMO", "README.md");
            shapeLink.NewWindow = true;
            connector.AddHyperlink("https://example.org/connector", "Connector details");

            page.SelectWithHyperlink("https://github.com/EvotecIT/OfficeIMO").Fill(Color.LightYellow);
            page.SelectConnectorsWithHyperlinks().EndArrow(EndArrow.Triangle);

            Assert.Single(page.ShapesWithHyperlinks());
            Assert.Single(page.ShapesWithHyperlink("https://github.com/EvotecIT/OfficeIMO"));
            Assert.Single(page.ConnectorsWithHyperlinks());

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertHyperlinkXml(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages[0];
            VisioShape loadedStart = loadedPage.Shapes.Single(shape => shape.Text == "Start");
            VisioHyperlink loadedShapeLink = Assert.Single(loadedStart.Hyperlinks);
            Assert.Equal("OfficeIMO", loadedShapeLink.Description);
            Assert.Equal("https://github.com/EvotecIT/OfficeIMO", loadedShapeLink.Address);
            Assert.Equal("README.md", loadedShapeLink.SubAddress);
            Assert.True(loadedShapeLink.NewWindow);

            VisioHyperlink loadedConnectorLink = Assert.Single(loadedPage.Connectors.Single().Hyperlinks);
            Assert.Equal("Connector details", loadedConnectorLink.Description);
            Assert.Equal("https://example.org/connector", loadedConnectorLink.Address);

            loadedPage.SelectWithHyperlinks().Hyperlink("https://example.org/extra", "Extra link");
            loaded.Save(roundTripPath);
            Assert.Empty(VisioValidator.Validate(roundTripPath));
        }

        private static void AssertHyperlinkXml(string filePath) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument page = ReadXml(archive, "visio/pages/page1.xml");

            XElement start = page.Descendants(ns + "Shape").Single(shape => (string?)shape.Attribute("ID") == "1");
            XElement shapeRow = start.Elements(ns + "Section")
                .Single(section => (string?)section.Attribute("N") == "Hyperlink")
                .Elements(ns + "Row")
                .Single();
            Assert.Equal("OfficeIMO", CellValue(shapeRow, ns, "Description"));
            Assert.Equal("https://github.com/EvotecIT/OfficeIMO", CellValue(shapeRow, ns, "Address"));
            Assert.Equal("README.md", CellValue(shapeRow, ns, "SubAddress"));
            Assert.Equal("1", CellValue(shapeRow, ns, "NewWindow"));
            Assert.Equal("0", CellValue(shapeRow, ns, "Invisible"));

            XElement connector = page.Descendants(ns + "Shape").Single(shape => (string?)shape.Attribute("ID") == "3");
            XElement connectorRow = connector.Elements(ns + "Section")
                .Single(section => (string?)section.Attribute("N") == "Hyperlink")
                .Elements(ns + "Row")
                .Single();
            Assert.Equal("Connector details", CellValue(connectorRow, ns, "Description"));
            Assert.Equal("https://example.org/connector", CellValue(connectorRow, ns, "Address"));
        }

        private static string CellValue(XElement row, XNamespace ns, string cellName) {
            return row.Elements(ns + "Cell")
                .Single(cell => (string?)cell.Attribute("N") == cellName)
                .Attribute("V")!
                .Value;
        }

        private static XDocument ReadXml(ZipArchive archive, string entryName) {
            ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException("Missing " + entryName);
            using Stream stream = entry.Open();
            return XDocument.Load(stream);
        }
    }
}
