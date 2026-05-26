using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Tests {
    public class VisioLayerTests {
        [Fact]
        public void PageLayersSaveLoadAndExposeLayerQueries() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Layered", 11, 8.5);
            VisioLayer infrastructure = page.AddLayer("Infrastructure");
            infrastructure.Color = 2;
            VisioLayer annotations = page.AddLayer("Annotations");
            annotations.Visible = false;
            annotations.Print = false;

            VisioShape server = page.AddStencilShape(VisioStencils.Network.Get("server"), "server", 2, 5, "Server");
            VisioShape note = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "note", 5, 5, "Note");
            VisioConnector connector = page.AddConnector(server, note, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            connector.Label = "documents";

            page.AddToLayer("Infrastructure", server);
            page.AddToLayer("Annotations", note);
            server.LayerNames.Add("Shared");
            connector.LayerNames.Add("Infrastructure");

            page.SelectLayer("Infrastructure").Stroke(Color.DodgerBlue, 0.02);
            Assert.Single(page.ShapesInLayer("Annotations"));
            Assert.Single(page.ConnectorsInLayer("Infrastructure"));
            Assert.Equal(Color.DodgerBlue, server.LineColor);

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertLayerXml(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages[0];
            Assert.Equal(3, loadedPage.Layers.Count);
            Assert.False(loadedPage.FindLayer("Annotations")!.Visible);
            Assert.False(loadedPage.FindLayer("Annotations")!.Print);
            Assert.Contains("Infrastructure", loadedPage.FindShapeById("server")!.LayerNames);
            Assert.Contains("Shared", loadedPage.FindShapeById("server")!.LayerNames);
            Assert.Contains("Infrastructure", loadedPage.Connectors.Single().LayerNames);
            Assert.Equal(2, loadedPage.ShapesInLayer("Infrastructure").Count + loadedPage.ShapesInLayer("Annotations").Count);
        }

        private static void AssertLayerXml(string filePath) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";

            XDocument pages = ReadXml(archive, "visio/pages/pages.xml");
            XElement layerSection = pages.Descendants(ns + "Section")
                .Single(section => string.Equals(section.Attribute("N")?.Value, "Layer", StringComparison.OrdinalIgnoreCase));
            XElement[] rows = layerSection.Elements(ns + "Row").ToArray();
            Assert.Equal(3, rows.Length);
            Assert.Equal("Infrastructure", rows[0].Elements(ns + "Cell").Single(cell => (string?)cell.Attribute("N") == "Name").Attribute("V")!.Value);
            Assert.Equal("0", rows[1].Elements(ns + "Cell").Single(cell => (string?)cell.Attribute("N") == "Visible").Attribute("V")!.Value);
            Assert.Equal("Shared", rows[2].Elements(ns + "Cell").Single(cell => (string?)cell.Attribute("N") == "Name").Attribute("V")!.Value);

            XDocument page = ReadXml(archive, "visio/pages/page1.xml");
            XElement server = page.Descendants(ns + "Shape").Single(shape => (string?)shape.Attribute("ID") == "1");
            XElement note = page.Descendants(ns + "Shape").Single(shape => (string?)shape.Attribute("ID") == "2");
            XElement connector = page.Descendants(ns + "Shape").Single(shape => (string?)shape.Attribute("ID") == "3");
            Assert.Equal("0;2", server.Elements(ns + "Cell").Single(cell => (string?)cell.Attribute("N") == "LayerMember").Attribute("V")!.Value);
            Assert.Equal("1", note.Elements(ns + "Cell").Single(cell => (string?)cell.Attribute("N") == "LayerMember").Attribute("V")!.Value);
            Assert.Equal("0", connector.Elements(ns + "Cell").Single(cell => (string?)cell.Attribute("N") == "LayerMember").Attribute("V")!.Value);
        }

        private static XDocument ReadXml(ZipArchive archive, string entryName) {
            ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException("Missing " + entryName);
            using Stream stream = entry.Open();
            return XDocument.Load(stream);
        }
    }
}
