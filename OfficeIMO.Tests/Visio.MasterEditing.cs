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
    public class VisioMasterEditingTests {
        [Fact]
        public void ReplaceMasterUpdatesSelectionWithoutLosingMetadataOrConnectors() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Replace Master", 11, 8.5);

            VisioShape source = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "source", 2, 6, "Source");
            VisioShape task = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "task", 5, 6, "Review");
            VisioShape target = page.AddStencilShape(VisioStencils.Flowchart.Get("data"), "target", 8, 6, "Archive");
            task.FillColor = Color.LightYellow;
            task.LineColor = Color.Orange;
            task.LayerNames.Add("Review");
            task.SetShapeData("Owner", "Operations", "Owner", VisioShapeDataType.String, "Owning team");
            task.SetUserCell("Stage", "Review", "STR");
            task.AddHyperlink("https://example.org/review", "Review docs");
            task.Protection.Text().Deletion();

            VisioConnector incoming = page.AddConnector(source, task, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            VisioConnector outgoing = page.AddConnector(task, target, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);

            page.SelectByMaster("Process")
                .ReplaceMaster(VisioStencils.Flowchart.Get("decision"), resizeToMaster: true);

            Assert.Equal("Decision", source.MasterNameU);
            Assert.Equal("Decision", task.MasterNameU);
            Assert.Equal(2.0, task.Width, 6);
            Assert.Equal(1.4, task.Height, 6);
            Assert.Equal(1.0, task.LocPinX, 6);
            Assert.Equal(0.7, task.LocPinY, 6);
            Assert.Equal(Color.LightYellow, task.FillColor);
            Assert.Equal(Color.Orange, task.LineColor);
            Assert.Contains("Review", task.LayerNames);
            Assert.Equal("Operations", task.GetShapeDataValue("Owner"));
            Assert.Equal("Owning team", task.FindShapeData("Owner")!.Prompt);
            Assert.Equal("Review", task.GetUserCellValue("Stage"));
            Assert.Equal("https://example.org/review", task.Hyperlinks.Single().Address);
            Assert.True(task.Protection.LockTextEdit);
            Assert.True(task.Protection.LockDelete);
            Assert.Same(task, incoming.To);
            Assert.Same(task, outgoing.From);
            Assert.Equal("Data", target.MasterNameU);

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertShapeMaster(filePath, "Review", "Decision");

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape loadedTask = loaded.Pages[0].Shapes.Single(shape => shape.Text == "Review");
            Assert.Equal("Decision", loadedTask.MasterNameU);
            Assert.Equal(2.0, loadedTask.Width, 6);
            Assert.Equal(1.4, loadedTask.Height, 6);
            Assert.Equal("Operations", loadedTask.GetShapeDataValue("Owner"));
            Assert.Equal("Review", loadedTask.GetUserCellValue("Stage"));
            Assert.Single(loaded.Pages[0].IncomingConnectors(loadedTask));
            Assert.Single(loaded.Pages[0].OutgoingConnectors(loadedTask));
        }

        [Fact]
        public void ReplaceMasterCanKeepCurrentShapeSize() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Keep Size");
            VisioShape shape = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "shape", 2, 4, 3.2, 0.8, "Large step");

            page.ReplaceMaster(shape, "Decision");

            Assert.Equal("Decision", shape.MasterNameU);
            Assert.Equal(3.2, shape.Width, 6);
            Assert.Equal(0.8, shape.Height, 6);
            Assert.Equal(1.6, shape.LocPinX, 6);
            Assert.Equal(0.4, shape.LocPinY, 6);
        }

        private static void AssertShapeMaster(string filePath, string shapeText, string expectedMasterNameU) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument pageDoc = ReadXml(archive, "visio/pages/page1.xml");
            XDocument mastersDoc = ReadXml(archive, "visio/masters/masters.xml");

            XElement shape = pageDoc.Descendants(ns + "Shape")
                .Single(element => element.Element(ns + "Text")?.Value == shapeText);
            string masterId = shape.Attribute("Master")?.Value ?? throw new InvalidOperationException("Shape did not reference a master.");
            XElement master = mastersDoc.Descendants(ns + "Master")
                .Single(element => element.Attribute("ID")?.Value == masterId);
            Assert.Equal(expectedMasterNameU, master.Attribute("NameU")?.Value);
        }

        private static XDocument ReadXml(ZipArchive archive, string entryName) {
            ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException("Missing " + entryName);
            using Stream stream = entry.Open();
            return XDocument.Load(stream);
        }
    }
}
