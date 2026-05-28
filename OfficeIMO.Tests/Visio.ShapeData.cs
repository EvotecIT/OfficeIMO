using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Tests {
    public class VisioShapeDataTests {
        [Fact]
        public void TypedShapeDataSavesLoadsPreservesMetadataAndKeepsDictionaryCompatibility() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Shape Data", 8.5, 6);
            VisioShape server = page.AddRectangle(2.5, 4, 2.2, 1, "Server");
            server.SetShapeData("Owner", "Operations", "Owner", VisioShapeDataType.String, "Owning support team");
            VisioShapeDataRow cost = server.SetShapeData("MonthlyCost", "1250", "Monthly cost", VisioShapeDataType.Currency, "Estimated monthly cost", "$#,##0");
            cost.SortKey = "020";
            cost.Verify = true;

            page.AddRectangle(6, 4, 2.2, 1, "Database")
                .SetShapeData("Owner", "Data", "Owner", VisioShapeDataType.String, "Owning support team");

            page.SelectWithShapeData("Owner", "Operations")
                .Fill(Color.LightBlue)
                .ShapeData("Reviewed", "Yes", "Reviewed", VisioShapeDataType.Boolean, "Architecture review complete");

            Assert.Single(page.ShapesWithShapeData("Owner", "Operations"));
            Assert.Single(page.ShapesWithData("Reviewed", "Yes"));

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertShapeDataXml(filePath, "Operations", "Yes");

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape loadedServer = loaded.Pages[0].Shapes.Single(shape => shape.Text == "Server");
            Assert.Equal("Operations", loadedServer.GetShapeDataValue("Owner"));
            Assert.Equal("Owner", loadedServer.FindShapeData("Owner")!.Label);
            Assert.Equal("Owning support team", loadedServer.FindShapeData("Owner")!.Prompt);
            Assert.Equal(VisioShapeDataType.Currency, loadedServer.FindShapeData("MonthlyCost")!.Type);
            Assert.Equal("$#,##0", loadedServer.FindShapeData("MonthlyCost")!.Format);
            Assert.True(loadedServer.FindShapeData("MonthlyCost")!.Verify);

            loadedServer.Data["Owner"] = "Platform";
            Assert.Equal("Platform", loadedServer.GetShapeDataValue("Owner"));
            loaded.Save(updatedPath);

            Assert.Empty(VisioValidator.Validate(updatedPath));
            AssertShapeDataXml(updatedPath, "Platform", "Yes");
        }

        [Fact]
        public void ConnectorShapeDataSetReusesExistingRowNameCasing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Connector Data", 8.5, 6);
            VisioShape source = page.AddRectangle(2, 4, 1.5, 0.75, "Source");
            VisioShape target = page.AddRectangle(5, 4, 1.5, 0.75, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Dynamic);
            connector.Label = "route";
            connector.SetShapeData("Owner", "Operations", "Owner", VisioShapeDataType.String);
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector loadedConnector = loaded.Pages[0].Connectors.Single();
            loadedConnector.SetShapeData("owner", "Platform");
            Assert.Equal("Platform", loadedConnector.Data["Owner"]);
            Assert.False(loadedConnector.Data.ContainsKey("owner"));
            loaded.Save(updatedPath);

            Assert.Empty(VisioValidator.Validate(updatedPath));
            AssertConnectorShapeDataXml(updatedPath, "Owner", "Platform");
        }

        [Fact]
        public void ConnectorShapeDataClearOverridesPreservedLoadedValue() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Connector Data", 8.5, 6);
            VisioShape source = page.AddRectangle(2, 4, 1.5, 0.75, "Source");
            VisioShape target = page.AddRectangle(5, 4, 1.5, 0.75, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Dynamic);
            connector.SetShapeData("Owner", "Operations", "Owner", VisioShapeDataType.String);
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector loadedConnector = loaded.Pages[0].Connectors.Single();
            loadedConnector.SetShapeData("Owner", null);
            Assert.Equal(string.Empty, loadedConnector.GetShapeDataValue("Owner"));
            Assert.False(loadedConnector.Data.ContainsKey("Owner"));
            loaded.Save(updatedPath);

            Assert.Empty(VisioValidator.Validate(updatedPath));
            AssertConnectorShapeDataXml(updatedPath, "Owner", string.Empty);
        }

        [Fact]
        public void ConnectorShapeDataRowValueEditWinsOverStaleDictionaryMirror() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Connector Data", 8.5, 6);
            VisioShape source = page.AddRectangle(2, 4, 1.5, 0.75, "Source");
            VisioShape target = page.AddRectangle(5, 4, 1.5, 0.75, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Dynamic);
            VisioShapeDataRow owner = connector.SetShapeData("Owner", "Operations", "Owner", VisioShapeDataType.String);
            owner.Value = "Platform";

            Assert.Equal("Platform", connector.GetShapeDataValue("Owner"));
            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertConnectorShapeDataXml(filePath, "Owner", "Platform");
        }

        private static void AssertShapeDataXml(string filePath, string ownerValue, string reviewedValue) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument page = ReadXml(archive, "visio/pages/page1.xml");
            XElement server = page.Descendants(ns + "Shape")
                .Single(shape => shape.Element(ns + "Text")?.Value == "Server");
            XElement propSection = server.Elements(ns + "Section")
                .Single(section => (string?)section.Attribute("N") == "Prop");

            XElement owner = Row(propSection, ns, "Owner");
            Assert.Equal(ownerValue, CellValue(owner, ns, "Value"));
            Assert.Equal("Owner", CellValue(owner, ns, "Label"));
            Assert.Equal("Owning support team", CellValue(owner, ns, "Prompt"));
            Assert.Equal(((int)VisioShapeDataType.String).ToString(), CellValue(owner, ns, "Type"));

            XElement cost = Row(propSection, ns, "MonthlyCost");
            Assert.Equal("1250", CellValue(cost, ns, "Value"));
            Assert.Equal(((int)VisioShapeDataType.Currency).ToString(), CellValue(cost, ns, "Type"));
            Assert.Equal("$#,##0", CellValue(cost, ns, "Format"));
            Assert.Equal("1", CellValue(cost, ns, "Verify"));

            XElement reviewed = Row(propSection, ns, "Reviewed");
            Assert.Equal(reviewedValue, CellValue(reviewed, ns, "Value"));
            Assert.Equal("Reviewed", CellValue(reviewed, ns, "Label"));
            Assert.Equal(((int)VisioShapeDataType.Boolean).ToString(), CellValue(reviewed, ns, "Type"));
        }

        private static void AssertConnectorShapeDataXml(string filePath, string rowName, string ownerValue) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument page = ReadXml(archive, "visio/pages/page1.xml");
            XElement connector = page.Descendants(ns + "Shape")
                .Single(shape => shape.Elements(ns + "Section")
                    .Any(section => (string?)section.Attribute("N") == "Prop" &&
                        section.Elements(ns + "Row")
                            .Any(row => string.Equals((string?)row.Attribute("N"), rowName, StringComparison.Ordinal))));
            XElement propSection = connector.Elements(ns + "Section")
                .Single(section => (string?)section.Attribute("N") == "Prop");

            Assert.Single(propSection.Elements(ns + "Row"),
                row => string.Equals((string?)row.Attribute("N"), rowName, StringComparison.Ordinal));
            Assert.DoesNotContain(propSection.Elements(ns + "Row"),
                row => string.Equals((string?)row.Attribute("N"), rowName.ToLowerInvariant(), StringComparison.Ordinal));
            Assert.Equal(ownerValue, CellValue(Row(propSection, ns, rowName), ns, "Value"));
        }

        private static XElement Row(XElement section, XNamespace ns, string name) {
            return section.Elements(ns + "Row")
                .Single(row => (string?)row.Attribute("N") == name);
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
