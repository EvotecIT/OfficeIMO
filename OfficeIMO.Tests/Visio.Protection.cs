using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Tests {
    public class VisioProtectionTests {
        [Fact]
        public void ShapeProtectionSavesLoadsAndSupportsBulkEditing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string unlockedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Protection", 8.5, 6);

            VisioShape background = page.AddRectangle(4.25, 3, 7.5, 4.8, "Background");
            background.FillColor = Color.LightCyan;
            background.Protect(protection => protection.Size().Position().Selection().Formatting());

            VisioShape label = page.AddRectangle(4.25, 5.2, 3.2, 0.6, "Generated label");
            label.LockSize();
            label.Protection.LockTextEdit = true;
            label.Protection.LockDelete = true;

            VisioShape editable = page.AddRectangle(4.25, 3, 2, 1, "Editable");
            VisioConnector connector = page.AddConnector(label, editable, ConnectorKind.Dynamic, VisioSide.Bottom, VisioSide.Top);
            connector.Label = "protected route";
            connector.Protect(protection => protection.Endpoints().Text().Deletion());

            page.SelectContainingText("Generated")
                .Protect(protection => protection.Deletion().Text())
                .Fill(Color.LightYellow);
            page.SelectConnectorsWithProtection()
                .Protect(protection => protection.Formatting());

            Assert.Equal(2, page.ShapesWithProtection().Count);
            Assert.Single(page.ShapesWithProtection(protection => protection.LockTextEdit == true));
            Assert.Single(page.ConnectorsWithProtection());
            Assert.Single(page.ConnectorsWithProtection(protection => protection.LockBegin == true && protection.LockEnd == true));

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertProtectionXml(filePath, "Background", expectedLockWidth: "1", expectedLockMoveX: "1", expectedLockSelect: "1", expectedLockDelete: null);
            AssertProtectionXml(filePath, "Generated label", expectedLockWidth: "1", expectedLockMoveX: null, expectedLockSelect: null, expectedLockDelete: "1");
            AssertConnectorProtectionXml(filePath, "protected route", expectedLockBegin: "1", expectedLockEnd: "1", expectedLockTextEdit: "1", expectedLockFormat: "1");

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape loadedBackground = loaded.Pages[0].Shapes.Single(shape => shape.Text == "Background");
            Assert.True(loadedBackground.Protection.LockWidth);
            Assert.True(loadedBackground.Protection.LockHeight);
            Assert.True(loadedBackground.Protection.LockMoveX);
            Assert.True(loadedBackground.Protection.LockMoveY);
            Assert.True(loadedBackground.Protection.LockSelect);

            VisioShape loadedLabel = loaded.Pages[0].Shapes.Single(shape => shape.Text == "Generated label");
            Assert.True(loadedLabel.Protection.LockTextEdit);
            Assert.True(loadedLabel.Protection.LockDelete);

            VisioConnector loadedConnector = loaded.Pages[0].Connectors.Single(conn => conn.Label == "protected route");
            Assert.True(loadedConnector.Protection.LockBegin);
            Assert.True(loadedConnector.Protection.LockEnd);
            Assert.True(loadedConnector.Protection.LockTextEdit);
            Assert.True(loadedConnector.Protection.LockFormat);

            loaded.Pages[0].SelectWithProtection(protection => protection.LockSelect == true)
                .ClearProtection();
            loaded.Pages[0].SelectConnectorsWithProtection(protection => protection.LockBegin == true)
                .ClearProtection();
            loaded.Save(unlockedPath);

            Assert.Empty(VisioValidator.Validate(unlockedPath));
            AssertProtectionXml(unlockedPath, "Background", expectedLockWidth: null, expectedLockMoveX: null, expectedLockSelect: null, expectedLockDelete: null);
            AssertProtectionXml(unlockedPath, "Generated label", expectedLockWidth: "1", expectedLockMoveX: null, expectedLockSelect: null, expectedLockDelete: "1");
            AssertConnectorProtectionXml(unlockedPath, "protected route", expectedLockBegin: null, expectedLockEnd: null, expectedLockTextEdit: null, expectedLockFormat: null);
        }

        private static void AssertProtectionXml(
            string filePath,
            string shapeText,
            string? expectedLockWidth,
            string? expectedLockMoveX,
            string? expectedLockSelect,
            string? expectedLockDelete) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument page = ReadXml(archive, "visio/pages/page1.xml");
            XElement shape = page.Descendants(ns + "Shape")
                .Single(element => element.Element(ns + "Text")?.Value == shapeText);

            AssertCellValue(shape, ns, "LockWidth", expectedLockWidth);
            AssertCellValue(shape, ns, "LockMoveX", expectedLockMoveX);
            AssertCellValue(shape, ns, "LockSelect", expectedLockSelect);
            AssertCellValue(shape, ns, "LockDelete", expectedLockDelete);
        }

        private static void AssertCellValue(XElement shape, XNamespace ns, string cellName, string? expectedValue) {
            XElement[] cells = shape.Elements(ns + "Cell")
                .Where(cell => (string?)cell.Attribute("N") == cellName)
                .ToArray();

            if (expectedValue == null) {
                Assert.Empty(cells);
                return;
            }

            XElement cell = Assert.Single(cells);
            Assert.Equal(expectedValue, cell.Attribute("V")!.Value);
        }

        private static void AssertConnectorProtectionXml(
            string filePath,
            string connectorText,
            string? expectedLockBegin,
            string? expectedLockEnd,
            string? expectedLockTextEdit,
            string? expectedLockFormat) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument page = ReadXml(archive, "visio/pages/page1.xml");
            XElement connector = page.Descendants(ns + "Shape")
                .Single(element => element.Element(ns + "Text")?.Value == connectorText);

            AssertCellValue(connector, ns, "LockBegin", expectedLockBegin);
            AssertCellValue(connector, ns, "LockEnd", expectedLockEnd);
            AssertCellValue(connector, ns, "LockTextEdit", expectedLockTextEdit);
            AssertCellValue(connector, ns, "LockFormat", expectedLockFormat);
        }

        private static XDocument ReadXml(ZipArchive archive, string entryName) {
            ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException("Missing " + entryName);
            using Stream stream = entry.Open();
            return XDocument.Load(stream);
        }
    }
}
