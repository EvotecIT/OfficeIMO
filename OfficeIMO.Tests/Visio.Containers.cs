using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Tests {
    public class VisioContainerTests {
        [Fact]
        public void ContainersSaveNativeUserCellsRelationshipsAndLoadAsSemanticShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string roundTripPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Containers", 11, 8.5);
            VisioShape api = page.AddRectangle(3, 5.5, 1.7, 0.8, "API");
            VisioShape worker = page.AddRectangle(6, 5.5, 1.7, 0.8, "Worker");

            VisioShape container = page.AddContainer("app-tier", "Application tier", new[] { api, worker }, new VisioContainerOptions {
                Margin = 0.3,
                HeadingHeight = 0.4,
                FillColor = Color.LightCyan,
                LineColor = Color.DodgerBlue
            });
            container.SetUserCell("OfficeIMO.Role", "Tier", "STR", prompt: "OfficeIMO semantic role");
            page.SelectContainers().Stroke(Color.DodgerBlue, 0.02);
            page.SelectWithUserCell("OfficeIMO.Role", "Tier").UserCell("OfficeIMO.Reviewed", "Yes", "STR");

            Assert.True(container.IsContainer);
            Assert.Equal(new[] { api.Id, worker.Id }, container.ContainerMemberIds.ToArray());
            Assert.Contains(container.Id, api.ContainerOwnerIds);
            Assert.Single(page.Containers());
            Assert.Single(page.ShapesWithUserCell("OfficeIMO.Role", "Tier"));

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            AssertContainerXml(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages[0];
            VisioShape loadedContainer = Assert.Single(loadedPage.Containers());
            Assert.True(loadedContainer.IsContainer);
            Assert.Equal("Container", loadedContainer.GetUserCellValue("msvStructureType"));
            Assert.Equal("Yes", loadedContainer.GetUserCellValue("OfficeIMO.Reviewed"));
            Assert.Equal(2, loadedContainer.ContainerMemberIds.Count);
            Assert.All(loadedContainer.ContainerMemberIds, memberId => Assert.Contains(loadedContainer.Id, loadedPage.FindShapeById(memberId)!.ContainerOwnerIds));

            loaded.Save(roundTripPath);
            Assert.Empty(VisioValidator.Validate(roundTripPath));
            AssertContainerXml(roundTripPath);
        }

        private static void AssertContainerXml(string filePath) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument page = ReadXml(archive, "visio/pages/page1.xml");

            XElement container = page.Descendants(ns + "Shape")
                .Single(shape => shape.Element(ns + "Text")?.Value == "Application tier");
            string containerId = container.Attribute("ID")!.Value;

            XElement userSection = container.Elements(ns + "Section")
                .Single(section => (string?)section.Attribute("N") == "User");
            Assert.Equal("Container", UserCellValue(userSection, ns, "msvStructureType"));
            Assert.Equal("0.3", UserCellValue(userSection, ns, "msvSDContainerMargin"));
            Assert.Equal("1", UserCellValue(userSection, ns, "msvSDContainerResize"));
            Assert.Equal("Tier", UserCellValue(userSection, ns, "OfficeIMO.Role"));
            Assert.Equal("Yes", UserCellValue(userSection, ns, "OfficeIMO.Reviewed"));

            XElement relationshipCell = container.Elements(ns + "Cell")
                .Single(cell => (string?)cell.Attribute("N") == "Relationships");
            string formula = relationshipCell.Attribute("F")!.Value;
            Assert.Contains("DEPENDSON(1,Sheet.", formula);

            XElement[] memberShapes = page.Descendants(ns + "Shape")
                .Where(shape => shape.Element(ns + "Text")?.Value is "API" or "Worker")
                .ToArray();
            Assert.Equal(2, memberShapes.Length);
            foreach (XElement memberShape in memberShapes) {
                XElement memberRelationship = memberShape.Elements(ns + "Cell")
                    .Single(cell => (string?)cell.Attribute("N") == "Relationships");
                Assert.Contains($"DEPENDSON(4,Sheet.{containerId}!SheetRef())", memberRelationship.Attribute("F")!.Value);
            }
        }

        private static string UserCellValue(XElement userSection, XNamespace ns, string rowName) {
            XElement row = userSection.Elements(ns + "Row")
                .Single(element => (string?)element.Attribute("N") == rowName);
            return row.Elements(ns + "Cell")
                .Single(cell => (string?)cell.Attribute("N") == "Value")
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
