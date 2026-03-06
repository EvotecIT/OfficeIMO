using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioMastersParity {
        private static string AssetsPath => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets"));

        [Fact]
        public void TemplateLearningRegistersNamesButGeneratesMastersFromCode() {
            string template = Path.Combine(AssetsPath, "VisioTemplates", "DrawingWithShapes.vsdx");
            Assert.True(File.Exists(template));

            string target = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument doc = VisioDocument.Create(target);
            doc.UseMastersByDefault = true;
            doc.UseMastersFromTemplate(template);

            Assert.True(doc.TryGetMaster("Rectangle", out _));
            Assert.True(doc.TryGetMaster("Square", out _));
            Assert.True(doc.TryGetMaster("Circle", out _));
            Assert.True(doc.TryGetMaster("Dynamic connector", out _));
            Assert.False(doc.TryGetMaster("Ellipse", out _));

            VisioPage page = doc.AddPage("Page-1", 29.7, 21, VisioMeasurementUnit.Centimeters);
            page.Shapes.Add(new VisioShape("1") { NameU = "Rectangle", PinX = 2, PinY = 6 });
            page.Shapes.Add(new VisioShape("2") { NameU = "Square", PinX = 4, PinY = 6 });
            page.Shapes.Add(new VisioShape("3") { NameU = "Circle", PinX = 6, PinY = 6 });
            page.Shapes.Add(new VisioShape("4") { NameU = "Ellipse", PinX = 8, PinY = 6 });
            page.Shapes.Add(new VisioShape("5") { NameU = "Diamond", PinX = 10, PinY = 6 });
            page.Shapes.Add(new VisioShape("6") { NameU = "Triangle", PinX = 12, PinY = 6 });
            page.Connectors.Add(new VisioConnector("connector-1", page.Shapes[0], page.Shapes[1]) { Kind = ConnectorKind.Dynamic });
            doc.Save();

            using ZipArchive zip = ZipFile.OpenRead(target);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";

            XDocument mastersDoc = XDocument.Load(zip.GetEntry("visio/masters/masters.xml")!.Open());
            XDocument relsDoc = XDocument.Load(zip.GetEntry("visio/masters/_rels/masters.xml.rels")!.Open());

            string[] expectedNames = {
                "Rectangle",
                "Square",
                "Circle",
                "Ellipse",
                "Diamond",
                "Triangle",
                "Dynamic connector"
            };

            Assert.Equal(expectedNames, mastersDoc.Root!.Elements(ns + "Master").Select(e => (string?)e.Attribute("NameU")));

            foreach (XElement masterElement in mastersDoc.Root!.Elements(ns + "Master")) {
                string nameU = (string)masterElement.Attribute("NameU")!;
                string relationshipId = (string)masterElement.Element(ns + "Rel")!.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))!;
                XElement relationship = relsDoc.Root!.Elements(relNs + "Relationship").Single(e => (string?)e.Attribute("Id") == relationshipId);
                string targetPart = (string)relationship.Attribute("Target")!;

                Assert.Equal(nameU, (string?)masterElement.Attribute("Name"));
                Assert.False(string.IsNullOrWhiteSpace((string?)masterElement.Attribute("Prompt")));
                Assert.False(string.IsNullOrWhiteSpace((string?)masterElement.Attribute("UniqueID")));
                Assert.False(string.IsNullOrWhiteSpace((string?)masterElement.Attribute("BaseID")));

                XElement pageSheet = masterElement.Element(ns + "PageSheet")!;
                Assert.NotNull(pageSheet);

                XDocument masterContents = XDocument.Load(zip.GetEntry("visio/masters/" + targetPart)!.Open());
                XElement shape = masterContents.Root!.Element(ns + "Shapes")!.Element(ns + "Shape")!;

                Assert.Equal(nameU, (string?)shape.Attribute("NameU"));
                Assert.NotNull(masterContents.Root!.Element(ns + "Shapes")!.Element(ns + "Shape")!.Elements(ns + "Section").FirstOrDefault(e => (string?)e.Attribute("N") == "User"));
                Assert.NotNull(masterContents.Root!.Element(ns + "Shapes")!.Element(ns + "Shape")!.Elements(ns + "Section").FirstOrDefault(e => (string?)e.Attribute("N") == "Character"));
                Assert.All(shape.Elements(ns + "Section").Where(e => (string?)e.Attribute("N") == "Connection").Elements(ns + "Row"), row =>
                    Assert.Equal("Connection", (string?)row.Attribute("T")));

                if (string.Equals(nameU, "Dynamic connector", StringComparison.OrdinalIgnoreCase)) {
                    Assert.Equal("1", (string?)masterElement.Attribute("MatchByName"));
                    Assert.Equal("0", (string?)masterElement.Attribute("MasterType"));
                    Assert.NotNull(pageSheet.Elements(ns + "Section").FirstOrDefault(e => (string?)e.Attribute("N") == "Layer"));
                    Assert.NotNull(shape.Elements(ns + "Section").FirstOrDefault(e => (string?)e.Attribute("N") == "Control"));
                    Assert.NotNull(shape.Element(ns + "XForm1D"));
                    Assert.Contains(shape.Elements(ns + "Cell"), cell => (string?)cell.Attribute("N") == "OneD");
                    continue;
                }

                Assert.Equal("2", (string?)masterElement.Attribute("MasterType"));
                Assert.Contains(pageSheet.Elements(ns + "Cell"), cell => (string?)cell.Attribute("N") == "ShapeKeywords");
                XElement geometrySection = shape.Elements(ns + "Section").Single(section => (string?)section.Attribute("N") == "Geometry");
                Assert.NotEmpty(geometrySection.Elements(ns + "Row"));

                if (string.Equals(nameU, "Circle", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(nameU, "Square", StringComparison.OrdinalIgnoreCase)) {
                    Assert.Contains(shape.Elements(ns + "Cell"), cell => (string?)cell.Attribute("N") == "LockAspect");
                }
            }
        }

        [Fact]
        public void ImportMastersSkipsUnsupportedTemplateMasters() {
            string template = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            CreateTemplateWithMasters(template, "FancyHexagon", "Rectangle");

            VisioDocument doc = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            IReadOnlyList<VisioMaster> imported = doc.ImportMastersAndGet(template);

            Assert.Single(imported);
            Assert.Equal("Rectangle", imported[0].NameU);
            Assert.True(doc.TryGetMaster("Rectangle", out _));
            Assert.False(doc.TryGetMaster("FancyHexagon", out _));
        }

        private static void CreateTemplateWithMasters(string path, params string[] masterNames) {
            const string visioNamespace = "http://schemas.microsoft.com/office/visio/2012/main";
            const string relationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            using ZipArchive zip = ZipFile.Open(path, ZipArchiveMode.Create);
            ZipArchiveEntry mastersEntry = zip.CreateEntry("visio/masters/masters.xml");
            using Stream stream = mastersEntry.Open();
            using StreamWriter writer = new(stream, new UTF8Encoding(false));

            XNamespace ns = visioNamespace;
            XNamespace rel = relationshipNamespace;
            XElement root = new(ns + "Masters",
                masterNames.Select((name, index) =>
                    new XElement(ns + "Master",
                        new XAttribute("ID", index + 1),
                        new XAttribute("Name", name),
                        new XAttribute("NameU", name),
                        new XElement(ns + "Rel", new XAttribute(rel + "id", $"rId{index + 1}")))));

            XDocument document = new(root);
            writer.Write(document.Declaration + Environment.NewLine + document.ToString(SaveOptions.DisableFormatting));
        }
    }
}
