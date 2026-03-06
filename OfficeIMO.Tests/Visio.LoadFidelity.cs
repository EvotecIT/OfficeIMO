using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioLoadFidelity {
        [Fact]
        public void LoadDetectsStraightConnectorWhenGeometryHasHeaderRow() {
            string filePath = CreateConnectorDocument(ConnectorKind.Straight);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement connectorShape = GetConnectorShape(pageDoc, ns);
                XElement geometry = connectorShape.Elements(ns + "Section").First(section => (string?)section.Attribute("N") == "Geometry");
                geometry.AddFirst(
                    new XElement(ns + "Row",
                        new XAttribute("T", "Geometry"),
                        new XElement(ns + "Cell", new XAttribute("N", "NoFill"), new XAttribute("V", "0")),
                        new XElement(ns + "Cell", new XAttribute("N", "NoLine"), new XAttribute("V", "0"))));
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector connector = Assert.Single(loaded.Pages[0].Connectors);
            Assert.Equal(ConnectorKind.Straight, connector.Kind);
        }

        [Fact]
        public void LoadDetectsRightAngleConnectorFromOrthogonalPolyline() {
            string filePath = CreateConnectorDocument(ConnectorKind.RightAngle);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement connectorShape = GetConnectorShape(pageDoc, ns);
                XElement geometry = connectorShape.Elements(ns + "Section").First(section => (string?)section.Attribute("N") == "Geometry");
                geometry.RemoveNodes();
                geometry.Add(
                    new XElement(ns + "Row",
                        new XAttribute("T", "Geometry"),
                        new XElement(ns + "Cell", new XAttribute("N", "NoFill"), new XAttribute("V", "0"))),
                    CreateRow(ns, "MoveTo", 1, 1),
                    CreateRow(ns, "LineTo", 1, 3),
                    CreateRow(ns, "LineTo", 4, 3),
                    CreateRow(ns, "LineTo", 4, 2));
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector connector = Assert.Single(loaded.Pages[0].Connectors);
            Assert.Equal(ConnectorKind.RightAngle, connector.Kind);
        }

        [Fact]
        public void LoadDetectsCurvedConnectorFromArcGeometry() {
            string filePath = CreateConnectorDocument(ConnectorKind.Straight);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement connectorShape = GetConnectorShape(pageDoc, ns);
                XElement geometry = connectorShape.Elements(ns + "Section").First(section => (string?)section.Attribute("N") == "Geometry");
                geometry.RemoveNodes();
                geometry.Add(
                    CreateRow(ns, "MoveTo", 1, 1),
                    new XElement(ns + "Row",
                        new XAttribute("T", "ArcTo"),
                        new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "4")),
                        new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "2")),
                        new XElement(ns + "Cell", new XAttribute("N", "A"), new XAttribute("V", "0.5"))));
            });

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector connector = Assert.Single(loaded.Pages[0].Connectors);
            Assert.Equal(ConnectorKind.Curved, connector.Kind);
        }

        [Fact]
        public void ThemeXmlRoundTripsWithoutLosingCustomContent() {
            string filePath = CreateThemedDocument();
            string originalThemeXml = """
                <?xml version="1.0" encoding="utf-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
                  <a:themeElements>
                    <a:clrScheme name="Custom Colors" />
                  </a:themeElements>
                  <a:objectDefaults />
                </a:theme>
                """;
            RewriteEntry(filePath, "visio/theme/theme1.xml", originalThemeXml);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.True(XNode.DeepEquals(
                NormalizeTheme(originalThemeXml),
                NormalizeTheme(ReadEntry(savedPath, "visio/theme/theme1.xml"))));
        }

        [Fact]
        public void ThemeNameCanChangeWithoutDroppingCustomThemeStructure() {
            string filePath = CreateThemedDocument();
            RewriteEntry(filePath, "visio/theme/theme1.xml", """
                <?xml version="1.0" encoding="utf-8"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
                  <a:themeElements>
                    <a:fmtScheme name="Formatting" />
                  </a:themeElements>
                </a:theme>
                """);

            VisioDocument loaded = VisioDocument.Load(filePath);
            loaded.Theme!.Name = "Renamed Theme";

            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            XDocument savedTheme = XDocument.Parse(ReadEntry(savedPath, "visio/theme/theme1.xml"));
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            Assert.Equal("Renamed Theme", savedTheme.Root?.Attribute("name")?.Value);
            Assert.NotNull(savedTheme.Root?.Element(a + "themeElements"));
            Assert.NotNull(savedTheme.Root?.Element(a + "themeElements")?.Element(a + "fmtScheme"));
        }

        [Fact]
        public void CustomConnectorGeometryIsPreservedOnRoundTrip() {
            string filePath = CreateConnectorDocument(ConnectorKind.Straight);

            RewritePage(filePath, pageDoc => {
                XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
                XElement connectorShape = GetConnectorShape(pageDoc, ns);
                XElement geometry = connectorShape.Elements(ns + "Section").First(section => (string?)section.Attribute("N") == "Geometry");
                geometry.RemoveNodes();
                geometry.Add(
                    new XElement(ns + "Row",
                        new XAttribute("T", "Geometry"),
                        new XElement(ns + "Cell", new XAttribute("N", "NoFill"), new XAttribute("V", "0"))),
                    CreateRow(ns, "MoveTo", 1, 1),
                    CreateRow(ns, "LineTo", 2, 4),
                    CreateRow(ns, "LineTo", 4, 3),
                    CreateRow(ns, "LineTo", 5, 2));
            });

            string originalGeometry = ReadConnectorGeometry(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            string savedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(savedPath);

            Assert.True(XNode.DeepEquals(
                XElement.Parse(originalGeometry),
                XElement.Parse(ReadConnectorGeometry(savedPath))));
        }

        private static string CreateConnectorDocument(ConnectorKind kind) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");
            VisioShape start = new("1", 1, 1, 1, 1, "Start");
            VisioShape end = new("2", 4, 2, 1, 1, "End");
            page.Shapes.Add(start);
            page.Shapes.Add(end);
            page.Connectors.Add(new VisioConnector(start, end) { Kind = kind });
            document.Save();
            return filePath;
        }

        private static string CreateThemedDocument() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.Theme = new VisioTheme { Name = "Office Theme" };
            VisioPage page = document.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, string.Empty));
            document.Save();
            return filePath;
        }

        private static XElement GetConnectorShape(XDocument pageDoc, XNamespace ns) {
            return pageDoc.Root!
                .Element(ns + "Shapes")!
                .Elements(ns + "Shape")
                .Last();
        }

        private static XElement CreateRow(XNamespace ns, string type, double x, double y) {
            return new XElement(ns + "Row",
                new XAttribute("T", type),
                new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", x.ToString(System.Globalization.CultureInfo.InvariantCulture))),
                new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", y.ToString(System.Globalization.CultureInfo.InvariantCulture))));
        }

        private static void RewritePage(string vsdxPath, Action<XDocument> transform) {
            using FileStream stream = File.Open(vsdxPath, FileMode.Open, FileAccess.ReadWrite);
            using ZipArchive archive = new(stream, ZipArchiveMode.Update);
            ZipArchiveEntry pageEntry = archive.GetEntry("visio/pages/page1.xml")!;
            XDocument pageDoc;
            using (Stream pageStream = pageEntry.Open()) {
                pageDoc = XDocument.Load(pageStream);
            }

            transform(pageDoc);
            pageEntry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry("visio/pages/page1.xml");
            using Stream replacementStream = replacement.Open();
            using StreamWriter writer = new(replacementStream, new UTF8Encoding(false));
            writer.Write(pageDoc.Declaration + Environment.NewLine + pageDoc.ToString(SaveOptions.DisableFormatting));
        }

        private static void RewriteEntry(string vsdxPath, string entryPath, string content) {
            using FileStream stream = File.Open(vsdxPath, FileMode.Open, FileAccess.ReadWrite);
            using ZipArchive archive = new(stream, ZipArchiveMode.Update);
            ZipArchiveEntry entry = archive.GetEntry(entryPath)!;
            entry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry(entryPath);
            using Stream replacementStream = replacement.Open();
            using StreamWriter writer = new(replacementStream, new UTF8Encoding(false));
            writer.Write(content.Replace("\r\n", "\n"));
        }

        private static string ReadEntry(string vsdxPath, string entryPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry(entryPath)!.Open();
            using StreamReader reader = new(stream, Encoding.UTF8);
            return reader.ReadToEnd();
        }

        private static string ReadConnectorGeometry(string vsdxPath) {
            using ZipArchive archive = ZipFile.OpenRead(vsdxPath);
            using Stream stream = archive.GetEntry("visio/pages/page1.xml")!.Open();
            XDocument pageDoc = XDocument.Load(stream);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement connectorShape = GetConnectorShape(pageDoc, ns);
            XElement geometry = connectorShape.Elements(ns + "Section").First(section => (string?)section.Attribute("N") == "Geometry");
            return geometry.ToString(SaveOptions.DisableFormatting);
        }

        private static XDocument NormalizeTheme(string xml) {
            return XDocument.Parse(xml);
        }
    }
}
