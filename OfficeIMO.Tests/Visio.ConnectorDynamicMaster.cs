using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using SixLabors.ImageSharp;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioConnectorDynamicMasterTests {
        [Fact]
        public void DynamicConnectorUsesMasterAndPageRels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            var doc = VisioDocument.Create(filePath);
            doc.UseMastersByDefault = true;
            var page = doc.AddPage("Page-1", 8.5, 11);
            var s1 = new VisioShape("1", 1, 1, 1.5, 1, "A") { NameU = "Rectangle", FillColor = Color.LightBlue };
            var s2 = new VisioShape("2", 4, 1, 1.5, 1, "B") { NameU = "Rectangle", FillColor = Color.LightGreen };
            page.Shapes.Add(s1);
            page.Shapes.Add(s2);
            // Kind defaults to Dynamic; leave it.
            var c = new VisioConnector(s1, s2);
            page.Connectors.Add(c);
            doc.Save();

            using ZipArchive z = ZipFile.OpenRead(filePath);
            var pageXml = LoadXml(z, "visio/pages/page1.xml");
            var relsXml = LoadXml(z, "visio/pages/_rels/page1.xml.rels");
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            var conn = pageXml.Root!.Element(v + "Shapes")!.Elements(v + "Shape").First(e => (string?)e.Attribute("ID") == c.Id);
            Assert.Equal("Dynamic connector", (string?)conn.Attribute("NameU"));
            Assert.NotNull(conn.Attribute("Master"));
            // Ensure at least one master relationship exists for page
            XNamespace pr = "http://schemas.openxmlformats.org/package/2006/relationships";
            Assert.Contains(relsXml.Root!.Elements(pr + "Relationship"), e => ((string?)e.Attribute("Type"))?.EndsWith("/master") == true);
        }

        private static XDocument LoadXml(ZipArchive zip, string path) {
            var e = zip.GetEntry(path);
            Assert.NotNull(e);
            using var s = e!.Open();
            return XDocument.Load(s);
        }
    }
}
