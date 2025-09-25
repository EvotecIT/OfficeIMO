using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioMasterGeometry {
        private static XDocument LoadZipXml(ZipArchive zip, string path) {
            var e = zip.GetEntry(path);
            Assert.NotNull(e);
            using var s = e!.Open();
            return XDocument.Load(s);
        }

        [Fact]
        public void CircleAndTriangleMastersHaveSpecificGeometry() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            var doc = VisioDocument.Create(filePath);
            doc.UseMastersByDefault = true;
            var page = doc.AddPage("Page-1");
            page.Shapes.Add(new VisioShape("C1", 2, 2, 1, 1, "Circle") { NameU = "Circle" });
            page.Shapes.Add(new VisioShape("T1", 4, 2, 1.5, 1, "Triangle") { NameU = "Triangle" });
            doc.Save();

            using var zip = ZipFile.OpenRead(filePath);
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace pr = "http://schemas.openxmlformats.org/package/2006/relationships";

            var masters = LoadZipXml(zip, "visio/masters/masters.xml");
            string GetMasterTarget(string nameU) {
                var m = masters.Root!.Elements(v + "Master").First(e => (string?)e.Attribute("NameU") == nameU);
                string relId = (string)m.Element(v + "Rel")!.Attribute(r + "id")!;
                var rels = LoadZipXml(zip, "visio/masters/_rels/masters.xml.rels");
                var rel = rels.Root!.Elements(pr + "Relationship").First(e => (string?)e.Attribute("Id") == relId);
                return "visio/masters/" + (string)rel.Attribute("Target")!;
            }

            // Circle master -> must contain EllipticalArcTo rows
            var circlePath = GetMasterTarget("Circle");
            var circle = LoadZipXml(zip, circlePath);
            var cGeom = circle.Root!.Element(v + "Shapes")!.Element(v + "Shape")!.Elements(v + "Section").First(e => (string?)e.Attribute("N") == "Geometry");
            Assert.Contains(cGeom.Elements(v + "Row"), row => (string?)row.Attribute("T") == "EllipticalArcTo");

            // Triangle master -> exactly 3 LineTo rows
            var triPath = GetMasterTarget("Triangle");
            var tri = LoadZipXml(zip, triPath);
            var tGeom = tri.Root!.Element(v + "Shapes")!.Element(v + "Shape")!.Elements(v + "Section").First(e => (string?)e.Attribute("N") == "Geometry");
            int lineToCount = tGeom.Elements(v + "Row").Count(e => (string?)e.Attribute("T") == "LineTo");
            Assert.Equal(3, lineToCount);
        }
    }
}

