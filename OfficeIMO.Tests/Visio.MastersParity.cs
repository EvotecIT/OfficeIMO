using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioMastersParity {
        private static string AssetsPath => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets"));

        [Fact]
        public void BasicMastersMatchTemplate() {
            string template = Path.Combine(AssetsPath, "VisioTemplates", "DrawingWithShapes.vsdx");
            Assert.True(File.Exists(template));

            string target = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            var doc = VisioDocument.Create(target);
            doc.UseMastersByDefault = true;
            doc.UseMastersFromTemplate(template);
            var page = doc.AddPage("Page-1", 29.7, 21, VisioMeasurementUnit.Centimeters);
            page.Shapes.Add(new VisioShape("1") { NameU = "Rectangle", PinX = 2, PinY = 6 });
            page.Shapes.Add(new VisioShape("2") { NameU = "Square", PinX = 4, PinY = 6 });
            page.Shapes.Add(new VisioShape("3") { NameU = "Circle", PinX = 6, PinY = 6 });
            page.Shapes.Add(new VisioShape("4") { NameU = "Ellipse", PinX = 8, PinY = 6 });
            page.Shapes.Add(new VisioShape("5") { NameU = "Diamond", PinX = 10, PinY = 6 });
            page.Shapes.Add(new VisioShape("6") { NameU = "Triangle", PinX = 12, PinY = 6 });
            page.Connectors.Add(new VisioConnector(page.Shapes[0], page.Shapes[1]) { Kind = ConnectorKind.Dynamic });
            doc.Save();

            using var expectedZip = ZipFile.OpenRead(template);
            using var actualZip = ZipFile.OpenRead(target);

            // For each master in actual masters.xml, compare MasterContents and masters.xml <Master> element with template's same NameU
            var ns = XNamespace.Get("http://schemas.microsoft.com/office/visio/2012/main");
            var rNs = XNamespace.Get("http://schemas.openxmlformats.org/package/2006/relationships");

            var mastersEntry = actualZip.GetEntry("visio/masters/masters.xml")!;
            var mastersDoc = XDocument.Load(mastersEntry.Open());
            // Restrict comparison to masters actually present in the template
            var expectedMastersDoc = XDocument.Load(expectedZip.GetEntry("visio/masters/masters.xml")!.Open());
            var templateNames = expectedMastersDoc.Root!.Elements(ns + "Master").Select(e => (string?)e.Attribute("NameU")).Where(n => !string.IsNullOrEmpty(n)).ToHashSet();
            foreach (var m in mastersDoc.Root!.Elements(ns + "Master").Where(e => templateNames.Contains((string?)e.Attribute("NameU") ?? string.Empty))) {
                string nameU = (string?)m.Attribute("NameU") ?? string.Empty;
                string id = (string?)m.Attribute("ID") ?? string.Empty;
                string rid = (string?)m.Element(ns + "Rel")?.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")) ?? string.Empty;
                Assert.False(string.IsNullOrEmpty(nameU));
                Assert.False(string.IsNullOrEmpty(rid));

                // Resolve actual master part path
                var actualRels = XDocument.Load(actualZip.GetEntry("visio/masters/_rels/masters.xml.rels")!.Open());
                var actualRel = actualRels.Root!.Elements(rNs + "Relationship").First(e => (string?)e.Attribute("Id") == rid);
                string actualTarget = (string)actualRel.Attribute("Target")!;
                var actualMasterDoc = XDocument.Load(actualZip.GetEntry("visio/masters/" + actualTarget)!.Open());

                // Find template master with same NameU and load its part
                var expectedMaster = expectedMastersDoc.Root!.Elements(ns + "Master").FirstOrDefault(e => (string?)e.Attribute("NameU") == nameU);
                Assert.NotNull(expectedMaster);
                string expectedRid = (string)expectedMaster!.Element(ns + "Rel")!.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))!;
                var expectedRels = XDocument.Load(expectedZip.GetEntry("visio/masters/_rels/masters.xml.rels")!.Open());
                var expectedRel = expectedRels.Root!.Elements(rNs + "Relationship").First(e => (string?)e.Attribute("Id") == expectedRid);
                string expectedTarget = (string)expectedRel.Attribute("Target")!;
                var expectedMasterDoc = XDocument.Load(expectedZip.GetEntry("visio/masters/" + expectedTarget)!.Open());

                Assert.True(XNode.DeepEquals(Normalize(expectedMasterDoc.Root!), Normalize(actualMasterDoc.Root!)), $"Master {nameU} differed");

                // Compare the <Master> metadata element as well (ignoring differing r:id values)
                XElement normalizeMaster(XElement e, string? rIdReplacement) {
                    var clone = new XElement(e);
                    // Normalize r:id to fixed token
                    var relElement = clone.Element(ns + "Rel");
                    if (relElement != null) {
                        relElement.SetAttributeValue(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"), rIdReplacement ?? "RID");
                    }
                    return clone;
                }
                Assert.True(
                    XNode.DeepEquals(
                        Normalize(normalizeMaster(expectedMaster!, "RID")),
                        Normalize(normalizeMaster(m, "RID"))),
                    $"Master metadata for {nameU} differed");
            }
        }

        private static XElement Normalize(XElement element) {
            return new XElement(element.Name,
                element.Attributes().OrderBy(a => a.Name.ToString()),
                element.Nodes().Select(n => n is XElement e ? Normalize(e) : n));
        }
    }
}
