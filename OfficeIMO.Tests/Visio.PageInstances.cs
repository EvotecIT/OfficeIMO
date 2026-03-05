using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioPageInstances {
        private static string AssetsPath => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets"));

        [Fact]
        public void ShapesReferenceMastersWithMinimalDeltas() {
            string template = Path.Combine(AssetsPath, "VisioTemplates", "DrawingWithShapes.vsdx");
            string target = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            var doc = VisioDocument.Create(target);
            doc.UseMastersByDefault = true;
            doc.UseMastersFromTemplate(template);
            var page = doc.AddPage("Page-1", 29.7, 21, VisioMeasurementUnit.Centimeters);
            page.Shapes.Add(new VisioShape("R1") { NameU = "Rectangle", PinX = 2, PinY = 6 });
            page.Shapes.Add(new VisioShape("S1") { NameU = "Square", PinX = 4, PinY = 6, Width = 1.2, Height = 1.2 });
            page.Shapes.Add(new VisioShape("C1") { NameU = "Circle", PinX = 6, PinY = 6, Width = 1.2, Height = 1.2 });
            page.Shapes.Add(new VisioShape("D1") { NameU = "Diamond", PinX = 8, PinY = 6 });
            page.Connectors.Add(new VisioConnector(page.Shapes[0], page.Shapes[1]) { Kind = ConnectorKind.Dynamic });
            doc.Save();

            using var zip = ZipFile.OpenRead(target);
            var pageEntry = zip.GetEntry("visio/pages/page1.xml")!;
            var pageDoc = XDocument.Load(pageEntry.Open());
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            var shapes = pageDoc.Root!.Element(ns + "Shapes")!.Elements(ns + "Shape").ToList();
            // First four are 2D shapes
            foreach (var s in shapes.Take(4)) {
                Assert.NotNull(s.Attribute("Master"));
                // Minimal delta: allow PinX/PinY, maybe Width/Height if size differs,
                // and a single reserved Prop row when the public shape id is non-numeric.
                var sections = s.Elements(ns + "Section").ToList();
                Assert.DoesNotContain(sections, section => (string?)section.Attribute("N") == "Geometry");
                foreach (var section in sections) {
                    Assert.Equal("Prop", (string?)section.Attribute("N"));
                    var rows = section.Elements(ns + "Row").ToList();
                    Assert.Single(rows);
                    Assert.Equal("OfficeIMOOriginalId", (string?)rows[0].Attribute("N"));
                }
            }
            // Last ones are connectors; dynamic has Master, others don't
            var dyn = shapes.Last();
            Assert.Equal("Dynamic connector", (string?)dyn.Attribute("NameU"));
            Assert.NotNull(dyn.Attribute("Master"));
            Assert.NotNull(dyn.Element(ns + "XForm1D"));

            var reloaded = VisioDocument.Load(target);
            Assert.Equal(new[] { "R1", "S1", "C1", "D1" }, reloaded.Pages[0].Shapes.Select(shape => shape.Id));
            Assert.Single(reloaded.Pages[0].Connectors);
        }
    }
}

