using System;
using System.IO;
using System.Globalization;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioAssetSamples {
        [Fact]
        public void EmptyDocument_BasicStructure_IsValid() {
            string target = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(target);
            VisioPage page = document.AddPage("Page-1", 21, 29.7, VisioMeasurementUnit.Centimeters);
            page.ViewCenterX = 4.1233127451916;
            page.ViewCenterY = 5.8492688900245;
            document.Save();

            using ZipArchive actual = ZipFile.OpenRead(target);
            // pages.xml
            var pagesDoc = LoadXml(actual, "visio/pages/pages.xml");
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            var pageEl = pagesDoc.Root!.Element(v + "Page");
            Assert.NotNull(pageEl);
            Assert.Equal("Page-1", (string?)pageEl!.Attribute("NameU"));

            // page1.xml
            var page1Doc = LoadXml(actual, "visio/pages/page1.xml");
            var shapes = page1Doc.Root!.Element(v + "Shapes");
            // No shapes for empty doc
            Assert.True(shapes == null || !shapes.Elements().Any());
        }

        [Fact]
        public void RectangleDocument_HasRectangleShape_WithPinCoordinates() {
            string target = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(target);
            document.UseMastersByDefault = true;
            VisioPage page = document.AddPage("Page-1", 29.7, 21, VisioMeasurementUnit.Centimeters);
            page.ViewCenterX = 5.8424184863857;
            page.ViewCenterY = 4.133858091015;
            page.Shapes.Add(new VisioShape("1") {
                NameU = "Rectangle",
                PinX = 2.047244040636296,
                PinY = 6.73228320203895
            });
            document.Save();

            using ZipArchive actual = ZipFile.OpenRead(target);
            var page1Doc = LoadXml(actual, "visio/pages/page1.xml");
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            var shape = page1Doc.Root!
                .Element(v + "Shapes")?
                .Elements(v + "Shape")
                .FirstOrDefault();
            Assert.NotNull(shape);
            Assert.Equal("Rectangle", (string?)shape!.Attribute("NameU"));

            double pinX = GetCellValue(shape, "PinX");
            double pinY = GetCellValue(shape, "PinY");
            AssertApproximately(2.047244040636296, pinX);
            AssertApproximately(6.73228320203895, pinY);
        }

        private static XDocument LoadXml(ZipArchive zip, string entryPath) {
            var entry = zip.GetEntry(entryPath);
            Assert.NotNull(entry);
            using var s = entry!.Open();
            return XDocument.Load(s);
        }

        private static double GetCellValue(XElement shape, string name) {
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            var cell = shape.Elements(v + "Cell").FirstOrDefault(c => (string?)c.Attribute("N") == name);
            Assert.NotNull(cell);
            return double.Parse((string?)cell!.Attribute("V") ?? "0", NumberStyles.Float, CultureInfo.InvariantCulture);
        }

        private static void AssertApproximately(double expected, double actual, double tol = 1e-9) {
            Assert.InRange(Math.Abs(expected - actual), 0, tol);
        }
    }
}

