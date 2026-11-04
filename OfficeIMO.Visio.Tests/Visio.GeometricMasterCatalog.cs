using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioGeometricMasterCatalog {
        private static string AssetsPath => Path.Combine(RepositoryTestPaths.Find(), "Assets");

        [Fact]
        public void PageHelpersCreateGeometricShapeNames() {
            VisioPage page = new("Shapes");

            VisioShape parallelogram = page.AddParallelogram(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Parallelogram");
            VisioShape hexagon = page.AddHexagon(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Hexagon");
            VisioShape trapezoid = page.AddTrapezoid(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Trapezoid");
            VisioShape pentagon = page.AddPentagon(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(10, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), "Pentagon");

            Assert.Equal("Parallelogram", parallelogram.NameU);
            Assert.Equal("Hexagon", hexagon.NameU);
            Assert.Equal("Trapezoid", trapezoid.NameU);
            Assert.Equal("Pentagon", pentagon.NameU);
        }

        [Fact]
        public void UseMastersByDefaultGeneratesGeometricMasters() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            document.UseMastersByDefault = true;
            VisioPage page = document.AddPage("Shapes");
            page.AddParallelogram(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Parallelogram");
            page.AddHexagon(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Hexagon");
            page.AddTrapezoid(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Trapezoid");
            page.AddPentagon(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(10, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), "Pentagon");
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(new[] { "Parallelogram", "Hexagon", "Trapezoid", "Pentagon" }, loaded.Pages[0].Shapes.Select(shape => shape.Master?.NameU));

            using ZipArchive zip = ZipFile.OpenRead(filePath);
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace pr = "http://schemas.openxmlformats.org/package/2006/relationships";

            XDocument masters = LoadZipXml(zip, "visio/masters/masters.xml");
            string parallelogramPart = GetMasterTarget(zip, masters, v, r, pr, "Parallelogram");
            string hexagonPart = GetMasterTarget(zip, masters, v, r, pr, "Hexagon");
            string trapezoidPart = GetMasterTarget(zip, masters, v, r, pr, "Trapezoid");
            string pentagonPart = GetMasterTarget(zip, masters, v, r, pr, "Pentagon");

            XElement parallelogramGeometry = LoadZipXml(zip, parallelogramPart).Root!.Element(v + "Shapes")!.Element(v + "Shape")!
                .Elements(v + "Section").Single(e => (string?)e.Attribute("N") == "Geometry");
            XElement hexagonGeometry = LoadZipXml(zip, hexagonPart).Root!.Element(v + "Shapes")!.Element(v + "Shape")!
                .Elements(v + "Section").Single(e => (string?)e.Attribute("N") == "Geometry");
            XElement trapezoidGeometry = LoadZipXml(zip, trapezoidPart).Root!.Element(v + "Shapes")!.Element(v + "Shape")!
                .Elements(v + "Section").Single(e => (string?)e.Attribute("N") == "Geometry");
            XElement pentagonGeometry = LoadZipXml(zip, pentagonPart).Root!.Element(v + "Shapes")!.Element(v + "Shape")!
                .Elements(v + "Section").Single(e => (string?)e.Attribute("N") == "Geometry");

            Assert.Equal(4, parallelogramGeometry.Elements(v + "Row").Count(e => (string?)e.Attribute("T") == "LineTo"));
            Assert.Equal(6, hexagonGeometry.Elements(v + "Row").Count(e => (string?)e.Attribute("T") == "LineTo"));
            Assert.Equal(4, trapezoidGeometry.Elements(v + "Row").Count(e => (string?)e.Attribute("T") == "LineTo"));
            Assert.Equal(5, pentagonGeometry.Elements(v + "Row").Count(e => (string?)e.Attribute("T") == "LineTo"));
        }

        [Fact]
        public void ImportMastersRecognizesTemplateGeometryNames() {
            string template = Path.Combine(AssetsPath, "VisioTemplates", "DrawingWithLotsOfShapresAndArrows.vsdx");
            Assert.True(File.Exists(template));

            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);

            IReadOnlyList<VisioMaster> imported = document.ImportMastersAndGet(template, new[] { "Parallelogram", "Hexagon", "Trapezoid", "Pentagon" });

            Assert.Equal(
                new[] { "Parallelogram", "Hexagon", "Trapezoid", "Pentagon" }.OrderBy(name => name),
                imported.Select(master => master.NameU).OrderBy(name => name));

            VisioPage page = document.AddPage("Imported");
            page.AddShape("shape-1", "Parallelogram", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "P");
            page.AddShape("shape-2", "Hexagon", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "H");
            page.AddShape("shape-3", "Trapezoid", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "T");
            page.AddShape("shape-4", "Pentagon", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(10, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), "Pg");
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(new[] { "Parallelogram", "Hexagon", "Trapezoid", "Pentagon" }, loaded.Pages[0].Shapes.Select(shape => shape.Master?.NameU));
        }

        [Fact]
        public void FluentHelpersCreateGeometricShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.UseMastersByDefault = true;

            document.AsFluent()
                .AddPage("Shapes", p => p
                    .AddParallelogram("para", OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Parallelogram")
                    .AddHexagon("hex", OfficeLength.Inches(4), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Hexagon")
                    .AddTrapezoid("trap", OfficeLength.Inches(7), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Trapezoid")
                    .AddPentagon("pent", OfficeLength.Inches(10), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(2), "Pentagon"))
                .End()
                .Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(new[] { "Parallelogram", "Hexagon", "Trapezoid", "Pentagon" }, loaded.Pages[0].Shapes.Select(shape => shape.NameU));
        }

        [Fact]
        public void SupportedBuiltinMastersExposeGeometricCatalog() {
            Assert.Contains("Parallelogram", VisioDocument.SupportedBuiltinMasters);
            Assert.Contains("Hexagon", VisioDocument.SupportedBuiltinMasters);
            Assert.Contains("Trapezoid", VisioDocument.SupportedBuiltinMasters);
            Assert.Contains("Pentagon", VisioDocument.SupportedBuiltinMasters);
        }

        private static XDocument LoadZipXml(ZipArchive zip, string path) {
            ZipArchiveEntry? entry = zip.GetEntry(path);
            Assert.NotNull(entry);
            using Stream stream = entry!.Open();
            return XDocument.Load(stream);
        }

        private static string GetMasterTarget(ZipArchive zip, XDocument masters, XNamespace v, XNamespace r, XNamespace pr, string nameU) {
            XElement master = masters.Root!.Elements(v + "Master").First(e => (string?)e.Attribute("NameU") == nameU);
            string relId = (string)master.Element(v + "Rel")!.Attribute(r + "id")!;
            XDocument rels = LoadZipXml(zip, "visio/masters/_rels/masters.xml.rels");
            XElement rel = rels.Root!.Elements(pr + "Relationship").First(e => (string?)e.Attribute("Id") == relId);
            return "visio/masters/" + (string)rel.Attribute("Target")!;
        }
    }
}
