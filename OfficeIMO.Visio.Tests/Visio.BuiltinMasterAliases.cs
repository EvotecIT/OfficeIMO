using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioBuiltinMasterAliases {
        [Fact]
        public void PageHelpersCreateSemanticFlowchartNames() {
            VisioPage page = new("Flow");

            VisioShape process = page.AddProcess(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Process");
            VisioShape decision = page.AddDecision(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), "Decision");
            VisioShape data = page.AddData(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Data");
            VisioShape preparation = page.AddPreparation(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(10, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Preparation");
            VisioShape manualOperation = page.AddManualOperation(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(13, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Manual");
            VisioShape offPageReference = page.AddOffPageReference(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(16, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), "Off-page");

            Assert.Equal("Process", process.NameU);
            Assert.Equal("Decision", decision.NameU);
            Assert.Equal("Data", data.NameU);
            Assert.Equal("Preparation", preparation.NameU);
            Assert.Equal("Manual operation", manualOperation.NameU);
            Assert.Equal("Off-page reference", offPageReference.NameU);
        }

        [Fact]
        public void UseMastersByDefaultGeneratesSemanticFlowchartMasters() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            document.UseMastersByDefault = true;
            VisioPage page = document.AddPage("Flow");
            page.AddProcess(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Step");
            page.AddDecision(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), "Branch");
            page.AddData(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Input");
            page.AddPreparation(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(10, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Prepare");
            page.AddManualOperation(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(13, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Manual");
            page.AddOffPageReference(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(16, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), "Jump");
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(new[] { "Process", "Decision", "Data", "Preparation", "Manual operation", "Off-page reference" }, loaded.Pages[0].Shapes.Select(shape => shape.Master?.NameU));

            using ZipArchive zip = ZipFile.OpenRead(filePath);
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XNamespace pr = "http://schemas.openxmlformats.org/package/2006/relationships";

            XDocument masters = LoadZipXml(zip, "visio/masters/masters.xml");
            Assert.Contains(masters.Root!.Elements(v + "Master"), e => (string?)e.Attribute("NameU") == "Process");
            Assert.Contains(masters.Root!.Elements(v + "Master"), e => (string?)e.Attribute("NameU") == "Decision");
            Assert.Contains(masters.Root!.Elements(v + "Master"), e => (string?)e.Attribute("NameU") == "Data");
            Assert.Contains(masters.Root!.Elements(v + "Master"), e => (string?)e.Attribute("NameU") == "Preparation");
            Assert.Contains(masters.Root!.Elements(v + "Master"), e => (string?)e.Attribute("NameU") == "Manual operation");
            Assert.Contains(masters.Root!.Elements(v + "Master"), e => (string?)e.Attribute("NameU") == "Off-page reference");

            string processPart = GetMasterTarget(zip, masters, v, r, pr, "Process");
            string decisionPart = GetMasterTarget(zip, masters, v, r, pr, "Decision");
            string dataPart = GetMasterTarget(zip, masters, v, r, pr, "Data");
            string preparationPart = GetMasterTarget(zip, masters, v, r, pr, "Preparation");
            string manualOperationPart = GetMasterTarget(zip, masters, v, r, pr, "Manual operation");
            string offPageReferencePart = GetMasterTarget(zip, masters, v, r, pr, "Off-page reference");

            XElement processGeometry = LoadZipXml(zip, processPart).Root!.Element(v + "Shapes")!.Element(v + "Shape")!
                .Elements(v + "Section").Single(e => (string?)e.Attribute("N") == "Geometry");
            XElement decisionGeometry = LoadZipXml(zip, decisionPart).Root!.Element(v + "Shapes")!.Element(v + "Shape")!
                .Elements(v + "Section").Single(e => (string?)e.Attribute("N") == "Geometry");
            XElement dataGeometry = LoadZipXml(zip, dataPart).Root!.Element(v + "Shapes")!.Element(v + "Shape")!
                .Elements(v + "Section").Single(e => (string?)e.Attribute("N") == "Geometry");
            XElement preparationGeometry = LoadZipXml(zip, preparationPart).Root!.Element(v + "Shapes")!.Element(v + "Shape")!
                .Elements(v + "Section").Single(e => (string?)e.Attribute("N") == "Geometry");
            XElement manualOperationGeometry = LoadZipXml(zip, manualOperationPart).Root!.Element(v + "Shapes")!.Element(v + "Shape")!
                .Elements(v + "Section").Single(e => (string?)e.Attribute("N") == "Geometry");
            XElement offPageReferenceGeometry = LoadZipXml(zip, offPageReferencePart).Root!.Element(v + "Shapes")!.Element(v + "Shape")!
                .Elements(v + "Section").Single(e => (string?)e.Attribute("N") == "Geometry");

            Assert.Equal(4, processGeometry.Elements(v + "Row").Count(e => (string?)e.Attribute("T") == "LineTo"));
            Assert.Equal(4, decisionGeometry.Elements(v + "Row").Count(e => (string?)e.Attribute("T") == "LineTo"));
            Assert.Equal(4, dataGeometry.Elements(v + "Row").Count(e => (string?)e.Attribute("T") == "LineTo"));
            Assert.Equal(6, preparationGeometry.Elements(v + "Row").Count(e => (string?)e.Attribute("T") == "LineTo"));
            Assert.Equal(4, manualOperationGeometry.Elements(v + "Row").Count(e => (string?)e.Attribute("T") == "LineTo"));
            Assert.Equal(5, offPageReferenceGeometry.Elements(v + "Row").Count(e => (string?)e.Attribute("T") == "LineTo"));
        }

        [Fact]
        public void FluentHelpersCreateSemanticShapesAndRegisteredMasters() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.UseMastersByDefault = true;
            document.RegisterMaster("CustomAudit", new VisioShape("1", OfficeLength.Inches(0), OfficeLength.Inches(0), OfficeLength.Inches(2), OfficeLength.Inches(1), string.Empty));

            document.AsFluent()
                .AddPage("Flow", p => p
                    .AddProcess("p1", OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Process")
                    .AddDecision("d1", OfficeLength.Inches(4), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(2), "Decision")
                    .AddData("data1", OfficeLength.Inches(7), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Data")
                    .AddPreparation("prep1", OfficeLength.Inches(10), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Prepare")
                    .AddManualOperation("manual1", OfficeLength.Inches(13), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Manual")
                    .AddOffPageReference("jump1", OfficeLength.Inches(16), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(2), "Jump")
                    .AddShape("audit1", "CustomAudit", OfficeLength.Inches(19), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "Audit"))
                .End()
                .Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(new[] { "Process", "Decision", "Data", "Preparation", "Manual operation", "Off-page reference", "CustomAudit" }, loaded.Pages[0].Shapes.Select(shape => shape.NameU));
            Assert.Equal("CustomAudit", loaded.Pages[0].Shapes.Last().Master?.NameU);
        }

        [Fact]
        public void StandaloneSemanticShapesUseSemanticGeometry() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Flow");
            page.AddData(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Input");
            page.AddPreparation(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(4, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Prepare");
            page.AddManualOperation(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), "Manual");
            page.AddOffPageReference(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(10, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), "Jump");
            document.Save();

            using ZipArchive zip = ZipFile.OpenRead(filePath);
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument pageDoc = LoadZipXml(zip, "visio/pages/page1.xml");
            XElement[] shapes = pageDoc.Root!.Element(v + "Shapes")!.Elements(v + "Shape").ToArray();

            XElement dataGeometry = shapes[0].Elements(v + "Section").Single(e => (string?)e.Attribute("N") == "Geometry");
            XElement preparationGeometry = shapes[1].Elements(v + "Section").Single(e => (string?)e.Attribute("N") == "Geometry");
            XElement manualOperationGeometry = shapes[2].Elements(v + "Section").Single(e => (string?)e.Attribute("N") == "Geometry");
            XElement offPageReferenceGeometry = shapes[3].Elements(v + "Section").Single(e => (string?)e.Attribute("N") == "Geometry");

            Assert.Equal(4, dataGeometry.Elements(v + "Row").Count(e => (string?)e.Attribute("T") == "LineTo"));
            Assert.Equal(6, preparationGeometry.Elements(v + "Row").Count(e => (string?)e.Attribute("T") == "LineTo"));
            Assert.Equal(4, manualOperationGeometry.Elements(v + "Row").Count(e => (string?)e.Attribute("T") == "LineTo"));
            Assert.Equal(5, offPageReferenceGeometry.Elements(v + "Row").Count(e => (string?)e.Attribute("T") == "LineTo"));
        }

        [Fact]
        public void SupportedBuiltinMastersExposeSemanticAliases() {
            Assert.Contains("Process", VisioDocument.SupportedBuiltinMasters);
            Assert.Contains("Decision", VisioDocument.SupportedBuiltinMasters);
            Assert.Contains("Data", VisioDocument.SupportedBuiltinMasters);
            Assert.Contains("Preparation", VisioDocument.SupportedBuiltinMasters);
            Assert.Contains("Manual operation", VisioDocument.SupportedBuiltinMasters);
            Assert.Contains("Off-page reference", VisioDocument.SupportedBuiltinMasters);
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
