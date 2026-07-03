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

            VisioShape process = page.AddProcess(1, 1, 2, 1, "Process");
            VisioShape decision = page.AddDecision(4, 1, 2, 2, "Decision");
            VisioShape data = page.AddData(7, 1, 2, 1, "Data");
            VisioShape preparation = page.AddPreparation(10, 1, 2, 1, "Preparation");
            VisioShape manualOperation = page.AddManualOperation(13, 1, 2, 1, "Manual");
            VisioShape offPageReference = page.AddOffPageReference(16, 1, 2, 2, "Off-page");

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
            page.AddProcess(1, 1, 2, 1, "Step");
            page.AddDecision(4, 1, 2, 2, "Branch");
            page.AddData(7, 1, 2, 1, "Input");
            page.AddPreparation(10, 1, 2, 1, "Prepare");
            page.AddManualOperation(13, 1, 2, 1, "Manual");
            page.AddOffPageReference(16, 1, 2, 2, "Jump");
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
            document.RegisterMaster("CustomAudit", new VisioShape("1", 0, 0, 2, 1, string.Empty) { NameU = "CustomAudit" });

            document.AsFluent()
                .Page("Flow", p => p
                    .Process("p1", 1, 1, 2, 1, "Process")
                    .Decision("d1", 4, 1, 2, 2, "Decision")
                    .Data("data1", 7, 1, 2, 1, "Data")
                    .Preparation("prep1", 10, 1, 2, 1, "Prepare")
                    .ManualOperation("manual1", 13, 1, 2, 1, "Manual")
                    .OffPageReference("jump1", 16, 1, 2, 2, "Jump")
                    .Master("audit1", "CustomAudit", 19, 1, 2, 1, "Audit"))
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
            page.AddData(1, 1, 2, 1, "Input");
            page.AddPreparation(4, 1, 2, 1, "Prepare");
            page.AddManualOperation(7, 1, 2, 1, "Manual");
            page.AddOffPageReference(10, 1, 2, 2, "Jump");
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
