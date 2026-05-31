using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioStencilProfileTests {
        [Fact]
        public void StencilProfileClassifiesGeneratedMastersAndBasicGeometry() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            document.UseMastersByDefault = true;
            VisioStencilCatalog catalog = VisioStencilCatalog.Create("Profile Catalog", builder => builder
                .Add("profile.cache", "Cache", "Process", "Infrastructure", 1.4, 0.8));
            VisioPage page = document.AddPage("Profile", 8, 5);

            VisioShape cache = page.AddStencilShape(catalog, "profile.cache", "cache", 2, 3, "Cache");
            cache.SetShapeData("Owner", "Platform", "Owner", VisioShapeDataType.String);
            VisioShape note = page.AddRectangle(5, 3, 1.6, 0.7, "Note");
            note.NameU = "Annotation";
            note.SetUserCell(VisioSemanticUserCells.Kind, "Annotation", "STR", prompt: "semantic kind");

            VisioStencilProfile profile = document.CreateStencilProfile();

            Assert.Equal(2, profile.TotalShapes);
            Assert.Equal(1, profile.MasterBackedShapeCount);
            Assert.Equal(1, profile.GeneratedMasterBackedShapeCount);
            Assert.Equal(1, profile.BasicGeometryShapeCount);
            Assert.Equal(new[] { "Owner" }, profile.ShapeDataKeys);
            Assert.Contains("Annotation", profile.SemanticKinds);
            Assert.Equal(new[] { "Profile Catalog" }, profile.StencilCatalogs);
            Assert.Equal(new[] { "Infrastructure" }, profile.StencilCategories);
            Assert.Contains("profile", profile.StencilTags);
            VisioStencilUsageProfile generated = Assert.Single(profile.Usages, usage => usage.Kind == VisioStencilProfileUsageKind.GeneratedMaster);
            Assert.Equal("stencil:profile.cache", generated.Key);
            Assert.Equal("Process", generated.MasterNameU);
            Assert.Equal("profile.cache", generated.StencilId);
            Assert.Equal("Cache", generated.StencilName);
            Assert.Equal("Infrastructure", generated.StencilCategory);
            Assert.Equal("Profile Catalog", generated.StencilCatalogName);
            Assert.Contains("Process", generated.StencilTags);
            Assert.Equal(new[] { "cache" }, generated.ShapeIds);
            VisioStencilUsageProfile geometry = Assert.Single(profile.Usages, usage => usage.Kind == VisioStencilProfileUsageKind.BasicGeometry);
            Assert.Equal("geometry:Annotation", geometry.Key);
            Assert.Equal("Annotation", geometry.ShapeNameU);
            Assert.Equal("Annotation", geometry.SemanticKind);
        }

        [Fact]
        public void StencilProfileCapturesConnectorShapeDataAndStableText() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.UseMastersByDefault = true;
            VisioPage page = document.AddPage("Stable", 8, 5);
            VisioShape source = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "source", 2, 3, "Source");
            VisioShape target = page.AddStencilShape(VisioStencils.Flowchart.Get("decision"), "target", 5, 3, "Target?");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            connector.SetShapeData("Protocol", "HTTPS", "Protocol", VisioShapeDataType.String);
            target.SetShapeData("Criticality", "Tier 0", "Criticality", VisioShapeDataType.String);

            VisioStencilProfile first = document.CreateStencilProfile();
            VisioStencilProfile second = document.CreateStencilProfile();
            string text = first.ToText();

            Assert.Equal(first.ToText(), second.ToText());
            Assert.Equal(2, first.GeneratedMasterBackedShapeCount);
            Assert.Equal(2, first.StencilBackedShapeCount);
            Assert.Equal(2, first.TotalConnectionPoints);
            Assert.Equal(2, first.ConnectionPointShapeCount);
            Assert.Equal(new[] { "Criticality" }, first.ShapeDataKeys);
            Assert.Equal(new[] { "Protocol" }, first.ConnectorShapeDataKeys);
            Assert.Contains("profile.generatedMasterBackedShapeCount=2", text, StringComparison.Ordinal);
            Assert.Contains("profile.stencilBackedShapeCount=2", text, StringComparison.Ordinal);
            Assert.Contains(first.Usages, usage => usage.MasterNameU == "Decision" && usage.StencilId == "flow.decision" && usage.Count == 1 && usage.ConnectionPointCount == 1);
            Assert.Contains(first.Usages, usage => usage.MasterNameU == "Process" && usage.StencilId == "flow.process" && usage.ShapeDataKeys.Count == 0 && usage.ConnectionPointCount == 1);
            Assert.Contains("profile.totalConnectionPoints=2", text, StringComparison.Ordinal);
            Assert.Contains("usage[stencil:flow.process].connectionPointShapeCount=1", text, StringComparison.Ordinal);
            Assert.Contains("profile.stencilCategories=Flowchart", text, StringComparison.Ordinal);
            Assert.Contains("usage[stencil:flow.decision].stencilCatalog=", text, StringComparison.Ordinal);
        }

        [Fact]
        public void StencilProfileRoundTripsPackageBackedMastersAfterLoad() {
            string packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            CreatePackageWithRawGroupMaster(packagePath, "FancyCloud", "Fancy Cloud");
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioStencilCatalog catalog = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                Category = "External",
                IncludeUnsupportedMasters = true,
                IdPrefix = "profile"
            });
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Package", 8, 5);
            page.AddStencilShape(catalog, "fancy-cloud", "source", 2, 3, "Source");
            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilProfile profile = loaded.CreateStencilProfile();

            Assert.Equal(1, profile.TotalShapes);
            Assert.Equal(1, profile.PackageBackedShapeCount);
            Assert.Equal(new[] { "External" }, profile.StencilCategories);
            Assert.Equal(new[] { Path.GetFullPath(packagePath) }, profile.StencilSourcePackagePaths);
            VisioStencilUsageProfile usage = Assert.Single(profile.Usages, item => item.Kind == VisioStencilProfileUsageKind.PackageBackedMaster);
            Assert.Equal("FancyCloud", usage.MasterNameU);
            Assert.Equal("profile.fancycloud", usage.StencilId);
            Assert.Equal("External", usage.StencilCategory);
            Assert.Equal(Path.GetFileNameWithoutExtension(packagePath), usage.StencilCatalogName);
            Assert.Equal(Path.GetFullPath(packagePath), usage.StencilSourcePackagePath);
            Assert.Contains("package", usage.StencilTags);
            Assert.Equal(new[] { "source" }, usage.ShapeIds);

            VisioInspectionMasterSnapshot master = Assert.Single(loaded.CreateInspectionSnapshot().Masters, item => item.NameU == "FancyCloud");
            Assert.True(master.IsPackageBacked);
            Assert.Equal("profile.fancycloud", master.StencilId);
            Assert.Equal(Path.GetFullPath(packagePath), master.StencilSourcePackagePath);
        }

        private static void CreatePackageWithRawGroupMaster(string path, string nameU, string name) {
            const string visioNamespace = "http://schemas.microsoft.com/office/visio/2012/main";
            const string officeRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            const string packageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

            using ZipArchive zip = ZipFile.Open(path, ZipArchiveMode.Create);
            XNamespace ns = visioNamespace;
            XNamespace rel = officeRelationshipNamespace;
            XElement mastersRoot = new(ns + "Masters",
                new XElement(ns + "Master",
                    new XAttribute("ID", "42"),
                    new XAttribute("Name", name),
                    new XAttribute("NameU", nameU),
                    new XElement(ns + "Rel", new XAttribute(rel + "id", "rId1"))));
            WriteZipXml(zip, "visio/masters/masters.xml", new XDocument(mastersRoot));

            XNamespace packageRel = packageRelationshipNamespace;
            XElement relationshipsRoot = new(packageRel + "Relationships",
                new XElement(packageRel + "Relationship",
                    new XAttribute("Id", "rId1"),
                    new XAttribute("Type", officeRelationshipNamespace + "/master"),
                    new XAttribute("Target", "master42.xml")));
            WriteZipXml(zip, "visio/masters/_rels/masters.xml.rels", new XDocument(relationshipsRoot));

            XElement childShape = new(ns + "Shape",
                new XAttribute("ID", "6"),
                new XAttribute("NameU", nameU + ".Icon"),
                new XAttribute("Type", "Shape"),
                new XElement(ns + "Cell", new XAttribute("N", "PinX"), new XAttribute("V", "0.5")),
                new XElement(ns + "Cell", new XAttribute("N", "PinY"), new XAttribute("V", "0.35")),
                new XElement(ns + "Cell", new XAttribute("N", "Width"), new XAttribute("V", "0.6")),
                new XElement(ns + "Cell", new XAttribute("N", "Height"), new XAttribute("V", "0.4")));
            XElement groupShape = new(ns + "Shape",
                new XAttribute("ID", "5"),
                new XAttribute("Name", name),
                new XAttribute("NameU", nameU),
                new XAttribute("Type", "Group"),
                new XElement(ns + "Cell", new XAttribute("N", "PinX"), new XAttribute("V", "0.5")),
                new XElement(ns + "Cell", new XAttribute("N", "PinY"), new XAttribute("V", "0.5")),
                new XElement(ns + "Cell", new XAttribute("N", "Width"), new XAttribute("V", "1")),
                new XElement(ns + "Cell", new XAttribute("N", "Height"), new XAttribute("V", "1")),
                new XElement(ns + "Shapes", childShape));
            XDocument masterDocument = new(new XElement(ns + "MasterContents",
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                new XAttribute(XNamespace.Xmlns + "r", officeRelationshipNamespace),
                new XElement(ns + "Shapes", groupShape)));
            WriteZipXml(zip, "visio/masters/master42.xml", masterDocument);
        }

        private static void WriteZipXml(ZipArchive zip, string path, XDocument document) {
            ZipArchiveEntry entry = zip.CreateEntry(path);
            using Stream stream = entry.Open();
            using StreamWriter writer = new(stream, new UTF8Encoding(false));
            writer.Write(document.Declaration + Environment.NewLine + document.ToString(SaveOptions.DisableFormatting));
        }
    }
}
