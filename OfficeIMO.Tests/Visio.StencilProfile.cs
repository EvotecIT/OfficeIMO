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
                .AddWithMetadata(
                    "profile.cache",
                    "Cache",
                    "Process",
                    "Infrastructure",
                    1.4,
                    0.8,
                    new[] { "memory", "redis" },
                    new[] { "fast-cache" },
                    new[] { "critical" }));
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
            Assert.Equal(new[] { "memory", "redis" }, profile.StencilKeywords);
            Assert.Equal(new[] { "cache", "fast-cache", "memory", "redis" }, profile.StencilAliases);
            Assert.Contains("profile", profile.StencilTags);
            Assert.Contains("critical", profile.StencilTags);
            Assert.Equal(new[] { "Process" }, profile.StencilIconNameUs);
            VisioStencilFamilyProfile family = Assert.Single(profile.StencilFamilies);
            Assert.Equal("stencil-family:Profile Catalog/Infrastructure", family.Key);
            Assert.Equal("Profile Catalog", family.StencilCatalogName);
            Assert.Equal("Infrastructure", family.StencilCategory);
            Assert.Equal(new[] { "memory", "redis" }, family.StencilKeywords);
            Assert.Equal(new[] { "cache", "fast-cache", "memory", "redis" }, family.StencilAliases);
            Assert.Contains("critical", family.StencilTags);
            Assert.Equal(new[] { "Process" }, family.StencilIconNameUs);
            Assert.Equal(1, family.ShapeCount);
            Assert.Equal(1, family.StencilBackedShapeCount);
            Assert.Equal(1, family.MasterBackedShapeCount);
            Assert.Equal(1, family.GeneratedMasterBackedShapeCount);
            Assert.Equal(0, family.ConnectionPointCount);
            Assert.Equal(new[] { "profile.cache" }, family.StencilIds);
            Assert.Equal(1.4, family.PlacedWidthMinimum);
            Assert.Equal(1.4, family.PlacedWidthMaximum);
            Assert.Equal(0.8, family.PlacedHeightMinimum);
            Assert.Equal(0.8, family.PlacedHeightMaximum);
            Assert.Equal(1.4, family.SourceDefaultWidthMinimum);
            Assert.Equal(1.4, family.SourceDefaultWidthMaximum);
            Assert.Equal(0.8, family.SourceDefaultHeightMinimum);
            Assert.Equal(0.8, family.SourceDefaultHeightMaximum);
            VisioStencilUsageProfile generated = Assert.Single(profile.Usages, usage => usage.Kind == VisioStencilProfileUsageKind.GeneratedMaster);
            Assert.Equal("stencil:profile.cache", generated.Key);
            Assert.Equal("Process", generated.MasterNameU);
            Assert.Equal("profile.cache", generated.StencilId);
            Assert.Equal("Cache", generated.StencilName);
            Assert.Equal("Infrastructure", generated.StencilCategory);
            Assert.Equal("Profile Catalog", generated.StencilCatalogName);
            Assert.Equal(new[] { "memory", "redis" }, generated.StencilKeywords);
            Assert.Equal(new[] { "cache", "fast-cache", "memory", "redis" }, generated.StencilAliases);
            Assert.Contains("Process", generated.StencilTags);
            Assert.Contains("critical", generated.StencilTags);
            Assert.Equal("Process", generated.StencilIconNameU);
            Assert.Equal(1.4, generated.SourceDefaultWidth);
            Assert.Equal(0.8, generated.SourceDefaultHeight);
            Assert.Null(generated.StencilDefaultUnit);
            Assert.Equal(new[] { "cache" }, generated.ShapeIds);
            Assert.Equal(1.4, generated.PlacedWidthMinimum);
            Assert.Equal(1.4, generated.PlacedWidthMaximum);
            Assert.Equal(0.8, generated.PlacedHeightMinimum);
            Assert.Equal(0.8, generated.PlacedHeightMaximum);
            VisioStencilUsageProfile geometry = Assert.Single(profile.Usages, usage => usage.Kind == VisioStencilProfileUsageKind.BasicGeometry);
            Assert.Equal("geometry:Annotation", geometry.Key);
            Assert.Equal("Annotation", geometry.ShapeNameU);
            Assert.Equal("Annotation", geometry.SemanticKind);

            VisioInspectionMasterSnapshot master = Assert.Single(document.CreateInspectionSnapshot().Masters, item => item.NameU == "Process");
            Assert.Null(master.StencilId);
            Assert.Null(master.StencilIconNameU);
            Assert.Null(master.StencilDefaultWidth);
            Assert.Null(master.StencilDefaultHeight);
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
            VisioStencilFamilyProfile flowchartFamily = Assert.Single(first.StencilFamilies);
            Assert.Equal("stencil-family:Flowchart", flowchartFamily.Key);
            Assert.Equal("Flowchart", flowchartFamily.StencilCategory);
            Assert.Equal(2, flowchartFamily.ShapeCount);
            Assert.Equal(2, flowchartFamily.StencilBackedShapeCount);
            Assert.Equal(2, flowchartFamily.MasterBackedShapeCount);
            Assert.Equal(2, flowchartFamily.GeneratedMasterBackedShapeCount);
            Assert.Equal(2, flowchartFamily.ConnectionPointCount);
            Assert.Equal(2, flowchartFamily.ConnectionPointShapeCount);
            Assert.Equal(new[] { "Decision", "Process" }, flowchartFamily.StencilIconNameUs);
            Assert.Equal(2.0, flowchartFamily.PlacedWidthMinimum);
            Assert.Equal(2.4, flowchartFamily.PlacedWidthMaximum);
            Assert.Equal(1.0, flowchartFamily.PlacedHeightMinimum);
            Assert.Equal(1.4, flowchartFamily.PlacedHeightMaximum);
            Assert.Equal(2.0, flowchartFamily.SourceDefaultWidthMinimum);
            Assert.Equal(2.4, flowchartFamily.SourceDefaultWidthMaximum);
            Assert.Equal(1.0, flowchartFamily.SourceDefaultHeightMinimum);
            Assert.Equal(1.4, flowchartFamily.SourceDefaultHeightMaximum);
            Assert.Equal(new[] { "flow.decision", "flow.process" }, flowchartFamily.StencilIds);
            Assert.Equal(new[] { "Criticality" }, first.ShapeDataKeys);
            Assert.Equal(new[] { "Protocol" }, first.ConnectorShapeDataKeys);
            Assert.Contains("profile.generatedMasterBackedShapeCount=2", text, StringComparison.Ordinal);
            Assert.Contains("profile.stencilBackedShapeCount=2", text, StringComparison.Ordinal);
            Assert.Contains("profile.stencilFamilyCount=1", text, StringComparison.Ordinal);
            Assert.Contains("profile.stencilAliases=", text, StringComparison.Ordinal);
            Assert.Contains("profile.stencilIconNameUs=Decision,Process", text, StringComparison.Ordinal);
            Assert.Contains("family[stencil-family:Flowchart].shapeCount=2", text, StringComparison.Ordinal);
            Assert.Contains("family[stencil-family:Flowchart].connectionPointCount=2", text, StringComparison.Ordinal);
            Assert.Contains("family[stencil-family:Flowchart].placedWidthMaximum=2.4", text, StringComparison.Ordinal);
            Assert.Contains("family[stencil-family:Flowchart].sourceDefaultHeightMaximum=1.4", text, StringComparison.Ordinal);
            Assert.Contains(first.Usages, usage => usage.MasterNameU == "Decision" && usage.StencilId == "flow.decision" && usage.Count == 1 && usage.ConnectionPointCount == 1);
            Assert.Contains(first.Usages, usage => usage.MasterNameU == "Process" && usage.StencilId == "flow.process" && usage.ShapeDataKeys.Count == 0 && usage.ConnectionPointCount == 1);
            Assert.Contains("profile.totalConnectionPoints=2", text, StringComparison.Ordinal);
            Assert.Contains("usage[stencil:flow.process].connectionPointShapeCount=1", text, StringComparison.Ordinal);
            Assert.Contains("usage[stencil:flow.process].placedHeightMaximum=1", text, StringComparison.Ordinal);
            Assert.Contains("usage[stencil:flow.process].stencilIconNameU=Process", text, StringComparison.Ordinal);
            Assert.Contains("usage[stencil:flow.process].sourceDefaultWidth=2.4", text, StringComparison.Ordinal);
            Assert.Contains("profile.stencilCategories=Flowchart", text, StringComparison.Ordinal);
            Assert.Contains("usage[stencil:flow.decision].stencilCatalog=", text, StringComparison.Ordinal);
        }

        [Fact]
        public void StencilProfilePreservesGeneratedStencilIdentityAfterLoad() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.UseMastersByDefault = true;
            VisioPage page = document.AddPage("Generated", 8, 5);
            page.AddStencilShape(VisioStencils.Flowchart, "flow.process", "process", 2, 3, "Process");

            VisioStencilProfile beforeSave = document.CreateStencilProfile();
            document.Save();
            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilProfile afterLoad = loaded.CreateStencilProfile();

            Assert.Equal(1, beforeSave.GeneratedMasterBackedShapeCount);
            Assert.Equal(1, afterLoad.GeneratedMasterBackedShapeCount);
            Assert.Equal(0, afterLoad.BasicGeometryShapeCount);
            Assert.Equal(1, afterLoad.StencilBackedShapeCount);
            VisioStencilUsageProfile usage = Assert.Single(afterLoad.Usages, item => item.Kind == VisioStencilProfileUsageKind.GeneratedMaster);
            Assert.Equal("flow.process", usage.StencilId);
            Assert.Equal("Process", usage.MasterNameU);
            Assert.Equal(new[] { "process" }, usage.ShapeIds);
        }

        [Fact]
        public void StencilProfileReportsPagesForRepeatedShapeIds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage first = document.AddPage("First", 8, 5);
            VisioPage second = document.AddPage("Second", 8, 5);
            VisioShape firstShape = first.AddRectangle(2, 3, 1, 0.5, "Cache");
            VisioShape secondShape = second.AddRectangle(2, 3, 1, 0.5, "Cache");
            firstShape.SetUserCell(VisioSemanticUserCells.StencilId, "shared.cache");
            secondShape.SetUserCell(VisioSemanticUserCells.StencilId, "shared.cache");

            Assert.Equal(firstShape.Id, secondShape.Id);
            VisioStencilUsageProfile usage = Assert.Single(document.CreateStencilProfile().Usages);

            Assert.Equal(new[] { "First", "Second" }, usage.PageNames);
        }

        [Fact]
        public void StencilProfilePreservesDistinctSameIdMasters() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioMaster first = new("1", "PackageA", new VisioShape("1", 0, 0, 1, 1, string.Empty) { NameU = "PackageA" }) {
                IsPackageBacked = true,
                StencilId = "pkg.a",
                StencilName = "Package A",
                StencilSourcePackagePath = "a.vssx"
            };
            VisioMaster second = new("1", "PackageB", new VisioShape("1", 0, 0, 1, 1, string.Empty) { NameU = "PackageB" }) {
                IsPackageBacked = true,
                StencilId = "pkg.b",
                StencilName = "Package B",
                StencilSourcePackagePath = "b.vssx"
            };
            document.RegisterMaster(first);
            document.RegisterMaster(second);
            VisioPage page = document.AddPage("Packages", 8, 5);
            page.AddShape("a", "PackageA", 2, 3, 1, 1, "A");
            page.AddShape("b", "PackageB", 5, 3, 1, 1, "B");

            VisioStencilProfile profile = document.CreateStencilProfile();

            Assert.Contains(profile.Usages, usage => usage.MasterNameU == "PackageA" && usage.StencilId == "pkg.a");
            Assert.Contains(profile.Usages, usage => usage.MasterNameU == "PackageB" && usage.StencilId == "pkg.b");
            Assert.Equal(2, profile.Usages.Count(usage => usage.Kind == VisioStencilProfileUsageKind.PackageBackedMaster));
        }

        [Fact]
        public void StencilProfileDisambiguatesDuplicateUsageSnapshotPrefixes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.UseMastersByDefault = true;
            VisioStencilCatalog catalog = VisioStencilCatalog.Create("Mixed", builder => builder
                .AddWithMetadata("shared.cache", "Cache", "Process", "Infrastructure", 1.4, 0.8));
            VisioPage page = document.AddPage("Mixed", 8, 5);
            page.AddStencilShape(catalog, "shared.cache", "master-cache", 2, 3, "Master");
            VisioShape geometry = page.AddRectangle(5, 3, 1.4, 0.8, "Geometry");
            geometry.SetUserCell(VisioSemanticUserCells.StencilId, "shared.cache");

            VisioStencilProfile profile = document.CreateStencilProfile();
            string text = profile.ToText();

            Assert.Equal(2, profile.Usages.Count);
            Assert.Contains("usage[GeneratedMaster:", text, StringComparison.Ordinal);
            Assert.Contains(":stencil:shared.cache].kind=GeneratedMaster", text, StringComparison.Ordinal);
            Assert.Contains("usage[BasicGeometry:", text, StringComparison.Ordinal);
            Assert.Contains(":stencil:shared.cache].kind=BasicGeometry", text, StringComparison.Ordinal);
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
            VisioStencilFamilyProfile family = Assert.Single(profile.StencilFamilies);
            Assert.Equal("stencil-family:" + Path.GetFileNameWithoutExtension(packagePath) + "/External", family.Key);
            Assert.Equal(Path.GetFileNameWithoutExtension(packagePath), family.StencilCatalogName);
            Assert.Equal("External", family.StencilCategory);
            Assert.Equal(1, family.PackageBackedShapeCount);
            Assert.Equal(new[] { Path.GetFullPath(packagePath) }, family.StencilSourcePackagePaths);
            VisioStencilUsageProfile usage = Assert.Single(profile.Usages, item => item.Kind == VisioStencilProfileUsageKind.PackageBackedMaster);
            Assert.Equal("FancyCloud", usage.MasterNameU);
            Assert.Equal("profile.fancy-cloud", usage.StencilId);
            Assert.Equal("External", usage.StencilCategory);
            Assert.Equal(Path.GetFileNameWithoutExtension(packagePath), usage.StencilCatalogName);
            Assert.Equal(Path.GetFullPath(packagePath), usage.StencilSourcePackagePath);
            Assert.Contains("package", usage.StencilTags);
            Assert.Equal("FancyCloud", usage.StencilIconNameU);
            Assert.Equal(1, usage.SourceDefaultWidth);
            Assert.Equal(1, usage.SourceDefaultHeight);
            Assert.Equal("Inches", usage.StencilDefaultUnit);
            Assert.Equal(new[] { "source" }, usage.ShapeIds);

            VisioInspectionMasterSnapshot master = Assert.Single(loaded.CreateInspectionSnapshot().Masters, item => item.NameU == "FancyCloud");
            Assert.True(master.IsPackageBacked);
            Assert.Equal("profile.fancy-cloud", master.StencilId);
            Assert.Equal(Path.GetFullPath(packagePath), master.StencilSourcePackagePath);
            Assert.Equal("FancyCloud", master.StencilIconNameU);
            Assert.Equal(1, master.StencilDefaultWidth);
            Assert.Equal(1, master.StencilDefaultHeight);
            Assert.Equal("Inches", master.StencilDefaultUnit);
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
