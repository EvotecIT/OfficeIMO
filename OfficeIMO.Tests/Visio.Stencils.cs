using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioStencilsTests {
        [Fact]
        public void BuiltInCatalogsExposeSearchableStencilDefinitions() {
            Assert.True(VisioStencils.Flowchart.TryGet("flow.process", out VisioStencilShape? process));
            Assert.NotNull(process);
            Assert.Equal("Process", process!.MasterNameU);
            Assert.Equal("Process", process.IconNameU);
            Assert.Contains("process", process.Aliases);
            Assert.Contains("Flowchart", process.Tags);

            VisioStencilShape byKeyword = VisioStencils.Flowchart.Get("branch");
            Assert.Equal("Decision", byKeyword.MasterNameU);

            Assert.Contains(VisioStencils.All.Shapes, shape => shape.Id == "basic.rectangle");
            Assert.Contains(VisioStencils.All.Shapes, shape => shape.Id == "flow.off-page-reference");
            Assert.Contains("Network", VisioStencils.All.Categories);
        }

        [Fact]
        public void CatalogCanSearchByCategoryAliasKeywordAndTag() {
            IReadOnlyList<VisioStencilShape> categoryMatches = VisioStencils.All.Search("Architecture");
            IReadOnlyList<VisioStencilShape> aliasMatches = VisioStencils.Network.Search("access-point");
            IReadOnlyList<VisioStencilShape> tagMatches = VisioStencils.All.Search("net");
            IReadOnlyList<VisioStencilShape> categoryOnly = VisioStencils.All.InCategory("Timeline");

            Assert.Contains(categoryMatches, shape => shape.Id == "arch.service");
            Assert.Equal("Wireless AP", Assert.Single(aliasMatches).Name);
            Assert.Contains(tagMatches, shape => shape.Id == "net.switch");
            Assert.Contains(tagMatches, shape => shape.Id == "net.firewall");
            Assert.All(categoryOnly, shape => Assert.Equal("Timeline", shape.Category));
            Assert.Contains(categoryOnly, shape => shape.Id == "time.milestone");
        }

        [Fact]
        public void PageCanPlaceStencilShapeWithGeneratedMaster() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Stencil Page", 29.7, 21, VisioMeasurementUnit.Centimeters);

            VisioShape process = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "process-1", 5, 15, "Review request");
            VisioShape decision = page.AddStencilShape(VisioStencils.Flowchart, "branch", "decision-1", 13, 15, "Approved?");
            page.AddConnector(process, decision, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);

            Assert.Equal("Process", process.NameU);
            Assert.Equal("Process", process.MasterNameU);
            Assert.Equal("Decision", decision.NameU);
            Assert.Equal("Decision", decision.MasterNameU);
            Assert.Equal(5.0.ToInches(VisioMeasurementUnit.Centimeters), process.PinX, 6);

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(2, loaded.Pages[0].Shapes.Count);
            Assert.Single(loaded.Pages[0].Connectors);
        }

        [Fact]
        public void FluentCanPlaceStencilShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            document.AsFluent()
                .Page("Stencil Page", page => page
                    .Stencil("input", VisioStencils.BlockDiagram.Get("block.storage"), 3, 4, "Storage")
                    .Stencil("worker", VisioStencils.BlockDiagram, "component", 6, 4, "Worker")
                    .Stencil("router", "net.router", 9, 4, "Router")
                    .Connect("input", "worker", VisioSide.Right, VisioSide.Left))
                .End();
            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(new[] { "input", "worker", "router" }, loaded.Pages[0].Shapes.Select(shape => shape.Id));
            Assert.Equal("Decision", loaded.Pages[0].Shapes[2].MasterNameU);
            Assert.Single(loaded.Pages[0].Connectors);
        }

        [Fact]
        public void PageCanPlaceBuiltInStencilFromAllCatalogByString() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("All Catalog");

            VisioShape switchShape = page.AddStencilShape("net.switch", "switch", 2, 4, "Switch");
            VisioShape milestone = page.AddStencilShape("time.milestone", "milestone", 5, 4, 0.4, 0.4, "M1");
            page.AddConnector(switchShape, milestone, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            document.Save();

            Assert.Equal("Rectangle", switchShape.MasterNameU);
            Assert.Equal("Diamond", milestone.MasterNameU);
            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(new[] { "switch", "milestone" }, loaded.Pages[0].Shapes.Select(shape => shape.Id));
        }

        [Fact]
        public void CustomStencilCatalogBuilderCreatesSearchablePaletteAndPlaceableShapes() {
            VisioStencilCatalog catalog = VisioStencilCatalog.Create("Custom Infrastructure", builder => builder
                .Add("custom.cache", "Cache", "Process", "Infrastructure", 1.8, 0.9, "redis", "memory-store")
                .AddWithMetadata(
                    "custom.archive",
                    "Object Archive",
                    "Data",
                    "Infrastructure",
                    1.8,
                    0.9,
                    keywords: new[] { "blob" },
                    aliases: new[] { "object-store" },
                    tags: new[] { "cloud", "storage" },
                    iconNameU: "Data"));

            Assert.Equal(new[] { "Infrastructure" }, catalog.Categories);
            Assert.Equal("Cache", Assert.Single(catalog.Search("redis")).Name);
            Assert.Equal("Object Archive", Assert.Single(catalog.Search("object-store")).Name);
            Assert.Contains(catalog.Search("cloud"), shape => shape.Id == "custom.archive");
            Assert.Contains("cache", catalog.Get("custom.cache").Aliases);
            Assert.Equal("Data", catalog.Get("custom.archive").IconNameU);

            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Custom Stencils");
            VisioShape cache = page.AddStencilShape(catalog, "redis", "cache", 2, 4, "Cache");
            VisioShape archive = page.AddStencilShape(catalog, "object-store", "archive", 5, 4, "Archive");
            page.AddConnector(cache, archive, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(new[] { "cache", "archive" }, loaded.Pages[0].Shapes.Select(shape => shape.Id));
            Assert.Single(loaded.Pages[0].Connectors);
        }

        [Fact]
        public void CustomStencilCatalogBuilderRejectsDuplicateIds() {
            VisioStencilCatalogBuilder builder = new("Custom");
            builder.Add("custom.node", "Node", "Process", "Custom", 1, 1);

            ArgumentException exception = Assert.Throws<ArgumentException>(() => builder.Add("custom.node", "Node 2", "Process", "Custom", 1, 1));

            Assert.Contains("custom.node", exception.Message);
        }

        [Fact]
        public void StencilCatalogManifestRoundTripsMetadataAndLoadedCatalogCanPlaceShapes() {
            string manifestPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".officeimo-visio-stencils.xml");
            VisioStencilCatalog source = VisioStencilCatalog.Create("Reusable Infrastructure", builder => builder
                .Add("infra.api", "API", "Process", "Infrastructure", 1.8, 0.9, "service", "http")
                .AddWithMetadata(
                    "infra.queue",
                    "Queue",
                    "Data",
                    "Infrastructure",
                    1.8,
                    0.9,
                    keywords: new[] { "bus" },
                    aliases: new[] { "message-queue" },
                    tags: new[] { "async", "cloud" },
                    iconNameU: "Data"));

            source.Save(manifestPath);
            VisioStencilCatalog loadedCatalog = VisioStencilCatalog.Load(manifestPath);

            Assert.Equal(source.Name, loadedCatalog.Name);
            Assert.Equal(source.Shapes.Count, loadedCatalog.Shapes.Count);
            Assert.Equal(new[] { "Infrastructure" }, loadedCatalog.Categories);
            Assert.Equal("API", loadedCatalog.Get("http").Name);
            Assert.Equal("Queue", Assert.Single(loadedCatalog.Search("message-queue")).Name);
            Assert.Contains("cloud", loadedCatalog.Get("infra.queue").Tags);
            Assert.Equal("Data", loadedCatalog.Get("infra.queue").IconNameU);

            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Reusable Catalog");
            VisioShape api = page.AddStencilShape(loadedCatalog, "http", "api", 2, 4, "API");
            VisioShape queue = page.AddStencilShape(loadedCatalog, "message-queue", "queue", 5, 4, "Queue");
            page.AddConnector(api, queue, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loadedDocument = VisioDocument.Load(filePath);
            Assert.Equal(new[] { "api", "queue" }, loadedDocument.Pages[0].Shapes.Select(shape => shape.Id));
            Assert.Single(loadedDocument.Pages[0].Connectors);
        }

        [Fact]
        public void StencilCatalogManifestRejectsUnsupportedVersion() {
            XNamespace ns = "urn:officeimo:visio:stencils";
            XDocument manifest = new(new XElement(ns + "StencilCatalog",
                new XAttribute("Version", "99"),
                new XAttribute("Name", "Future")));

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => VisioStencilCatalogManifest.FromXml(manifest));

            Assert.Contains("version", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void PackageStencilCatalogLoadsSupportedMastersFromVssxWithoutRuntimeTemplateDependency() {
            string packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            CreatePackageWithMasters(packagePath, "Rectangle", "FancyCloud", "Decision");

            VisioStencilCatalog catalog = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                CatalogName = "Imported Network",
                Category = "Imported",
                IdPrefix = "imported"
            });

            Assert.Equal("Imported Network", catalog.Name);
            Assert.Equal(new[] { "Imported" }, catalog.Categories);
            Assert.Equal(new[] { "Decision", "Rectangle" }, catalog.Shapes.Select(shape => shape.MasterNameU).OrderBy(name => name).ToArray());
            Assert.DoesNotContain(catalog.Shapes, shape => shape.MasterNameU == "FancyCloud");
            Assert.Contains(catalog.Search("vssx"), shape => shape.Id == "imported.rectangle");
            Assert.Equal("Rectangle", catalog.Get("rectangle").MasterNameU);

            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Package Catalog");
            VisioShape source = page.AddStencilShape(catalog, "rectangle", "source", 2, 4, "Source");
            VisioShape decision = page.AddStencilShape(catalog, "decision", "decision", 5, 4, "Decision");
            page.AddConnector(source, decision, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(new[] { "Rectangle", "Decision" }, loaded.Pages[0].Shapes.Select(shape => shape.MasterNameU));
            Assert.Single(loaded.Pages[0].Connectors);
        }

        [Fact]
        public void PackageStencilCatalogLoadsFromVstxAndCanIncludeUnsupportedMasters() {
            string packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vstx");
            CreatePackageWithMasters(packagePath, "Rectangle", "FancyCloud");

            VisioStencilCatalog supportedOnly = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                MasterNames = new[] { "FancyCloud", "Rectangle" }
            });
            VisioStencilCatalog withGeneric = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                IncludeUnsupportedMasters = true,
                Category = "Template Masters",
                MasterNames = new[] { "FancyCloud" }
            });

            Assert.Single(supportedOnly.Shapes);
            Assert.Equal("Rectangle", supportedOnly.Shapes[0].MasterNameU);
            Assert.Single(withGeneric.Shapes);
            Assert.Equal("FancyCloud", withGeneric.Shapes[0].MasterNameU);
            Assert.Contains("generic", withGeneric.Shapes[0].Tags);
            Assert.Contains(withGeneric.Search("vstx"), shape => shape.MasterNameU == "FancyCloud");
        }

        [Fact]
        public void PackageStencilCatalogFiltersByVisibleAndPackageMasterIdentities() {
            string packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            CreatePackageWithMasterMetadata(
                packagePath,
                ("Rectangle", "Basic Box"),
                ("Decision", "Branch Point"),
                ("Manual operation", "Manual Step"));

            VisioStencilCatalog catalog = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                IdPrefix = "learned",
                MasterNames = new[] { "Basic Box", "2", "rId3", "manual-operation" }
            });

            Assert.Equal(new[] { "Decision", "Manual operation", "Rectangle" }, catalog.Shapes.Select(shape => shape.MasterNameU).OrderBy(name => name).ToArray());
            Assert.Equal("Rectangle", catalog.Get("basic-box").MasterNameU);
            Assert.Equal("Decision", catalog.Get("2").MasterNameU);
            Assert.Equal("Manual operation", catalog.Get("rId3").MasterNameU);
            Assert.Equal("Manual operation", catalog.Get("manual-operation").MasterNameU);
        }

        [Fact]
        public void PackageStencilCatalogPlacesUnsupportedMastersAsGeneratedPlaceholders() {
            string packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vstx");
            CreatePackageWithMasters(packagePath, "FancyCloud");
            VisioStencilCatalog catalog = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                IncludeUnsupportedMasters = true,
                Category = "Template Masters"
            });

            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Generic Placeholder");
            VisioShape cloud = page.AddStencilShape(catalog, "fancycloud", "cloud", 2, 4, "Cloud");
            document.Save();

            Assert.Equal("FancyCloud", cloud.MasterNameU);
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal("FancyCloud", loaded.Pages[0].Shapes[0].MasterNameU);
        }

        [Fact]
        public void CatalogThrowsForUnknownStencilShape() {
            KeyNotFoundException exception = Assert.Throws<KeyNotFoundException>(() => VisioStencils.BasicShapes.Get("not-here"));

            Assert.Contains("not-here", exception.Message);
        }

        private static void CreatePackageWithMasters(string path, params string[] masterNames) {
            CreatePackageWithMasterMetadata(path, masterNames.Select(name => (name, (string?)null)).ToArray());
        }

        private static void CreatePackageWithMasterMetadata(string path, params (string NameU, string? Name)[] masters) {
            const string visioNamespace = "http://schemas.microsoft.com/office/visio/2012/main";
            const string relationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            using ZipArchive zip = ZipFile.Open(path, ZipArchiveMode.Create);
            ZipArchiveEntry mastersEntry = zip.CreateEntry("visio/masters/masters.xml");
            using Stream stream = mastersEntry.Open();
            using StreamWriter writer = new(stream, new UTF8Encoding(false));

            XNamespace ns = visioNamespace;
            XNamespace rel = relationshipNamespace;
            XElement root = new(ns + "Masters",
                masters.Select((master, index) =>
                    new XElement(ns + "Master",
                        new XAttribute("ID", index + 1),
                        new XAttribute("Name", master.Name ?? master.NameU),
                        new XAttribute("NameU", master.NameU),
                        new XElement(ns + "Rel", new XAttribute(rel + "id", $"rId{index + 1}")))));

            XDocument document = new(root);
            writer.Write(document.Declaration + Environment.NewLine + document.ToString(SaveOptions.DisableFormatting));
        }
    }
}
