using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
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
            Assert.True(VisioStencils.All.TryFindBest(new[] { "missing", "access-point" }, out VisioStencilShape? best));
            Assert.Equal("Wireless AP", best!.Name);
            Assert.Equal("Storage", VisioStencils.All.FindBest("not-present", "data-store").Name);
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
        public void GeneratedStencilMasterInstancesUseRendererFriendlyShapeReferences() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.UseMastersByDefault = true;
            VisioPage page = document.AddPage("Stencil Page");

            VisioShape service = page.AddStencilShape(VisioStencils.Architecture.Get("service"), "service", 2, 5, "Service");
            service.FillColor = OfficeColor.FromRgb(31, 119, 180);
            service.LineColor = OfficeColor.FromRgb(10, 70, 120);
            VisioShape database = page.AddStencilShape(VisioStencils.Architecture.Get("database"), "database", 5, 5, "Database");
            database.FillColor = OfficeColor.FromRgb(46, 160, 67);
            database.LineColor = OfficeColor.FromRgb(18, 92, 43);

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));

            using ZipArchive zip = ZipFile.OpenRead(filePath);
            XNamespace v = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument pageXml = XDocument.Load(zip.GetEntry("visio/pages/page1.xml")!.Open());
            XElement[] pageShapes = pageXml.Root!.Element(v + "Shapes")!.Elements(v + "Shape").ToArray();
            XElement serviceShape = pageShapes.Single(shape => GetOriginalId(shape, v) == "service");
            XElement databaseShape = pageShapes.Single(shape => GetOriginalId(shape, v) == "database");

            Assert.NotNull(serviceShape.Attribute("Master"));
            Assert.NotNull(databaseShape.Attribute("Master"));
            Assert.Null(serviceShape.Attribute("MasterShape"));
            Assert.Null(databaseShape.Attribute("MasterShape"));
            Assert.Equal("#1F77B4", GetCellValue(serviceShape, v, "FillForegnd"));
            Assert.Equal("#0A4678", GetCellValue(serviceShape, v, "LineColor"));
            Assert.Equal("#2EA043", GetCellValue(databaseShape, v, "FillForegnd"));
            Assert.Equal("#125C2B", GetCellValue(databaseShape, v, "LineColor"));
            Assert.DoesNotContain(serviceShape.Elements(v + "Section"), section => (string?)section.Attribute("N") == "Geometry");

            XDocument mastersXml = XDocument.Load(zip.GetEntry("visio/masters/masters.xml")!.Open());
            Assert.Contains(mastersXml.Root!.Elements(v + "Master"), master => (string?)master.Attribute("NameU") == "Process");
            Assert.Contains(mastersXml.Root!.Elements(v + "Master"), master => (string?)master.Attribute("NameU") == "Data");

            static string? GetCellValue(XElement shape, XNamespace ns, string name) {
                return (string?)shape.Elements(ns + "Cell")
                    .FirstOrDefault(cell => (string?)cell.Attribute("N") == name)
                    ?.Attribute("V");
            }

            static string? GetOriginalId(XElement shape, XNamespace ns) {
                return (string?)shape.Elements(ns + "Section")
                    .Where(section => (string?)section.Attribute("N") == "Prop")
                    .Elements(ns + "Row")
                    .FirstOrDefault(row => (string?)row.Attribute("N") == "OfficeIMOOriginalId")
                    ?.Elements(ns + "Cell")
                    .FirstOrDefault(cell => (string?)cell.Attribute("N") == "Value")
                    ?.Attribute("V");
            }
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
        public void PageCanRenderStencilCatalogGallery() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioStencilCatalog catalog = VisioStencilCatalog.Create("Gallery Catalog", builder => builder
                .Add("gallery.api", "API", "Process", "Integration", 1.8, 0.9)
                .Add("gallery.queue", "Queue", "Data", "Integration", 1.4, 0.8)
                .Add("gallery.worker", "Worker", "Rectangle", "Compute", 1.6, 0.9));
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Gallery", 5, 4);

            IReadOnlyList<VisioShape> placed = page.AddStencilGallery(catalog, new VisioStencilGalleryOptions {
                IdPrefix = "gallery",
                Columns = 2,
                MaxShapes = 3,
                Title = "Reusable palette",
                AutoResizePage = true
            });

            Assert.Equal(3, placed.Count);
            Assert.True(page.Width > 5);
            Assert.Contains(page.Shapes, shape => shape.Id == "gallery-title" && shape.Text == "Reusable palette");
            Assert.Contains(page.Shapes, shape => shape.Id == "gallery-0-name" && shape.Text == "API");
            Assert.Contains(page.Shapes, shape => shape.Id == "gallery-1-category" && shape.Text == "Integration");
            Assert.Equal("Process", page.Shapes.Single(shape => shape.Id == "gallery-0-shape").MasterNameU);

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(13, loaded.Pages[0].Shapes.Count);
        }

        [Fact]
        public void StencilCatalogGalleryReservesIdsAndKeepsUnitlessStencilSizesInInches() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioStencilCatalog catalog = VisioStencilCatalog.Create("Metric Gallery Catalog", builder => builder
                .Add("gallery.api", "API", "Process", "Integration", 1.8, 0.9));
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Metric Gallery", 29.7, 21, VisioMeasurementUnit.Centimeters);
            VisioStencilGalleryOptions options = new() {
                IdPrefix = "gallery",
                Columns = 1,
                MaxShapes = 1,
                Title = "Reusable palette",
                AutoResizePage = false,
                IconMaxWidth = 1D,
                IconMaxHeight = 0.8D
            };

            page.AddStencilGallery(catalog, options);
            page.AddStencilGallery(catalog, options);

            Assert.Contains(page.Shapes, shape => shape.Id == "gallery-title");
            Assert.Contains(page.Shapes, shape => shape.Id == "gallery-title-2");
            Assert.Equal(page.Shapes.Count, page.Shapes.Select(shape => shape.Id).Distinct(StringComparer.Ordinal).Count());
            Assert.Equal(1D, page.Shapes.Single(shape => shape.Id == "gallery-0-shape").Width, 6);
            Assert.Equal(0.5D, page.Shapes.Single(shape => shape.Id == "gallery-0-shape").Height, 6);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void StencilCatalogGalleryReservesExistingConnectorIds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioStencilCatalog catalog = VisioStencilCatalog.Create("Gallery Catalog", builder => builder
                .Add("gallery.api", "API", "Process", "Integration", 1.8, 0.9));
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Gallery", 5, 4);
            VisioShape left = new VisioShape("left", 0.8, 0.8, 0.5, 0.5, "L");
            VisioShape right = new VisioShape("right", 1.8, 0.8, 0.5, 0.5, "R");
            page.Shapes.Add(left);
            page.Shapes.Add(right);
            page.AddConnector("gallery-title", left, right, ConnectorKind.Dynamic);

            page.AddStencilGallery(catalog, new VisioStencilGalleryOptions {
                IdPrefix = "gallery",
                Columns = 1,
                MaxShapes = 1,
                Title = "Reusable palette"
            });

            Assert.Contains(page.Connectors, connector => connector.Id == "gallery-title");
            Assert.Contains(page.Shapes, shape => shape.Id == "gallery-title-2" && shape.Text == "Reusable palette");
            Assert.DoesNotContain(page.Shapes, shape => shape.Id == "gallery-title");

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
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
        public void StencilMetadataKeepsPreviousPublicOverloads() {
            Type enumerableType = typeof(IEnumerable<string>);

            Assert.NotNull(typeof(VisioStencilShape).GetConstructor(new[] {
                typeof(string),
                typeof(string),
                typeof(string),
                typeof(string),
                typeof(double),
                typeof(double),
                enumerableType,
                enumerableType,
                enumerableType,
                typeof(string)
            }));
            Assert.Contains(
                typeof(VisioStencilCatalogBuilder).GetMethods().Where(method => method.Name == nameof(VisioStencilCatalogBuilder.AddWithMetadata)),
                method => method.GetParameters().Length == 10);
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
        public void PackageStencilCatalogLearnsNativeMasterDimensionsWithoutRuntimeTemplateDependency() {
            string packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            CreatePackageWithMasterDimensions(
                packagePath,
                ("Rectangle", "Wide Box", 3.2, 1.1, null),
                ("Decision", "Metric Decision", 2.0, 1.0, "MM"));

            VisioStencilCatalog catalog = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                CatalogName = "Learned Sizes",
                Category = "Learned",
                IdPrefix = "learned",
                DefaultWidth = 9,
                DefaultHeight = 7
            });

            VisioStencilShape wideBox = catalog.Get("wide-box");
            VisioStencilShape metricDecision = catalog.Get("metric-decision");

            Assert.Equal(3.2, wideBox.DefaultWidth, 6);
            Assert.Equal(1.1, wideBox.DefaultHeight, 6);
            Assert.Equal(VisioMeasurementUnit.Inches, wideBox.DefaultUnit);
            Assert.Equal(2.0, metricDecision.DefaultWidth, 6);
            Assert.Equal(1.0, metricDecision.DefaultHeight, 6);
            Assert.Equal(VisioMeasurementUnit.Inches, metricDecision.DefaultUnit);

            using MemoryStream manifest = new();
            catalog.Save(manifest);
            manifest.Position = 0;
            VisioStencilCatalog loadedCatalog = VisioStencilCatalog.Load(manifest);
            Assert.Equal(VisioMeasurementUnit.Inches, loadedCatalog.Get("wide-box").DefaultUnit);

            VisioStencilCatalog fallbackCatalog = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                LearnMasterDimensions = false,
                DefaultWidth = 9,
                DefaultHeight = 7
            });
            Assert.Equal(9, fallbackCatalog.Get("wide-box").DefaultWidth);
            Assert.Null(fallbackCatalog.Get("wide-box").DefaultUnit);

            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Learned Stencils", 20, 15, VisioMeasurementUnit.Centimeters);
            VisioShape shape = page.AddStencilShape(catalog, "wide-box", "wide", 5, 8);
            VisioShape coordinateUnitShape = page.AddStencilShape(wideBox, "wide-cm", 6, 9, "Wide in cm", VisioMeasurementUnit.Centimeters);
            VisioShape explicitShape = page.AddStencilShape(catalog, "wide-box", "explicit", 10, 8, 4, 2, "Explicit size");
            VisioShape resized = page.AddRectangle(14, 8, 1, 1, "Resize me", VisioMeasurementUnit.Centimeters);
            page.ReplaceMaster(resized, wideBox, resizeToMaster: true);
            document.Save();

            Assert.Equal(5.0 / 2.54, shape.PinX, 6);
            Assert.Equal(8.0 / 2.54, shape.PinY, 6);
            Assert.Equal(3.2, shape.Width, 6);
            Assert.Equal(1.1, shape.Height, 6);
            Assert.Equal(6.0 / 2.54, coordinateUnitShape.PinX, 6);
            Assert.Equal(9.0 / 2.54, coordinateUnitShape.PinY, 6);
            Assert.Equal(3.2, coordinateUnitShape.Width, 6);
            Assert.Equal(1.1, coordinateUnitShape.Height, 6);
            Assert.Equal(4.0 / 2.54, explicitShape.Width, 6);
            Assert.Equal(2.0 / 2.54, explicitShape.Height, 6);
            Assert.Equal(3.2, resized.Width, 6);
            Assert.Equal(1.1, resized.Height, 6);
            Assert.Empty(VisioValidator.Validate(filePath));
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
        public void ImportedStencilMastersPreserveExternalMasterArtwork() {
            string packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            CreatePackageWithRawGroupMaster(packagePath, "FancyCloud", "Fancy Cloud");
            VisioStencilCatalog catalog = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                IncludeUnsupportedMasters = true,
                Category = "External"
            });

            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            IReadOnlyList<VisioMaster> imported = document.ImportStencilMastersAndGet(packagePath, new[] { "fancy-cloud" });
            VisioPage page = document.AddPage("External Stencils");
            VisioShape cloud = page.AddStencilShape(catalog, "fancy-cloud", "cloud", 2, 4, "Cloud");
            document.Save();

            Assert.Single(imported);
            Assert.Same(imported[0], cloud.Master);
            Assert.Equal("FancyCloud", cloud.MasterNameU);
            Assert.Empty(VisioValidator.Validate(filePath));

            using ZipArchive zip = ZipFile.OpenRead(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XDocument masterDocument = XDocument.Load(zip.GetEntry("visio/masters/master1.xml")!.Open());
            XElement rootShape = masterDocument.Root!.Element(ns + "Shapes")!.Element(ns + "Shape")!;
            Assert.Equal("5", (string?)rootShape.Attribute("ID"));
            Assert.Equal("Group", (string?)rootShape.Attribute("Type"));
            Assert.NotNull(rootShape.Element(ns + "Shapes")?.Element(ns + "Shape"));

            XDocument pageDocument = XDocument.Load(zip.GetEntry("visio/pages/page1.xml")!.Open());
            XElement pageShape = pageDocument.Root!.Element(ns + "Shapes")!.Element(ns + "Shape")!;
            Assert.Null(pageShape.Attribute("MasterShape"));
            XElement pageChildShape = pageShape.Element(ns + "Shapes")!.Element(ns + "Shape")!;
            Assert.Equal("6", (string?)pageChildShape.Attribute("MasterShape"));
            Assert.DoesNotContain(pageShape.Elements(ns + "Section"), section => (string?)section.Attribute("N") == "Geometry");

            XDocument documentXml = XDocument.Load(zip.GetEntry("visio/document.xml")!.Open());
            Assert.NotNull(documentXml.Root!.Element(ns + "Colors")!.Elements(ns + "ColorEntry").FirstOrDefault(element => (string?)element.Attribute("IX") == "24"));
            Assert.NotNull(documentXml.Root!.Element(ns + "StyleSheets")!.Elements(ns + "StyleSheet").FirstOrDefault(element => (string?)element.Attribute("ID") == "8"));
            Assert.NotNull(zip.GetEntry("visio/theme/theme1.xml"));
            Assert.NotNull(zip.GetEntry("visio/media/officeimo-master1-rel1.emf"));
            XNamespace packageRel = "http://schemas.openxmlformats.org/package/2006/relationships";
            XDocument masterRelationships = XDocument.Load(zip.GetEntry("visio/masters/_rels/master1.xml.rels")!.Open());
            Assert.NotNull(masterRelationships.Root!.Elements(packageRel + "Relationship").FirstOrDefault(element =>
                (string?)element.Attribute("Id") == "rIdImage" &&
                ((string?)element.Attribute("Target"))!.Contains("officeimo-master1-rel1.emf", StringComparison.OrdinalIgnoreCase)));
        }

        [Fact]
        public void PackageCatalogLoadManyAutoImportsSourceMasters() {
            string firstPackage = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            string secondPackage = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            CreatePackageWithRawGroupMaster(firstPackage, "FancyCloud", "Fancy Cloud");
            CreatePackageWithRawGroupMaster(secondPackage, "DataVault", "Data Vault");

            VisioStencilCatalog catalog = VisioStencilPackageCatalog.LoadMany(new[] { firstPackage, secondPackage }, new VisioStencilPackageLoadOptions {
                CatalogName = "Combined",
                IncludeUnsupportedMasters = true
            });

            VisioStencilShape cloudStencil = catalog.Get("fancy-cloud");
            VisioStencilShape vaultStencil = catalog.Get("data-vault");
            string manifestPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xml");
            catalog.Save(manifestPath);
            VisioStencilCatalog reloadedCatalog = VisioStencilCatalog.Load(manifestPath);
            Assert.Equal(Path.GetFullPath(firstPackage), reloadedCatalog.Get("fancy-cloud").SourcePackagePath);

            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("External Stencils");
            VisioShape cloud = page.AddStencilShape(cloudStencil, "cloud", 2, 4);
            VisioShape vault = page.AddStencilShape(vaultStencil, "vault", 5, 4);
            document.Save();

            Assert.Equal(Path.GetFullPath(firstPackage), cloudStencil.SourcePackagePath);
            Assert.Equal(Path.GetFullPath(secondPackage), vaultStencil.SourcePackagePath);
            Assert.Equal(cloudStencil.MasterNameU, cloud.MasterNameU);
            Assert.Equal(vaultStencil.MasterNameU, vault.MasterNameU);
            Assert.NotNull(cloud.Master);
            Assert.NotNull(vault.Master);
            Assert.Empty(VisioValidator.Validate(filePath));

            using ZipArchive zip = ZipFile.OpenRead(filePath);
            Assert.NotNull(zip.GetEntry("visio/masters/master1.xml"));
            Assert.NotNull(zip.GetEntry("visio/masters/master2.xml"));
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

            XElement documentRoot = new(ns + "VisioDocument",
                new XElement(ns + "DocumentSettings"),
                new XElement(ns + "Colors",
                    new XElement(ns + "ColorEntry", new XAttribute("IX", "24"), new XAttribute("RGB", "#50E6FF"))),
                new XElement(ns + "FaceNames",
                    new XElement(ns + "FaceName", new XAttribute("NameU", "Sample UI"))),
                new XElement(ns + "StyleSheets",
                    new XElement(ns + "StyleSheet",
                        new XAttribute("ID", "8"),
                        new XAttribute("Name", "External Azure"),
                        new XAttribute("NameU", "External Azure"),
                        new XAttribute("LineStyle", "0"),
                        new XAttribute("FillStyle", "0"),
                        new XAttribute("TextStyle", "0"),
                        new XElement(ns + "Cell", new XAttribute("N", "FillForegnd"), new XAttribute("V", "#50E6FF")),
                        new XElement(ns + "Cell", new XAttribute("N", "LineColor"), new XAttribute("V", "#0078D4")))));
            WriteZipXml(zip, "visio/document.xml", new XDocument(documentRoot));
            WriteZipXml(zip, "visio/theme/theme1.xml", new XDocument(new XElement(XName.Get("theme", "http://schemas.openxmlformats.org/drawingml/2006/main"), new XAttribute("name", "External Theme"))));

            XElement childShape = new(ns + "Shape",
                new XAttribute("ID", "6"),
                new XAttribute("NameU", "FancyCloud.Icon"),
                new XAttribute("Type", "Shape"),
                new XElement(ns + "Cell", new XAttribute("N", "PinX"), new XAttribute("V", "0.5")),
                new XElement(ns + "Cell", new XAttribute("N", "PinY"), new XAttribute("V", "0.35")),
                new XElement(ns + "Cell", new XAttribute("N", "Width"), new XAttribute("V", "0.6")),
                new XElement(ns + "Cell", new XAttribute("N", "Height"), new XAttribute("V", "0.4")),
                new XElement(ns + "Cell", new XAttribute("N", "LocPinX"), new XAttribute("V", "0.3")),
                new XElement(ns + "Cell", new XAttribute("N", "LocPinY"), new XAttribute("V", "0.2")),
                new XElement(ns + "Section",
                    new XAttribute("N", "Geometry"),
                    new XAttribute("IX", "0"),
                    new XElement(ns + "Row", new XAttribute("T", "MoveTo"),
                        new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                        new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.2"))),
                    new XElement(ns + "Row", new XAttribute("T", "LineTo"),
                        new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.2")),
                        new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.4"))),
                    new XElement(ns + "Row", new XAttribute("T", "LineTo"),
                        new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.6")),
                        new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.3"))),
                    new XElement(ns + "Row", new XAttribute("T", "LineTo"),
                        new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0.6")),
                        new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.1"))),
                    new XElement(ns + "Row", new XAttribute("T", "LineTo"),
                        new XElement(ns + "Cell", new XAttribute("N", "X"), new XAttribute("V", "0")),
                        new XElement(ns + "Cell", new XAttribute("N", "Y"), new XAttribute("V", "0.2")))));
            XElement groupShape = new(ns + "Shape",
                new XAttribute("ID", "5"),
                new XAttribute("Name", name),
                new XAttribute("NameU", nameU),
                new XAttribute("Type", "Group"),
                new XAttribute("LineStyle", "8"),
                new XAttribute("FillStyle", "8"),
                new XAttribute("TextStyle", "8"),
                new XElement(ns + "Cell", new XAttribute("N", "PinX"), new XAttribute("V", "0.5")),
                new XElement(ns + "Cell", new XAttribute("N", "PinY"), new XAttribute("V", "0.5")),
                new XElement(ns + "Cell", new XAttribute("N", "Width"), new XAttribute("V", "1")),
                new XElement(ns + "Cell", new XAttribute("N", "Height"), new XAttribute("V", "1")),
                new XElement(ns + "Cell", new XAttribute("N", "LocPinX"), new XAttribute("V", "0.5")),
                new XElement(ns + "Cell", new XAttribute("N", "LocPinY"), new XAttribute("V", "0.5")),
                new XElement(ns + "Shapes", childShape));
            XDocument masterDocument = new(new XElement(ns + "MasterContents",
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                new XAttribute(XNamespace.Xmlns + "r", officeRelationshipNamespace),
                new XElement(ns + "Shapes", groupShape),
                new XElement(ns + "ForeignData",
                    new XAttribute("ForeignType", "Bitmap"),
                    new XAttribute(rel + "id", "rIdImage"))));
            WriteZipXml(zip, "visio/masters/master42.xml", masterDocument);

            XElement masterRelRoot = new(packageRel + "Relationships",
                new XElement(packageRel + "Relationship",
                    new XAttribute("Id", "rIdImage"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
                    new XAttribute("Target", "../media/image1.emf")));
            WriteZipXml(zip, "visio/masters/_rels/master42.xml.rels", new XDocument(masterRelRoot));

            ZipArchiveEntry mediaEntry = zip.CreateEntry("visio/media/image1.emf");
            using Stream mediaStream = mediaEntry.Open();
            byte[] media = { 1, 0, 0, 0, 32, 69, 77, 70 };
            mediaStream.Write(media, 0, media.Length);
        }

        private static void CreatePackageWithMasterDimensions(string path, params (string NameU, string? Name, double Width, double Height, string? Unit)[] masters) {
            const string visioNamespace = "http://schemas.microsoft.com/office/visio/2012/main";
            const string officeRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            const string packageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

            using ZipArchive zip = ZipFile.Open(path, ZipArchiveMode.Create);
            XNamespace ns = visioNamespace;
            XNamespace rel = officeRelationshipNamespace;
            XElement mastersRoot = new(ns + "Masters",
                masters.Select((master, index) =>
                    new XElement(ns + "Master",
                        new XAttribute("ID", index + 1),
                        new XAttribute("Name", master.Name ?? master.NameU),
                        new XAttribute("NameU", master.NameU),
                        new XElement(ns + "Rel", new XAttribute(rel + "id", $"rId{index + 1}")))));
            WriteZipXml(zip, "visio/masters/masters.xml", new XDocument(mastersRoot));

            XNamespace packageRel = packageRelationshipNamespace;
            XElement relationshipsRoot = new(packageRel + "Relationships",
                masters.Select((master, index) =>
                    new XElement(packageRel + "Relationship",
                        new XAttribute("Id", $"rId{index + 1}"),
                        new XAttribute("Type", officeRelationshipNamespace + "/master"),
                        new XAttribute("Target", $"master{index + 1}.xml"))));
            WriteZipXml(zip, "visio/masters/_rels/masters.xml.rels", new XDocument(relationshipsRoot));

            for (int index = 0; index < masters.Length; index++) {
                (string nameU, string? name, double width, double height, string? unit) = masters[index];
                XElement shape = new(ns + "Shape",
                    new XAttribute("ID", "1"),
                    new XAttribute("Name", name ?? nameU),
                    new XAttribute("NameU", nameU),
                    DimensionCell(ns, "Width", width, unit),
                    DimensionCell(ns, "Height", height, unit));
                XDocument masterDocument = new(new XElement(ns + "MasterContents", new XElement(ns + "Shapes", shape)));
                WriteZipXml(zip, $"visio/masters/master{index + 1}.xml", masterDocument);
            }
        }

        private static XElement DimensionCell(XNamespace ns, string name, double value, string? unit) {
            XElement cell = new(ns + "Cell",
                new XAttribute("N", name),
                new XAttribute("V", value.ToString(System.Globalization.CultureInfo.InvariantCulture)));
            if (!string.IsNullOrWhiteSpace(unit)) {
                cell.Add(new XAttribute("U", unit));
            }

            return cell;
        }

        private static void WriteZipXml(ZipArchive zip, string path, XDocument document) {
            ZipArchiveEntry entry = zip.CreateEntry(path);
            using Stream stream = entry.Open();
            using StreamWriter writer = new(stream, new UTF8Encoding(false));
            writer.Write(document.Declaration + Environment.NewLine + document.ToString(SaveOptions.DisableFormatting));
        }
    }
}
