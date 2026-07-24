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
            Assert.Equal("Start/End", VisioStencils.Flowchart.Get("terminator").Name);
            Assert.Equal("Continuation", VisioStencils.Flowchart.Get("continuation").Name);

            Assert.Contains(VisioStencils.All.Shapes, shape => shape.Id == "basic.rectangle");
            Assert.Contains(VisioStencils.All.Shapes, shape => shape.Id == "flow.off-page-reference");
            Assert.Equal("Block", VisioStencils.BlockDiagram.Get("component").Name);
            Assert.Equal("Decision Block", VisioStencils.BlockDiagram.Get("branch").Name);
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
        public void ExpandedBuiltInCatalogsExposePremiumDomainStencilPacks() {
            Assert.Contains(VisioStencils.All.Categories, category => category == "Infrastructure");
            Assert.Contains(VisioStencils.All.Categories, category => category == "Cloud");
            Assert.Contains(VisioStencils.All.Categories, category => category == "Security and Identity");
            Assert.Contains(VisioStencils.All.Categories, category => category == "Containers and Kubernetes");
            Assert.Contains(VisioStencils.All.Categories, category => category == "Data and Platform");
            Assert.Contains(VisioStencils.All.Categories, category => category == "Collaboration and Business Process");
            Assert.True(VisioStencils.All.Shapes.Count >= 120);

            Assert.Equal("Load Balancer", VisioStencils.Infrastructure.Get("traffic").Name);
            Assert.Equal("Function", VisioStencils.Cloud.Get("serverless").Name);
            Assert.Equal("Policy", VisioStencils.SecurityIdentity.Get("conditional-access").Name);
            Assert.Equal("Cluster", VisioStencils.ContainersKubernetes.Get("kubernetes").Name);
            Assert.Equal("Pipeline", VisioStencils.DataPlatform.Get("etl").Name);
            Assert.Equal("Approval", VisioStencils.CollaborationBusiness.Get("sign-off").Name);
            Assert.Equal("Security Alert", VisioStencils.All.FindBest("security-alert", "incident").Name);

            IReadOnlyList<VisioStencilShape> identityMatches = VisioStencils.All.Search("identity");
            Assert.Contains(identityMatches, shape => shape.Id == "sec.identity-provider");
            Assert.Contains(VisioStencils.All.Search("kubernetes"), shape => shape.Id == "k8s.cluster");
            Assert.Contains(VisioStencils.All.Search("event-stream"), shape => shape.Id == "data.stream");
            Assert.All(VisioStencils.All.InCategory("Security and Identity"), shape => Assert.Equal("Security and Identity", shape.Category));
        }

        [Fact]
        public void ExpandedBuiltInCatalogShapesArePlaceableAndProfiled() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.UseMastersByDefault = true;
            VisioPage page = document.AddPage("Expanded Stencils", 11, 8.5);

            VisioShape idp = page.AddStencilShape(VisioStencils.SecurityIdentity, "idp", "idp", 1.8, 6.3, "Entra ID");
            VisioShape cluster = page.AddStencilShape(VisioStencils.ContainersKubernetes, "kubernetes", "cluster", 4.3, 6.3, "AKS Cluster");
            VisioShape lake = page.AddStencilShape(VisioStencils.DataPlatform, "data.lake", "lake", 6.9, 6.3, "Lake");
            VisioShape team = page.AddStencilShape(VisioStencils.CollaborationBusiness, "team", "team", 9.1, 6.3, "Ops Team");
            page.AddConnector(idp, cluster, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left).Label = "tokens";
            page.AddConnector(cluster, lake, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left).Label = "events";
            page.AddConnector(team, cluster, ConnectorKind.Dynamic, VisioSide.Left, VisioSide.Right).Label = "runbook";

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(4, loaded.Pages[0].Shapes.Count);
            Assert.Equal("sec.identity-provider", GetUserCellValue(loaded.Pages[0].Shapes.Single(shape => shape.Id == "idp"), "OfficeIMO.StencilId"));
            Assert.Equal("k8s.cluster", GetUserCellValue(loaded.Pages[0].Shapes.Single(shape => shape.Id == "cluster"), "OfficeIMO.StencilId"));
            Assert.Equal("data.lake", GetUserCellValue(loaded.Pages[0].Shapes.Single(shape => shape.Id == "lake"), "OfficeIMO.StencilId"));
            Assert.Equal("collab.team", GetUserCellValue(loaded.Pages[0].Shapes.Single(shape => shape.Id == "team"), "OfficeIMO.StencilId"));

            string profile = loaded.CreateStencilProfile().ToText();
            Assert.Contains("Security and Identity", profile);
            Assert.Contains("Containers and Kubernetes", profile);
            Assert.Contains("Data and Platform", profile);
            Assert.Contains("Collaboration and Business Process", profile);
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
        public void GeneratedStencilShapesUseRendererFriendlyLocalGeometry() {
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
            Assert.Contains(serviceShape.Elements(v + "Section"), section => (string?)section.Attribute("N") == "Geometry");
            Assert.Contains(databaseShape.Elements(v + "Section"), section => (string?)section.Attribute("N") == "Geometry");

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
        public void StencilGalleryDocumentCreatesPagedCategoryReviewDocument() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioStencilCatalog catalog = VisioStencilCatalog.Create("Review Catalog", builder => builder
                .AddWithMetadata("review.worker", "Worker", "Process", "Compute", 1.6, 0.9, keywords: new[] { "service" }, aliases: new[] { "processor" }, tags: new[] { "runtime" })
                .AddWithMetadata("review.runtime", "Runtime", "Rectangle", "Compute", 1.6, 0.9, keywords: new[] { "container" })
                .AddWithMetadata("review.api", "API", "Process", "Integration", 1.7, 0.9, keywords: new[] { "gateway" })
                .AddWithMetadata("review.queue", "Queue", "Data", "Integration", 1.5, 0.85, keywords: new[] { "broker" })
                .AddWithMetadata("review.topic", "Topic", "Data", "Integration", 1.5, 0.85, keywords: new[] { "event" }));

            VisioDocument document = VisioStencilGalleryDocument.Create(filePath, catalog, new VisioStencilGalleryDocumentOptions {
                Title = "Reusable stencil review",
                IdPrefix = "review",
                Columns = 2,
                ShapesPerPage = 2,
                PageWidth = 7,
                PageHeight = 5,
                IncludeStencilMetadataShapeData = true
            });

            Assert.Equal("Reusable stencil review", document.Title);
            Assert.Equal("OfficeIMO.Visio", document.Author);
            Assert.True(document.UseMastersByDefault);
            Assert.Equal(4, document.Pages.Count);
            Assert.Equal("Stencil Gallery Overview", document.Pages[0].Name);
            Assert.Equal("01 Compute", document.Pages[1].Name);
            Assert.Equal("02 Integration (1 of 2)", document.Pages[2].Name);
            Assert.Equal("03 Integration (2 of 2)", document.Pages[3].Name);
            Assert.Contains(document.Pages[0].Shapes, shape => shape.Id == "review-overview-summary" && shape.Text == "5 stencils across 2 categories");

            VisioShape firstComputeShape = document.Pages[1].Shapes.Single(shape => shape.Id == "review-01-0-shape");
            Assert.Equal("Runtime", firstComputeShape.GetShapeDataValue("StencilName"));
            Assert.Equal("Compute", firstComputeShape.GetShapeDataValue("StencilCategory"));
            Assert.Equal("Review Catalog - Compute", firstComputeShape.GetShapeDataValue("StencilCatalog"));
            Assert.Equal("Rectangle", firstComputeShape.GetShapeDataValue("MasterNameU"));
            Assert.Equal("0", firstComputeShape.GetShapeDataValue("GalleryIndex"));
            Assert.Equal("Rectangle", firstComputeShape.MasterNameU);

            VisioShape integrationShape = document.Pages[2].Shapes.Single(shape => shape.Id == "review-02-1-shape");
            Assert.Equal("Queue", integrationShape.GetShapeDataValue("StencilName"));
            Assert.Equal("broker", integrationShape.GetShapeDataValue("Keywords"));

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));
            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(4, loaded.Pages.Count);
            Assert.Equal("Runtime", loaded.Pages[1].Shapes.Single(shape => shape.Id == "review-01-0-shape").GetShapeDataValue("StencilName"));
            Assert.Equal("Queue", loaded.Pages[2].Shapes.Single(shape => shape.Id == "review-02-1-shape").GetShapeDataValue("StencilName"));
        }

        [Fact]
        public void StencilGalleryDocumentRestoresExistingMasterPreferenceAndReservesOverviewIds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioStencilCatalog catalog = VisioStencilCatalog.Create("Review Catalog", builder => builder
                .AddWithMetadata("review.api", "API", "Process", "A/B", 1.7, 0.9)
                .AddWithMetadata("review.worker", "Worker", "Process", "A B", 1.6, 0.9));
            VisioDocument document = VisioDocument.Create(filePath);
            document.UseMastersByDefault = false;

            document.AddStencilGalleryDocument(catalog, new VisioStencilGalleryDocumentOptions {
                IdPrefix = "review",
                ShapesPerPage = 4,
                UseMastersByDefault = true
            });

            VisioPage overview = document.Pages.Single(page => page.Name == "Stencil Gallery Overview");
            string[] ids = overview.Shapes.Select(shape => shape.Id).ToArray();
            Assert.False(document.UseMastersByDefault);
            Assert.Equal(ids.Length, ids.Distinct(StringComparer.OrdinalIgnoreCase).Count());
            Assert.Contains(overview.Shapes, shape => shape.Id == "review-overview-category-a-b");
            Assert.Contains(overview.Shapes, shape => shape.Id == "review-overview-category-a-b-2");
        }

        [Fact]
        public void StencilGalleryDocumentValidatesPagingOptions() {
            VisioStencilCatalog catalog = VisioStencilCatalog.Create("Review Catalog", builder => builder
                .Add("review.worker", "Worker", "Process", "Compute", 1.6, 0.9));
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            Assert.Throws<ArgumentOutOfRangeException>(() => VisioStencilGalleryDocument.Create(filePath, catalog, new VisioStencilGalleryDocumentOptions {
                ShapesPerPage = 0
            }));
            Assert.Throws<ArgumentException>(() => VisioStencilGalleryDocument.Create(filePath, catalog, new VisioStencilGalleryDocumentOptions {
                IdPrefix = " "
            }));
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
        public void StencilCatalogManifestDropsSourcePackagePathsByDefault() {
            VisioStencilCatalog source = VisioStencilCatalog.Create("Untrusted Catalog", builder => builder
                .AddWithMetadata(
                    "external.master",
                    "External Master",
                    "ExternalMaster",
                    "External",
                    1.8,
                    0.9,
                    keywords: null,
                    aliases: null,
                    tags: null,
                    iconNameU: "ExternalMaster",
                    defaultUnit: null,
                    sourcePackagePath: Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx")));
            using MemoryStream manifest = new();
            source.Save(manifest);
            manifest.Position = 0;

            VisioStencilCatalog loaded = VisioStencilCatalog.Load(manifest);
            VisioStencilShape stencil = loaded.Get("external.master");
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Untrusted");
            VisioShape shape = page.AddStencilShape(stencil, "shape", 2, 4);

            Assert.Null(stencil.SourcePackagePath);
            Assert.Null(shape.Master);
        }

        [Fact]
        public void StencilCatalogManifestRequiresExternalSourcePackagePathOptIn() {
            string baseDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(baseDirectory);
            string externalPackagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            string manifestPath = Path.Combine(baseDirectory, "catalog.xml");

            try {
                VisioStencilCatalog source = VisioStencilCatalog.Create("Untrusted Catalog", builder => builder
                    .AddWithMetadata(
                        "external.master",
                        "External Master",
                        "ExternalMaster",
                        "External",
                        1.8,
                        0.9,
                        keywords: null,
                        aliases: null,
                        tags: null,
                        iconNameU: "ExternalMaster",
                        defaultUnit: null,
                        sourcePackagePath: externalPackagePath));
                source.Save(manifestPath);

                VisioStencilCatalog bounded = VisioStencilCatalog.Load(manifestPath, new VisioStencilCatalogManifestLoadOptions {
                    AllowSourcePackagePaths = true
                });
                VisioStencilCatalog trustedExternal = VisioStencilCatalog.Load(manifestPath, new VisioStencilCatalogManifestLoadOptions {
                    AllowSourcePackagePaths = true,
                    AllowExternalSourcePackagePaths = true
                });

                Assert.Null(bounded.Get("external.master").SourcePackagePath);
                Assert.Equal(Path.GetFullPath(externalPackagePath), trustedExternal.Get("external.master").SourcePackagePath);
            } finally {
                if (Directory.Exists(baseDirectory)) {
                    Directory.Delete(baseDirectory, recursive: true);
                }
            }
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
        public void PackageStencilCatalogLearnsNativeConnectionPointsAndAppliesThemToPlacedShapes() {
            string packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            CreatePackageWithMasterConnectionPoints(packagePath, "Rectangle", "Connected Box");

            VisioStencilCatalog catalog = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                CatalogName = "Connected Stencils",
                Category = "Connected",
                IdPrefix = "connected",
                IncludeUnsupportedMasters = true
            });

            VisioStencilShape stencil = catalog.Get("connected-box");
            Assert.Equal(3, stencil.SourceConnectionPoints.Count);
            Assert.Equal(0, stencil.SourceConnectionPoints[0].SectionIndex);
            Assert.Equal(3.2, stencil.SourceConnectionPoints[0].SourceWidth);
            Assert.Equal(1.1, stencil.SourceConnectionPoints[0].SourceHeight);
            Assert.Equal(3.2, stencil.DefaultWidth, 6);
            Assert.Equal(1.1, stencil.DefaultHeight, 6);

            using MemoryStream manifest = new();
            catalog.Save(manifest);
            manifest.Position = 0;
            VisioStencilShape reloadedStencil = VisioStencilCatalog.Load(manifest).Get("connected-box");
            Assert.Equal(3, reloadedStencil.SourceConnectionPoints.Count);
            Assert.Equal(1.6, reloadedStencil.SourceConnectionPoints[2].X, 6);
            Assert.Equal(3.2, reloadedStencil.SourceConnectionPoints[2].SourceWidth);

            VisioStencilCatalog withoutConnectionMetadata = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                ExtractConnectionPointMetadata = false,
                IncludeUnsupportedMasters = true
            });
            Assert.Empty(withoutConnectionMetadata.Get("connected-box").SourceConnectionPoints);
            VisioStencilShape nativeSizedPoints = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                IncludeUnsupportedMasters = true,
                LearnMasterDimensions = false,
                DefaultWidth = 1,
                DefaultHeight = 1
            }).Get("connected-box");

            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Connected Stencils");
            VisioShape shape = page.AddStencilShape(catalog, "connected-box", "node", 3, 4, 6.4, 2.2, "Connected");
            VisioShape overriddenDefaultShape = page.AddStencilShape(nativeSizedPoints, "native-sized", 8, 4, 6.4, 2.2, "Native sized");
            document.Save();

            Assert.Equal(3, shape.ConnectionPoints.Count);
            Assert.Equal(0, shape.ConnectionPoints[0].X, 6);
            Assert.Equal(1.1, shape.ConnectionPoints[0].Y, 6);
            Assert.Equal(6.4, shape.ConnectionPoints[1].X, 6);
            Assert.Equal(1.1, shape.ConnectionPoints[1].Y, 6);
            Assert.Equal(3.2, shape.ConnectionPoints[2].X, 6);
            Assert.Equal(2.2, shape.ConnectionPoints[2].Y, 6);
            Assert.Equal(6.4, overriddenDefaultShape.ConnectionPoints[1].X, 6);
            Assert.Equal(2.2, overriddenDefaultShape.ConnectionPoints[2].Y, 6);
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape loadedShape = loaded.Pages[0].Shapes.Single(item => item.Id == "node");
            Assert.Equal(3, loadedShape.ConnectionPoints.Count);
            VisioStencilProfile profile = loaded.CreateStencilProfile();
            Assert.Equal(6, profile.TotalConnectionPoints);
            Assert.Equal(2, profile.ConnectionPointShapeCount);
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
        public void PackageStencilCatalogExtractsPreviewImageMetadataFromMasterRelationships() {
            string packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            CreatePackageWithRawGroupMaster(packagePath, "FancyCloud", "Fancy Cloud");

            VisioStencilCatalog catalog = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                IncludeUnsupportedMasters = true,
                Category = "External",
                IdPrefix = "preview"
            });
            VisioStencilShape stencil = catalog.Get("fancy-cloud");
            VisioStencilCatalog metadataWithoutDimensions = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                IncludeUnsupportedMasters = true,
                LearnMasterDimensions = false
            });
            VisioStencilCatalog metadataDisabled = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                IncludeUnsupportedMasters = true,
                ExtractPreviewImageMetadata = false
            });

            Assert.NotNull(stencil.PreviewImage);
            Assert.Equal("rIdImage", stencil.PreviewImage!.RelationshipId);
            Assert.Equal("../media/image1.emf", stencil.PreviewImage.Target);
            Assert.Equal("image/x-emf", stencil.PreviewImage.ContentType);
            Assert.Equal(".emf", stencil.PreviewImage.Extension);
            Assert.Equal(8, stencil.PreviewImage.ByteLength);
            Assert.Equal(".emf", metadataWithoutDimensions.Get("fancy-cloud").PreviewImage?.Extension);
            Assert.Null(metadataDisabled.Get("fancy-cloud").PreviewImage);

            VisioStencilPreviewImageData extracted = Assert.Single(VisioStencilPackageCatalog.ExtractPreviewImages(packagePath, new VisioStencilPackageLoadOptions {
                IncludeUnsupportedMasters = true
            }));
            Assert.Equal("42", extracted.MasterId);
            Assert.Equal("FancyCloud", extracted.MasterNameU);
            Assert.Equal("Fancy Cloud", extracted.MasterName);
            Assert.Equal("42-FancyCloud.emf", extracted.SuggestedFileName);
            Assert.Equal(new byte[] { 1, 0, 0, 0, 32, 69, 77, 70 }, extracted.ToBytes());
            VisioStencilPreviewImageData explicitlyExtracted = Assert.Single(VisioStencilPackageCatalog.ExtractPreviewImages(packagePath, new VisioStencilPackageLoadOptions {
                IncludeUnsupportedMasters = true,
                ExtractPreviewImageMetadata = false
            }));
            Assert.Equal(extracted.ToBytes(), explicitlyExtracted.ToBytes());

            string outputDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            string extractedPath = Assert.Single(VisioStencilPackageCatalog.ExtractPreviewImagesToDirectory(packagePath, outputDirectory, new VisioStencilPackageLoadOptions {
                IncludeUnsupportedMasters = true
            }));
            Assert.True(File.Exists(extractedPath));
            Assert.Equal(extracted.ToBytes(), File.ReadAllBytes(extractedPath));

            using MemoryStream manifest = new();
            catalog.Save(manifest);
            manifest.Position = 0;
            VisioStencilCatalog reloadedCatalog = VisioStencilCatalog.Load(manifest);
            Assert.Equal(".emf", reloadedCatalog.Get("fancy-cloud").PreviewImage?.Extension);

            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Preview Metadata");
            page.AddStencilShape(catalog, "fancy-cloud", "cloud", 2, 4);
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioInspectionMasterSnapshot master = Assert.Single(loaded.CreateInspectionSnapshot().Masters, item => item.NameU == "FancyCloud");
            Assert.Equal("rIdImage", master.StencilPreviewImageRelationshipId);
            Assert.Equal("../media/image1.emf", master.StencilPreviewImageTarget);
            Assert.Equal("image/x-emf", master.StencilPreviewImageContentType);
            Assert.Equal(".emf", master.StencilPreviewImageExtension);
            Assert.Equal(8, master.StencilPreviewImageByteLength);

            VisioStencilProfile profile = loaded.CreateStencilProfile();
            Assert.Equal(new[] { "image/x-emf" }, profile.StencilPreviewImageContentTypes);
            Assert.Equal(new[] { ".emf" }, profile.StencilPreviewImageExtensions);
            VisioStencilUsageProfile usage = Assert.Single(profile.Usages, item => item.Kind == VisioStencilProfileUsageKind.PackageBackedMaster);
            Assert.Equal("image/x-emf", usage.StencilPreviewImageContentType);
            Assert.Equal(".emf", usage.StencilPreviewImageExtension);
        }

        [Fact]
        public void PackageStencilCatalogRecognizesExternalPreviewImageBySharedTargetExtension() {
            string packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            CreatePackageWithRawGroupMaster(
                packagePath,
                "FancyWeb",
                "Fancy Web",
                "webp",
                relationshipType: "https://schemas.example.org/not-an-image",
                externalImageRelationship: true);

            VisioStencilCatalog catalog = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                IncludeUnsupportedMasters = true,
                Category = "External"
            });

            VisioStencilPreviewImage? previewImage = catalog.Get("fancy-web").PreviewImage;

            Assert.NotNull(previewImage);
            Assert.Equal("rIdImage", previewImage!.RelationshipId);
            Assert.Equal("../media/image1.webp", previewImage.Target);
            Assert.Null(previewImage.ContentType);
            Assert.Null(previewImage.Extension);
            Assert.Null(previewImage.ByteLength);
        }

        [Fact]
        public void PackageStencilPreviewGalleryWritesReviewableHtmlIndex() {
            string packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            CreatePackageWithRawGroupMaster(packagePath, "FancyCloud", "Fancy Cloud");
            string outputDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));

            VisioStencilPreviewGallery gallery = VisioStencilPackageCatalog.CreatePreviewGallery(
                packagePath,
                outputDirectory,
                new VisioStencilPackageLoadOptions {
                    IncludeUnsupportedMasters = true
                },
                new VisioStencilPreviewGalleryOptions {
                    Title = "External Preview Review",
                    PreviewDirectoryName = "assets",
                    IndexFileName = "gallery.html"
                });

            VisioStencilPreviewGalleryEntry entry = Assert.Single(gallery.Entries);
            Assert.Equal(Path.GetFullPath(packagePath), gallery.PackagePath);
            Assert.Equal(Path.Combine(outputDirectory, "assets"), gallery.PreviewDirectory);
            Assert.Equal(Path.Combine(outputDirectory, "gallery.html"), gallery.IndexPath);
            Assert.Equal("assets/42-FancyCloud.emf", entry.RelativePath);
            Assert.Equal("FancyCloud", entry.Image.MasterNameU);
            Assert.False(entry.IsBrowserRenderable);
            Assert.Equal(0, gallery.BrowserRenderableCount);
            Assert.Equal(0, gallery.ThumbnailCount);
            Assert.False(entry.HasThumbnail);
            Assert.True(File.Exists(entry.FilePath));
            Assert.True(File.Exists(gallery.IndexPath));
            Assert.Equal(new byte[] { 1, 0, 0, 0, 32, 69, 77, 70 }, File.ReadAllBytes(entry.FilePath));

            string html = File.ReadAllText(gallery.IndexPath!);
            Assert.Contains("<h1>External Preview Review</h1>", html);
            Assert.Contains("Fancy Cloud", html);
            Assert.Contains("image/x-emf", html);
            Assert.Contains("assets/42-FancyCloud.emf", html);
            Assert.Contains("<div class=\"fallback\">emf</div>", html);
        }

        [Fact]
        public void PackageStencilPreviewGalleryWritesBrowserRenderableThumbnailArtifacts() {
            string packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAFgwJ/lM9sWQAAAABJRU5ErkJggg==");
            const string displayName = "Fancy <Cloud> & \"QA\"";
            const string escapedDisplayName = "Fancy &lt;Cloud&gt; &amp; &quot;QA&quot;";
            CreatePackageWithRawGroupMaster(packagePath, "FancyCloud", displayName, "png", png);
            string outputDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));

            VisioStencilPreviewGallery gallery = VisioStencilPackageCatalog.CreatePreviewGallery(
                packagePath,
                outputDirectory,
                new VisioStencilPackageLoadOptions {
                    IncludeUnsupportedMasters = true
                },
                new VisioStencilPreviewGalleryOptions {
                    Title = "Renderable Preview Review",
                    PreviewDirectoryName = "assets",
                    ThumbnailDirectoryName = "thumbs",
                    ThumbnailWidth = 180,
                    ThumbnailHeight = 120
                });

            VisioStencilPreviewGalleryEntry entry = Assert.Single(gallery.Entries);
            Assert.True(entry.IsBrowserRenderable);
            Assert.True(entry.HasThumbnail);
            Assert.Equal(1, gallery.BrowserRenderableCount);
            Assert.Equal(1, gallery.ThumbnailCount);
            Assert.Equal(Path.Combine(outputDirectory, "thumbs"), gallery.ThumbnailDirectory);
            Assert.Equal("assets/42-FancyCloud.png", entry.RelativePath);
            Assert.Equal("thumbs/42-FancyCloud.thumbnail.svg", entry.ThumbnailRelativePath);
            Assert.True(File.Exists(entry.FilePath));
            Assert.True(File.Exists(entry.ThumbnailFilePath));
            Assert.Equal(png, File.ReadAllBytes(entry.FilePath));

            string thumbnail = File.ReadAllText(entry.ThumbnailFilePath!);
            Assert.Contains("width=\"180\"", thumbnail);
            Assert.Contains("height=\"120\"", thumbnail);
            Assert.Contains("data:image/png;base64,", thumbnail);
            Assert.Contains(Convert.ToBase64String(png), thumbnail);
            Assert.Contains("aria-label=\"" + escapedDisplayName + "\"", thumbnail);
            Assert.Contains(escapedDisplayName, thumbnail);
            Assert.Contains("<rect x=\"0\" y=\"0\" width=\"180\" height=\"120\" rx=\"8\" ry=\"8\" fill=\"#FFFFFF\"/>", thumbnail);
            Assert.Contains("<rect x=\"0.5\" y=\"0.5\" width=\"179\" height=\"119\" rx=\"7.5\" ry=\"7.5\" fill=\"none\" stroke=\"#D3E0EC\"/>", thumbnail);
            Assert.Contains("<clipPath id=\"visio-thumbnail-image-clip\"><rect x=\"14\" y=\"12\" width=\"152\" height=\"78\"/></clipPath>", thumbnail);
            Assert.Contains("<image x=\"14\" y=\"12\" width=\"152\" height=\"78\" clip-path=\"url(#visio-thumbnail-image-clip)\" preserveAspectRatio=\"xMidYMid meet\" href=\"data:image/png;base64,", thumbnail);
            Assert.Contains("<text x=\"14\" y=\"106\" font-family=\"Aptos, Segoe UI, Arial, sans-serif\" font-size=\"12\" text-anchor=\"start\" fill=\"#657586\">" + escapedDisplayName + "</text>", thumbnail);

            string html = File.ReadAllText(gallery.IndexPath!);
            Assert.Contains("<strong>1</strong> thumbnails", html);
            Assert.Contains("thumbs/42-FancyCloud.thumbnail.svg", html);
            Assert.Contains("assets/42-FancyCloud.png", html);
            Assert.Contains("<h2>" + escapedDisplayName + "</h2>", html);
        }

        [Fact]
        public void PackageStencilPreviewGalleryStoresSvgPreviewsAsDownloadOnlyText() {
            string packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\"><script>alert(1)</script><rect width=\"10\" height=\"10\"/></svg>");
            CreatePackageWithRawGroupMaster(packagePath, "FancyCloud", "Fancy Cloud", "svg", svg);
            string outputDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));

            VisioStencilPreviewGallery gallery = VisioStencilPackageCatalog.CreatePreviewGallery(
                packagePath,
                outputDirectory,
                new VisioStencilPackageLoadOptions {
                    IncludeUnsupportedMasters = true
                },
                new VisioStencilPreviewGalleryOptions {
                    Title = "SVG Preview Review",
                    PreviewDirectoryName = "assets",
                    ThumbnailDirectoryName = "thumbs"
                });

            VisioStencilPreviewGalleryEntry entry = Assert.Single(gallery.Entries);
            Assert.False(entry.IsBrowserRenderable);
            Assert.False(entry.HasThumbnail);
            Assert.Equal(0, gallery.BrowserRenderableCount);
            Assert.Equal(0, gallery.ThumbnailCount);
            Assert.Equal("assets/42-FancyCloud.svg.txt", entry.RelativePath);
            Assert.True(File.Exists(entry.FilePath));
            Assert.Equal(svg, File.ReadAllBytes(entry.FilePath));

            string html = File.ReadAllText(gallery.IndexPath!);
            Assert.Contains("<a download href=\"assets/42-FancyCloud.svg.txt\">", html);
            Assert.Contains("<div class=\"fallback\">svg</div>", html);
            Assert.DoesNotContain("href=\"assets/42-FancyCloud.svg\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<img src=\"assets/42-FancyCloud.svg.txt\"", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain(".thumbnail.svg", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ImportedStencilMastersPreserveExternalMasterArtwork() {
            string packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            CreatePackageWithRawGroupMaster(packagePath, "FancyCloud", "Fancy Cloud");
            VisioStencilCatalog catalog = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                IncludeUnsupportedMasters = true,
                Category = "External"
            });
            Assert.Equal("fancy-cloud", catalog.Get("fancy-cloud").Id);

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
            XElement masterUserSection = masterDocument.Root!.Element(ns + "PageSheet")!.Elements(ns + "Section").Single(section => (string?)section.Attribute("N") == "User");
            Assert.Equal("1", GetUserCellValue(masterUserSection, ns, "OfficeIMO.PackageBackedMaster"));
            Assert.Equal("fancy-cloud", GetUserCellValue(masterUserSection, ns, VisioSemanticUserCells.StencilId));

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
        public void ImportStencilMastersRejectsOversizedEmbeddedMasterRelationship() {
            string packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            byte[] oversizedMedia = new byte[checked((int)VisioAssets.MaxMasterRelationshipBytes + 1)];
            CreatePackageWithRawGroupMaster(packagePath, "FancyCloud", "Fancy Cloud", "png", oversizedMedia);
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                document.ImportStencilMastersAndGet(packagePath, new[] { "fancy-cloud" }));

            Assert.Contains(VisioAssets.MaxMasterRelationshipBytes.ToString(), exception.Message);
        }

        [Fact]
        public void ReplaceMasterImportsPackageMasterWhenNameUAlreadyRegistered() {
            string packagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            CreatePackageWithRawGroupMaster(packagePath, "Rectangle", "Package Rectangle");
            VisioStencilCatalog catalog = VisioStencilPackageCatalog.Load(packagePath, new VisioStencilPackageLoadOptions {
                IncludeUnsupportedMasters = true,
                Category = "External"
            });
            VisioStencilShape stencil = catalog.Get("rectangle");
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            document.RegisterMaster("Rectangle", new VisioShape("1", 0, 0, 1, 1, string.Empty) { NameU = "Rectangle" }, "builtin-rectangle");
            VisioPage page = document.AddPage("Replace");
            VisioShape shape = page.AddRectangle(2, 4, 1, 1, "Shape");

            page.ReplaceMaster(shape, stencil);

            Assert.NotNull(shape.Master);
            Assert.True(shape.Master!.IsPackageBacked);
            Assert.Equal(Path.GetFullPath(packagePath), shape.Master.StencilSourcePackagePath);
            Assert.Equal(stencil.Id, shape.GetUserCellValue(VisioSemanticUserCells.StencilId));
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
            VisioStencilCatalog untrustedCatalog = VisioStencilCatalog.Load(manifestPath);
            VisioStencilCatalog reloadedCatalog = VisioStencilCatalog.Load(manifestPath, new VisioStencilCatalogManifestLoadOptions {
                AllowSourcePackagePaths = true
            });
            Assert.Null(untrustedCatalog.Get("fancy-cloud").SourcePackagePath);
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

        private static string? GetUserCellValue(VisioShape shape, string name) {
            return shape.UserCells
                .FirstOrDefault(cell => string.Equals(cell.Name, name, StringComparison.OrdinalIgnoreCase))
                ?.Value;
        }

        private static string? GetUserCellValue(XElement userSection, XNamespace ns, string name) {
            return userSection.Elements(ns + "Row")
                .FirstOrDefault(row => string.Equals((string?)row.Attribute("N"), name, StringComparison.OrdinalIgnoreCase))
                ?.Elements(ns + "Cell")
                .FirstOrDefault(cell => string.Equals((string?)cell.Attribute("N"), "Value", StringComparison.OrdinalIgnoreCase))
                ?.Attribute("V")
                ?.Value;
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

        private static void CreatePackageWithRawGroupMaster(
            string path,
            string nameU,
            string name,
            string imageExtension = "emf",
            byte[]? imageData = null,
            string relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            bool externalImageRelationship = false) {
            const string visioNamespace = "http://schemas.microsoft.com/office/visio/2012/main";
            const string officeRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            const string packageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
            string normalizedExtension = imageExtension.TrimStart('.');
            byte[] media = imageData ?? new byte[] { 1, 0, 0, 0, 32, 69, 77, 70 };

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
                    new XAttribute("Type", relationshipType),
                    new XAttribute("Target", "../media/image1." + normalizedExtension),
                    externalImageRelationship ? new XAttribute("TargetMode", "External") : null));
            WriteZipXml(zip, "visio/masters/_rels/master42.xml.rels", new XDocument(masterRelRoot));

            if (!externalImageRelationship) {
                ZipArchiveEntry mediaEntry = zip.CreateEntry("visio/media/image1." + normalizedExtension);
                using Stream mediaStream = mediaEntry.Open();
                mediaStream.Write(media, 0, media.Length);
            }
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

        private static void CreatePackageWithMasterConnectionPoints(string path, string nameU, string name) {
            const string visioNamespace = "http://schemas.microsoft.com/office/visio/2012/main";
            const string officeRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            const string packageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

            using ZipArchive zip = ZipFile.Open(path, ZipArchiveMode.Create);
            XNamespace ns = visioNamespace;
            XNamespace rel = officeRelationshipNamespace;
            XElement mastersRoot = new(ns + "Masters",
                new XElement(ns + "Master",
                    new XAttribute("ID", "1"),
                    new XAttribute("Name", name),
                    new XAttribute("NameU", nameU),
                    new XElement(ns + "Rel", new XAttribute(rel + "id", "rId1"))));
            WriteZipXml(zip, "visio/masters/masters.xml", new XDocument(mastersRoot));

            XNamespace packageRel = packageRelationshipNamespace;
            XElement relationshipsRoot = new(packageRel + "Relationships",
                new XElement(packageRel + "Relationship",
                    new XAttribute("Id", "rId1"),
                    new XAttribute("Type", officeRelationshipNamespace + "/master"),
                    new XAttribute("Target", "master1.xml")));
            WriteZipXml(zip, "visio/masters/_rels/masters.xml.rels", new XDocument(relationshipsRoot));

            XElement shape = new(ns + "Shape",
                new XAttribute("ID", "1"),
                new XAttribute("Name", name),
                new XAttribute("NameU", nameU),
                DimensionCell(ns, "Width", 3.2, null),
                DimensionCell(ns, "Height", 1.1, null),
                new XElement(ns + "Section",
                    new XAttribute("N", "Connection"),
                    ConnectionRow(ns, 0, 0, 0.55, 1, 0),
                    ConnectionRow(ns, 1, 3.2, 0.55, -1, 0),
                    ConnectionRow(ns, 2, 1.6, 1.1, 0, -1)));
            XDocument masterDocument = new(new XElement(ns + "MasterContents", new XElement(ns + "Shapes", shape)));
            WriteZipXml(zip, "visio/masters/master1.xml", masterDocument);
        }

        private static XElement ConnectionRow(XNamespace ns, int index, double x, double y, double dirX, double dirY) {
            return new XElement(ns + "Row",
                new XAttribute("T", "Connection"),
                new XAttribute("IX", index),
                DimensionCell(ns, "X", x, null),
                DimensionCell(ns, "Y", y, null),
                DimensionCell(ns, "DirX", dirX, null),
                DimensionCell(ns, "DirY", dirY, null));
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
