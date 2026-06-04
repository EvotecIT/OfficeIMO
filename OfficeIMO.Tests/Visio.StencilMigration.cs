using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using OfficeIMO.Visio.Stencils;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Tests {
    public class VisioStencilMigrationTests {
        [Fact]
        public void StencilMigrationMapAppliesFirstMatchingRulesAndReportsChanges() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Migration", 11, 8.5);
            VisioShape intake = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "intake", 2, 6, "Intake");
            VisioShape review = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "review", 5, 6, "Review");
            VisioShape archive = page.AddStencilShape(VisioStencils.Flowchart.Get("data"), "archive", 8, 6, "Archive");
            VisioShape custom = new("custom", 5, 3, 2.5, 1, "Custom") { NameU = "LegacyCustom" };
            page.Shapes.Add(custom);
            review.FillColor = Color.LightYellow;
            review.SetShapeData("Owner", "Ops");
            review.SetUserCell("Stage", "Review", "STR");
            review.AddHyperlink("https://example.org/review");
            page.AddConnector(intake, review, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left).Label = "submit";
            page.AddConnector(review, archive, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left).Label = "store";
            document.Save();

            VisioStencilMigrationMap map = VisioStencilMigrationMap.Create(builder => builder
                .MapStencilId("flow.process", VisioStencils.Flowchart.Get("preparation"), resizeToStencil: true)
                .MapMaster("Process", VisioStencils.Flowchart.Get("decision"), resizeToStencil: true)
                .MapMaster("Data", VisioStencils.DataPlatform.Get("database"), resizeToStencil: true)
                .MapNameU("LegacyCustom", VisioStencils.Network.Get("server"), resizeToStencil: true));

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(map);
            loaded.Save(updatedPath);

            Assert.True(result.HasChanges);
            Assert.Equal(4, result.Count);
            Assert.Equal(new[] { "archive", "custom", "intake", "review" }, result.Replacements.Select(replacement => replacement.ShapeId).OrderBy(id => id).ToArray());
            Assert.Contains(result.Replacements, replacement =>
                replacement.ShapeId == "review" &&
                replacement.OldStencilId == "flow.process" &&
                replacement.NewStencilId == "flow.preparation" &&
                replacement.OldMasterNameU == "Process" &&
                replacement.NewMasterNameU == "Preparation");

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage updatedPage = Assert.Single(updated.Pages);
            VisioShape migratedReview = updatedPage.FindShapeById("review")!;
            VisioShape migratedCustom = updatedPage.FindShapeById("custom")!;
            VisioShape migratedArchive = updatedPage.FindShapeById("archive")!;

            Assert.Equal("Preparation", migratedReview.MasterNameU);
            Assert.Equal("flow.preparation", migratedReview.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal(Color.LightYellow, migratedReview.FillColor);
            Assert.Equal("Ops", migratedReview.GetShapeDataValue("Owner"));
            Assert.Equal("Review", migratedReview.GetUserCellValue("Stage"));
            Assert.Equal("https://example.org/review", migratedReview.Hyperlinks.Single().Address);
            Assert.Single(updatedPage.IncomingConnectors(migratedReview));
            Assert.Single(updatedPage.OutgoingConnectors(migratedReview));
            Assert.Equal("Process", migratedCustom.MasterNameU);
            Assert.Equal("net.server", migratedCustom.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("Data", migratedArchive.MasterNameU);
            Assert.Equal("data.database", migratedArchive.GetUserCellValue(VisioSemanticUserCells.StencilId));

            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void FluentCanApplyStencilMigrationToLoadedPage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Fluent Migration", page => page
                    .Stencil("worker", VisioStencils.Flowchart.Get("process"), 2, 5, "Worker")
                    .Stencil("queue", VisioStencils.Flowchart.Get("data"), 5, 5, "Queue")
                    .Rect("legacy", 8, 5, 2, 1, "Legacy")
                    .Shape("legacy", shape => shape.UserCell("Family", "Legacy")))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationMap.Create(builder => builder
                .MapStencilId("flow.process", VisioStencils.Infrastructure.Get("server"), resizeToStencil: true)
                .Map(shape => string.Equals(shape.GetUserCellValue("Family"), "Legacy", StringComparison.OrdinalIgnoreCase),
                    VisioStencils.CollaborationBusiness.Get("system"),
                    resizeToStencil: true));

            VisioDocument.Load(filePath)
                .AsFluent()
                .ExistingPage("Fluent Migration", page => page.ApplyStencilMigration(map))
                .End()
                .Save(updatedPath);

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            VisioShape worker = page.FindShapeById("worker")!;
            VisioShape legacy = page.FindShapeById("legacy")!;
            VisioShape queue = page.FindShapeById("queue")!;

            Assert.Equal("Process", worker.MasterNameU);
            Assert.Equal("infra.server", worker.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("Process", legacy.MasterNameU);
            Assert.Equal("collab.system", legacy.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("Legacy", legacy.GetUserCellValue("Family"));
            Assert.Equal("Data", queue.MasterNameU);

            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void CatalogQueryHelpersResolveMigrationTargets() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Catalog Queries", 11, 8.5);
            page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "worker", 2, 5, "Worker");
            VisioShape custom = new("legacy-system", 5, 5, 2.2, 1, "Legacy System") { NameU = "LegacySystem" };
            page.Shapes.Add(custom);
            VisioShape database = new("legacy-db", 8, 5, 2.2, 1, "Legacy DB") { NameU = "LegacyDataStore" };
            database.SetUserCell("Family", "DataStore");
            page.Shapes.Add(database);
            document.Save();

            VisioStencilMigrationMap map = VisioStencilMigrationMap.Create(builder => builder
                .MapStencilId("flow.process", VisioStencils.Infrastructure, new[] { "host", "server" }, resizeToStencil: true)
                .MapNameU("LegacySystem", VisioStencils.CollaborationBusiness, "business system", resizeToStencil: true)
                .Map(shape => string.Equals(shape.GetUserCellValue("Family"), "DataStore", StringComparison.OrdinalIgnoreCase),
                    VisioStencils.DataPlatform,
                    new[] { "relational", "database" },
                    resizeToStencil: true));

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(map);
            loaded.Save(updatedPath);

            Assert.Equal(3, result.Count);

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage updatedPage = Assert.Single(updated.Pages);
            Assert.Equal("infra.server", updatedPage.FindShapeById("worker")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("collab.system", updatedPage.FindShapeById("legacy-system")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("data.database", updatedPage.FindShapeById("legacy-db")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void BasicFlowchartPresetMigratesUnstenciledShapesOnly() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Legacy Flow", page => page
                    .Ellipse("start", 2, 6, 1.6, 0.7, "Start")
                    .Rect("task", 4, 6, 2, 0.9, "Task")
                    .Diamond("choice", 6.5, 6, 1.6, 1.1, "Choice?")
                    .Parallelogram("input", 8.7, 6, 1.8, 0.8, "Input")
                    .Hexagon("setup", 4, 4, 1.8, 0.8, "Setup")
                    .Stencil("already", VisioStencils.Flowchart.Get("process"), 7, 4, "Already stenciled"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationPresets.BasicFlowchart(VisioStencils.Flowchart);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(map);
            loaded.Save(updatedPath);

            Assert.Equal(5, result.Count);
            Assert.DoesNotContain(result.Replacements, replacement => replacement.ShapeId == "already");

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            Assert.Equal("flow.start-end", page.FindShapeById("start")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("flow.process", page.FindShapeById("task")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("flow.decision", page.FindShapeById("choice")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("flow.data", page.FindShapeById("input")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("flow.preparation", page.FindShapeById("setup")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("flow.process", page.FindShapeById("already")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void PlanStencilMigrationReportsChangesWithoutMutatingDocument() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Plan", page => page
                    .Stencil("worker", VisioStencils.Flowchart.Get("process"), 2, 5, "Worker")
                    .Stencil("queue", VisioStencils.Flowchart.Get("data"), 5, 5, "Queue"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationMap.Create(builder => builder
                .MapStencilId("flow.process", VisioStencils.Infrastructure, "server", resizeToStencil: true));

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage page = Assert.Single(loaded.Pages);

            VisioStencilMigrationPlan documentPlan = loaded.PlanStencilMigration(map);
            VisioStencilMigrationPlan pagePlan = page.PlanStencilMigration(map);
            VisioStencilMigrationPlan selectionPlan = page.SelectByMaster("Process").PlanStencilMigration(map);

            Assert.Equal(1, documentPlan.Count);
            Assert.Equal(1, pagePlan.Count);
            Assert.Equal(1, selectionPlan.Count);
            Assert.True(documentPlan.HasChanges);

            VisioStencilMigrationPlannedReplacement replacement = documentPlan.Replacements.Single();
            Assert.Equal("Plan", replacement.PageName);
            Assert.Equal("worker", replacement.ShapeId);
            Assert.Equal("flow.process", replacement.OldStencilId);
            Assert.Equal("infra.server", replacement.NewStencilId);
            Assert.Equal("Process", replacement.OldMasterNameU);
            Assert.Equal("Process", replacement.NewMasterNameU);
            Assert.Equal(VisioStencilMigrationMatchKind.StencilId, replacement.MatchKind);
            Assert.Equal("flow.process", replacement.MatchValue);
            Assert.Equal("Server", replacement.ReplacementStencilName);
            Assert.Equal("Infrastructure", replacement.ReplacementStencilCategory);
            Assert.True(replacement.ResizeToStencil);

            string report = documentPlan.ToText();
            Assert.Contains("migration.hasChanges=true", report, StringComparison.Ordinal);
            Assert.Contains("migration.count=1", report, StringComparison.Ordinal);
            Assert.Contains("migration.replacement[0].shapeId=worker", report, StringComparison.Ordinal);
            Assert.Contains("migration.replacement[0].newStencilId=infra.server", report, StringComparison.Ordinal);
            Assert.Contains("migration.replacement[0].resizeToStencil=true", report, StringComparison.Ordinal);

            VisioShape worker = page.FindShapeById("worker")!;
            Assert.Equal("flow.process", worker.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("Process", worker.MasterNameU);
            Assert.Equal("flow.data", page.FindShapeById("queue")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
        }

        [Fact]
        public void ApplyStencilMigrationPlanValidatesReviewedPlanAndDisambiguatesPages() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("First", page => page
                    .Stencil("worker", VisioStencils.Flowchart.Get("process"), 2, 5, "First worker"))
                .Page("Second", page => page
                    .Stencil("worker", VisioStencils.Flowchart.Get("process"), 2, 5, "Second worker"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationMap.Create(builder => builder
                .MapStencilId("flow.process", VisioStencils.Infrastructure.Get("server"), resizeToStencil: true));

            VisioDocument reviewCopy = VisioDocument.Load(filePath);
            VisioStencilMigrationPlan plan = reviewCopy.PlanStencilMigration(map);
            string report = plan.ToText();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(plan, map);
            loaded.Save(updatedPath);

            Assert.Equal(2, plan.Count);
            Assert.Equal(2, result.Count);
            Assert.Contains("migration.replacement[0].pageId=", report, StringComparison.Ordinal);
            Assert.Contains("migration.replacement[0].pageNameU=First", report, StringComparison.Ordinal);
            Assert.Contains("migration.replacement[1].pageNameU=Second", report, StringComparison.Ordinal);

            VisioDocument updated = VisioDocument.Load(updatedPath);
            Assert.Equal("infra.server", updated.Pages[0].FindShapeById("worker")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("infra.server", updated.Pages[1].FindShapeById("worker")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void ApplyStencilMigrationPlanFailsWhenShapeChangedAfterReview() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Drift", page => page
                    .Stencil("worker", VisioStencils.Flowchart.Get("process"), 2, 5, "Worker"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationMap.Create(builder => builder
                .MapStencilId("flow.process", VisioStencils.Infrastructure.Get("server"), resizeToStencil: true));

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationPlan plan = loaded.PlanStencilMigration(map);
            loaded.Pages[0].FindShapeById("worker")!.Text = "Worker changed after review";

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => loaded.ApplyStencilMigration(plan, map));

            Assert.Contains("text changed", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal("flow.process", loaded.Pages[0].FindShapeById("worker")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
        }

        [Fact]
        public void NetworkInfrastructurePresetMigratesLegacyNetworkLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Legacy Network", page => page
                    .Rect("server", 2, 6, 1.8, 0.8, "Web Server")
                    .Data("database", 4.5, 6, 1.8, 0.8, "SQL Database")
                    .Diamond("firewall", 7, 6, 1.3, 1.0, "Firewall")
                    .Rect("switch", 2, 4, 1.8, 0.7, "Core Switch")
                    .Ellipse("internet", 4.5, 4, 1.0, 1.0, "Internet")
                    .Stencil("already", VisioStencils.Network.Get("server"), 7, 4, "Already stenciled"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationPresets.NetworkInfrastructure(VisioStencils.Network);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationPlan plan = loaded.PlanStencilMigration(map);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(map);
            loaded.Save(updatedPath);

            Assert.Equal(5, plan.Count);
            Assert.Equal(5, result.Count);
            Assert.DoesNotContain(result.Replacements, replacement => replacement.ShapeId == "already");

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            Assert.Equal("net.server", page.FindShapeById("server")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("net.database", page.FindShapeById("database")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("net.firewall", page.FindShapeById("firewall")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("net.switch", page.FindShapeById("switch")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("net.internet", page.FindShapeById("internet")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("net.server", page.FindShapeById("already")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void ArchitectureInfrastructurePresetMigratesLegacyArchitectureLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Legacy Architecture", page => page
                    .Rect("api", 2, 6, 1.8, 0.8, "API Service")
                    .Rect("worker", 4.2, 6, 1.8, 0.8, "Worker VM")
                    .Data("database", 6.4, 6, 1.8, 0.8, "SQL Database")
                    .Data("queue", 8.6, 6, 1.8, 0.8, "Message Queue")
                    .Diamond("gateway", 2, 4, 1.4, 1.0, "API Gateway")
                    .Rect("region", 5.5, 3.7, 3.2, 1.4, "Region Boundary")
                    .Stencil("already", VisioStencils.Architecture.Get("service"), 9, 4, "Already stenciled"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationPresets.ArchitectureInfrastructure(VisioStencils.Architecture);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationPlan plan = loaded.PlanStencilMigration(map);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(map);
            loaded.Save(updatedPath);

            Assert.Equal(6, plan.Count);
            Assert.Equal(6, result.Count);
            Assert.DoesNotContain(result.Replacements, replacement => replacement.ShapeId == "already");

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            Assert.Equal("arch.service", page.FindShapeById("api")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("arch.compute", page.FindShapeById("worker")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("arch.database", page.FindShapeById("database")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("arch.queue", page.FindShapeById("queue")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("arch.gateway", page.FindShapeById("gateway")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("arch.region", page.FindShapeById("region")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("arch.service", page.FindShapeById("already")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void OrgChartPresetMigratesLegacyOrganizationLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Legacy Org", page => page
                    .Rect("ceo", 2, 6, 1.8, 0.8, "CEO Executive")
                    .Rect("manager", 4.2, 6, 1.8, 0.8, "Engineering Manager")
                    .Rect("employee", 6.4, 6, 1.8, 0.8, "Platform Engineer")
                    .Rect("assistant", 2, 4.5, 1.6, 0.65, "Executive Assistant")
                    .Rect("vacancy", 4.2, 4.5, 1.8, 0.8, "Open Position")
                    .Rect("external", 6.4, 4.5, 1.8, 0.8, "External Advisor")
                    .Rect("band", 4.2, 3.2, 3.2, 1.2, "Platform Team Band")
                    .Stencil("already", VisioStencils.OrgChart.Get("position"), 8.6, 6, "Already stenciled"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationPresets.OrgChart(VisioStencils.OrgChart);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationPlan plan = loaded.PlanStencilMigration(map);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(map);
            loaded.Save(updatedPath);

            Assert.Equal(7, plan.Count);
            Assert.Equal(7, result.Count);
            Assert.DoesNotContain(result.Replacements, replacement => replacement.ShapeId == "already");

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            Assert.Equal("org.executive", page.FindShapeById("ceo")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("org.manager", page.FindShapeById("manager")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("org.position", page.FindShapeById("employee")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("org.assistant", page.FindShapeById("assistant")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("org.vacancy", page.FindShapeById("vacancy")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("org.external", page.FindShapeById("external")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("org.team-band", page.FindShapeById("band")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("org.position", page.FindShapeById("already")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void TimelinePresetMigratesLegacyRoadmapLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Legacy Timeline", page => page
                    .Rect("axis", 5, 6, 7.0, 0.08, "Timeline Axis")
                    .Diamond("milestone", 2, 5, 0.4, 0.4, "Milestone")
                    .Ellipse("release", 3.5, 5, 0.4, 0.4, "Release Launch")
                    .Ellipse("decision", 5, 5, 0.4, 0.4, "Approval Gate")
                    .Ellipse("risk", 6.5, 5, 0.4, 0.4, "Risk Issue")
                    .Rect("span", 4.2, 4, 2.6, 0.4, "Phase Duration")
                    .Rect("label", 7.6, 4.2, 1.4, 0.5, "Callout Label")
                    .Stencil("already", VisioStencils.Timeline.Get("milestone"), 8.5, 5, "Already stenciled"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationPresets.Timeline(VisioStencils.Timeline);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationPlan plan = loaded.PlanStencilMigration(map);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(map);
            loaded.Save(updatedPath);

            Assert.Equal(7, plan.Count);
            Assert.Equal(7, result.Count);
            Assert.DoesNotContain(result.Replacements, replacement => replacement.ShapeId == "already");

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            Assert.Equal("time.axis", page.FindShapeById("axis")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("time.milestone", page.FindShapeById("milestone")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("time.release", page.FindShapeById("release")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("time.decision", page.FindShapeById("decision")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("time.risk", page.FindShapeById("risk")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("time.span", page.FindShapeById("span")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("time.label", page.FindShapeById("label")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("time.milestone", page.FindShapeById("already")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void SequencePresetMigratesLegacySequenceLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Legacy Sequence", page => page
                    .Ellipse("actor", 1.5, 6, 0.8, 0.8, "Client Actor")
                    .Rect("participant", 3, 6, 1.4, 0.7, "Payment Service")
                    .Rect("boundary", 4.8, 6, 1.4, 0.7, "API Boundary")
                    .Rect("control", 6.6, 6, 1.4, 0.7, "Retry Controller")
                    .Rect("entity", 8.4, 6, 1.4, 0.7, "Domain Object")
                    .Data("database", 3, 4.6, 1.4, 0.7, "Order Store")
                    .Rect("activation", 4.8, 4.6, 0.3, 1.1, "Activation")
                    .Rect("fragment", 6.6, 4.4, 2.2, 1.2, "alt retry fragment")
                    .Rect("note", 8.8, 4.4, 1.6, 0.7, "Sequence Note")
                    .Stencil("already", VisioStencils.Sequence.Get("participant"), 1.5, 4.4, "Already stenciled"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationPresets.Sequence(VisioStencils.Sequence);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationPlan plan = loaded.PlanStencilMigration(map);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(map);
            loaded.Save(updatedPath);

            Assert.Equal(9, plan.Count);
            Assert.Equal(9, result.Count);
            Assert.DoesNotContain(result.Replacements, replacement => replacement.ShapeId == "already");

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            Assert.Equal("seq.actor", page.FindShapeById("actor")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("seq.participant", page.FindShapeById("participant")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("seq.boundary", page.FindShapeById("boundary")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("seq.control", page.FindShapeById("control")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("seq.entity", page.FindShapeById("entity")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("seq.database", page.FindShapeById("database")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("seq.activation", page.FindShapeById("activation")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("seq.fragment", page.FindShapeById("fragment")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("seq.note", page.FindShapeById("note")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("seq.participant", page.FindShapeById("already")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void SwimlaneProcessMapPresetMigratesLegacyProcessLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Legacy Swimlane", page => page
                    .Rect("lane", 4.5, 6, 6.0, 1.1, "Responsibility Lane")
                    .Rect("phase", 2, 7.2, 2.0, 0.5, "Phase Stage")
                    .Rect("activity", 2, 5.8, 1.6, 0.7, "Review Task")
                    .Diamond("decision", 4, 5.8, 1.2, 0.9, "Decision Choice")
                    .Data("data", 6, 5.8, 1.6, 0.7, "Input Document")
                    .Ellipse("start", 8, 5.8, 1.4, 0.7, "Start")
                    .Stencil("already", VisioStencils.Swimlane.Get("activity"), 8, 7.1, "Already stenciled"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationPresets.SwimlaneProcessMap(VisioStencils.Swimlane);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationPlan plan = loaded.PlanStencilMigration(map);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(map);
            loaded.Save(updatedPath);

            Assert.Equal(6, plan.Count);
            Assert.Equal(6, result.Count);
            Assert.DoesNotContain(result.Replacements, replacement => replacement.ShapeId == "already");

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            Assert.Equal("swim.lane", page.FindShapeById("lane")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("swim.phase", page.FindShapeById("phase")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("swim.activity", page.FindShapeById("activity")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("swim.decision", page.FindShapeById("decision")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("swim.data", page.FindShapeById("data")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("swim.start-end", page.FindShapeById("start")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("swim.activity", page.FindShapeById("already")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void CloudInfrastructurePresetMigratesLegacyCloudLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Legacy Cloud", page => page
                    .Rect("subscription", 4.5, 6.4, 6.0, 1.1, "Subscription Boundary")
                    .Rect("region", 2, 5, 1.8, 0.8, "Region Zone")
                    .Rect("service", 4, 5, 1.8, 0.8, "Cloud Service")
                    .Hexagon("function", 6, 5, 1.4, 0.8, "Serverless Function")
                    .Diamond("gateway", 8, 5, 1.2, 0.9, "API Gateway")
                    .Data("queue", 2, 3.8, 1.6, 0.7, "Message Queue")
                    .Data("secret", 4, 3.8, 1.6, 0.7, "Secret Store")
                    .Ellipse("monitoring", 6, 3.8, 0.9, 0.9, "Monitoring Logs")
                    .Stencil("already", VisioStencils.Cloud.Get("service"), 8, 3.8, "Already stenciled"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationPresets.CloudInfrastructure(VisioStencils.Cloud);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationPlan plan = loaded.PlanStencilMigration(map);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(map);
            loaded.Save(updatedPath);

            Assert.Equal(8, plan.Count);
            Assert.Equal(8, result.Count);
            Assert.DoesNotContain(result.Replacements, replacement => replacement.ShapeId == "already");

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            Assert.Equal("cloud.subscription", page.FindShapeById("subscription")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("cloud.region", page.FindShapeById("region")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("cloud.service", page.FindShapeById("service")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("cloud.function", page.FindShapeById("function")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("cloud.gateway", page.FindShapeById("gateway")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("cloud.queue", page.FindShapeById("queue")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("cloud.secret-store", page.FindShapeById("secret")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("cloud.monitoring", page.FindShapeById("monitoring")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("cloud.service", page.FindShapeById("already")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void SecurityIdentityPresetMigratesLegacySecurityLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Legacy Security", page => page
                    .Rect("idp", 2, 6, 1.8, 0.8, "Identity Provider")
                    .Ellipse("user", 4, 6, 0.8, 0.8, "User Principal")
                    .Rect("group", 6, 6, 1.5, 0.8, "Security Group")
                    .Diamond("policy", 8, 6, 1.3, 0.9, "Conditional Access Policy")
                    .Diamond("key", 2, 4.5, 0.9, 0.75, "Certificate Secret")
                    .Diamond("firewall", 4, 4.5, 1.3, 0.9, "Firewall Policy")
                    .Data("audit", 6, 4.5, 1.6, 0.8, "Audit Log")
                    .Rect("boundary", 8, 4.5, 2.0, 1.0, "Trust Boundary")
                    .Triangle("alert", 2, 3.2, 0.9, 0.8, "Security Alert")
                    .Stencil("already", VisioStencils.SecurityIdentity.Get("user"), 4, 3.2, "Already stenciled"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationPresets.SecurityIdentity(VisioStencils.SecurityIdentity);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationPlan plan = loaded.PlanStencilMigration(map);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(map);
            loaded.Save(updatedPath);

            Assert.Equal(9, plan.Count);
            Assert.Equal(9, result.Count);
            Assert.DoesNotContain(result.Replacements, replacement => replacement.ShapeId == "already");

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            Assert.Equal("sec.identity-provider", page.FindShapeById("idp")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("sec.user", page.FindShapeById("user")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("sec.group", page.FindShapeById("group")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("sec.policy", page.FindShapeById("policy")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("sec.key", page.FindShapeById("key")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("sec.firewall", page.FindShapeById("firewall")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("sec.audit", page.FindShapeById("audit")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("sec.trust-boundary", page.FindShapeById("boundary")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("sec.alert", page.FindShapeById("alert")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("sec.user", page.FindShapeById("already")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Empty(VisioValidator.Validate(updatedPath));
        }
    }
}
