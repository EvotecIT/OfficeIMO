using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using OfficeIMO.Visio.Stencils;
using Xunit;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Tests {
    public class VisioStencilMigrationTests {
        [Fact]
        public void StencilMigrationMapAppliesFirstMatchingRulesAndReportsChanges() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Migration", OfficeLength.Inches(11), OfficeLength.Inches(8.5));
            VisioShape intake = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "intake", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), "Intake");
            VisioShape review = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "review", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), "Review");
            VisioShape archive = page.AddStencilShape(VisioStencils.Flowchart.Get("data"), "archive", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(8, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(6, page.DefaultUnit)), "Archive");
            VisioShape custom = new("custom", OfficeLength.Inches(5), OfficeLength.Inches(3), OfficeLength.Inches(2.5), OfficeLength.Inches(1), "Custom") { NameU = "LegacyCustom" };
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
                .AddPage("Fluent Migration", page => page
                    .AddStencilShape(id: "worker", stencil: VisioStencils.Flowchart.Get("process"), x: OfficeLength.Inches(2), y: OfficeLength.Inches(5), text: "Worker")
                    .AddStencilShape(id: "queue", stencil: VisioStencils.Flowchart.Get("data"), x: OfficeLength.Inches(5), y: OfficeLength.Inches(5), text: "Queue")
                    .AddRectangle("legacy", OfficeLength.Inches(8), OfficeLength.Inches(5), OfficeLength.Inches(2), OfficeLength.Inches(1), "Legacy")
                    .GetShape("legacy", shape => shape.UserCell("Family", "Legacy")))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationMap.Create(builder => builder
                .MapStencilId("flow.process", VisioStencils.Infrastructure.Get("server"), resizeToStencil: true)
                .Map(shape => string.Equals(shape.GetUserCellValue("Family"), "Legacy", StringComparison.OrdinalIgnoreCase),
                    VisioStencils.CollaborationBusiness.Get("system"),
                    resizeToStencil: true));

            VisioDocument.Load(filePath)
                .AsFluent()
                .GetPage("Fluent Migration", page => page.ApplyStencilMigration(map))
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
            VisioPage page = document.AddPage("Catalog Queries", OfficeLength.Inches(11), OfficeLength.Inches(8.5));
            page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "worker", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(2, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), "Worker");
            VisioShape custom = new("legacy-system", OfficeLength.Inches(5), OfficeLength.Inches(5), OfficeLength.Inches(2.2), OfficeLength.Inches(1), "Legacy System") { NameU = "LegacySystem" };
            page.Shapes.Add(custom);
            VisioShape database = new("legacy-db", OfficeLength.Inches(8), OfficeLength.Inches(5), OfficeLength.Inches(2.2), OfficeLength.Inches(1), "Legacy DB") { NameU = "LegacyDataStore" };
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
                .AddPage("Legacy Flow", page => page
                    .AddEllipse("start", OfficeLength.Inches(2), OfficeLength.Inches(6), OfficeLength.Inches(1.6), OfficeLength.Inches(0.7), "Start")
                    .AddRectangle("task", OfficeLength.Inches(4), OfficeLength.Inches(6), OfficeLength.Inches(2), OfficeLength.Inches(0.9), "Task")
                    .AddDiamond("choice", OfficeLength.Inches(6.5), OfficeLength.Inches(6), OfficeLength.Inches(1.6), OfficeLength.Inches(1.1), "Choice?")
                    .AddParallelogram("input", OfficeLength.Inches(8.7), OfficeLength.Inches(6), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Input")
                    .AddHexagon("setup", OfficeLength.Inches(4), OfficeLength.Inches(4), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Setup")
                    .AddStencilShape(id: "already", stencil: VisioStencils.Flowchart.Get("process"), x: OfficeLength.Inches(7), y: OfficeLength.Inches(4), text: "Already stenciled"))
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
                .AddPage("Plan", page => page
                    .AddStencilShape(id: "worker", stencil: VisioStencils.Flowchart.Get("process"), x: OfficeLength.Inches(2), y: OfficeLength.Inches(5), text: "Worker")
                    .AddStencilShape(id: "queue", stencil: VisioStencils.Flowchart.Get("data"), x: OfficeLength.Inches(5), y: OfficeLength.Inches(5), text: "Queue"))
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
        public void DocumentStencilMigrationSkipsBackgroundPages() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage background = document.AddBackgroundPage("Watermark", OfficeLength.Inches(11), OfficeLength.Inches(8.5));
            background.AddStencilShape(VisioStencils.Flowchart.Get("process"), "background-process", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, background.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(7, background.DefaultUnit)), "Background process");
            VisioPage page = document.AddPage("Foreground", OfficeLength.Inches(11), OfficeLength.Inches(8.5));
            page.SetBackgroundPage(background);
            page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "foreground-process", OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(3, page.DefaultUnit)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(5, page.DefaultUnit)), "Foreground process");
            document.Save();

            VisioStencilMigrationMap map = VisioStencilMigrationMap.Create(builder => builder
                .MapStencilId("flow.process", VisioStencils.Infrastructure.Get("server"), resizeToStencil: true));

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationPlan plan = loaded.PlanStencilMigration(map);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(map);

            Assert.Single(plan.Replacements);
            Assert.Single(result.Replacements);
            Assert.Equal("foreground-process", plan.Replacements.Single().ShapeId);
            Assert.Equal("flow.process", loaded.Pages.Single(current => current.IsBackground).FindShapeById("background-process")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("infra.server", loaded.Pages.Single(current => !current.IsBackground).FindShapeById("foreground-process")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
        }

        [Fact]
        public void ApplyStencilMigrationPlanValidatesReviewedPlanAndDisambiguatesPages() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .AddPage("First", page => page
                    .AddStencilShape(id: "worker", stencil: VisioStencils.Flowchart.Get("process"), x: OfficeLength.Inches(2), y: OfficeLength.Inches(5), text: "First worker"))
                .AddPage("Second", page => page
                    .AddStencilShape(id: "worker", stencil: VisioStencils.Flowchart.Get("process"), x: OfficeLength.Inches(2), y: OfficeLength.Inches(5), text: "Second worker"))
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
                .AddPage("Drift", page => page
                    .AddStencilShape(id: "worker", stencil: VisioStencils.Flowchart.Get("process"), x: OfficeLength.Inches(2), y: OfficeLength.Inches(5), text: "Worker"))
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
        public void ApplyStencilMigrationPlanValidatesAllShapesBeforeMutating() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .AddPage("Atomic", page => page
                    .AddStencilShape(id: "first", stencil: VisioStencils.Flowchart.Get("process"), x: OfficeLength.Inches(2), y: OfficeLength.Inches(5), text: "First")
                    .AddStencilShape(id: "second", stencil: VisioStencils.Flowchart.Get("process"), x: OfficeLength.Inches(5), y: OfficeLength.Inches(5), text: "Second"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationMap.Create(builder => builder
                .MapStencilId("flow.process", VisioStencils.Infrastructure.Get("server"), resizeToStencil: true));

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationPlan plan = loaded.PlanStencilMigration(map);
            loaded.Pages[0].FindShapeById("second")!.Text = "Second changed after review";

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => loaded.ApplyStencilMigration(plan, map));

            Assert.Contains("text changed", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal("flow.process", loaded.Pages[0].FindShapeById("first")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("flow.process", loaded.Pages[0].FindShapeById("second")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
        }

        [Fact]
        public void NetworkInfrastructurePresetMigratesLegacyNetworkLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .AddPage("Legacy Network", page => page
                    .AddRectangle("server", OfficeLength.Inches(2), OfficeLength.Inches(6), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Web Server")
                    .AddData("database", OfficeLength.Inches(4.5), OfficeLength.Inches(6), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "SQL Database")
                    .AddDiamond("firewall", OfficeLength.Inches(7), OfficeLength.Inches(6), OfficeLength.Inches(1.3), OfficeLength.Inches(1.0), "Firewall")
                    .AddRectangle("switch", OfficeLength.Inches(2), OfficeLength.Inches(4), OfficeLength.Inches(1.8), OfficeLength.Inches(0.7), "Core Switch")
                    .AddEllipse("internet", OfficeLength.Inches(4.5), OfficeLength.Inches(4), OfficeLength.Inches(1.0), OfficeLength.Inches(1.0), "Internet")
                    .AddStencilShape(id: "already", stencil: VisioStencils.Network.Get("server"), x: OfficeLength.Inches(7), y: OfficeLength.Inches(4), text: "Already stenciled"))
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
                .AddPage("Legacy Architecture", page => page
                    .AddRectangle("api", OfficeLength.Inches(2), OfficeLength.Inches(6), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "API Service")
                    .AddRectangle("worker", OfficeLength.Inches(4.2), OfficeLength.Inches(6), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Worker VM")
                    .AddData("database", OfficeLength.Inches(6.4), OfficeLength.Inches(6), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "SQL Database")
                    .AddData("queue", OfficeLength.Inches(8.6), OfficeLength.Inches(6), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Message Queue")
                    .AddDiamond("gateway", OfficeLength.Inches(2), OfficeLength.Inches(4), OfficeLength.Inches(1.4), OfficeLength.Inches(1.0), "API Gateway")
                    .AddRectangle("region", OfficeLength.Inches(5.5), OfficeLength.Inches(3.7), OfficeLength.Inches(3.2), OfficeLength.Inches(1.4), "Region Boundary")
                    .AddStencilShape(id: "already", stencil: VisioStencils.Architecture.Get("service"), x: OfficeLength.Inches(9), y: OfficeLength.Inches(4), text: "Already stenciled"))
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
                .AddPage("Legacy Org", page => page
                    .AddRectangle("ceo", OfficeLength.Inches(2), OfficeLength.Inches(6), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "CEO Executive")
                    .AddRectangle("manager", OfficeLength.Inches(4.2), OfficeLength.Inches(6), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Engineering Manager")
                    .AddRectangle("employee", OfficeLength.Inches(6.4), OfficeLength.Inches(6), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Platform Engineer")
                    .AddRectangle("assistant", OfficeLength.Inches(2), OfficeLength.Inches(4.5), OfficeLength.Inches(1.6), OfficeLength.Inches(0.65), "Executive Assistant")
                    .AddRectangle("vacancy", OfficeLength.Inches(4.2), OfficeLength.Inches(4.5), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Open Position")
                    .AddRectangle("external", OfficeLength.Inches(6.4), OfficeLength.Inches(4.5), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "External Advisor")
                    .AddRectangle("band", OfficeLength.Inches(4.2), OfficeLength.Inches(3.2), OfficeLength.Inches(3.2), OfficeLength.Inches(1.2), "Platform Team Band")
                    .AddStencilShape(id: "already", stencil: VisioStencils.OrgChart.Get("position"), x: OfficeLength.Inches(8.6), y: OfficeLength.Inches(6), text: "Already stenciled"))
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
                .AddPage("Legacy Timeline", page => page
                    .AddRectangle("axis", OfficeLength.Inches(5), OfficeLength.Inches(6), OfficeLength.Inches(7.0), OfficeLength.Inches(0.08), "Timeline Axis")
                    .AddDiamond("milestone", OfficeLength.Inches(2), OfficeLength.Inches(5), OfficeLength.Inches(0.4), OfficeLength.Inches(0.4), "Milestone")
                    .AddEllipse("release", OfficeLength.Inches(3.5), OfficeLength.Inches(5), OfficeLength.Inches(0.4), OfficeLength.Inches(0.4), "Release Launch")
                    .AddEllipse("decision", OfficeLength.Inches(5), OfficeLength.Inches(5), OfficeLength.Inches(0.4), OfficeLength.Inches(0.4), "Approval Gate")
                    .AddEllipse("risk", OfficeLength.Inches(6.5), OfficeLength.Inches(5), OfficeLength.Inches(0.4), OfficeLength.Inches(0.4), "Risk Issue")
                    .AddRectangle("span", OfficeLength.Inches(4.2), OfficeLength.Inches(4), OfficeLength.Inches(2.6), OfficeLength.Inches(0.4), "Phase Duration")
                    .AddRectangle("label", OfficeLength.Inches(7.6), OfficeLength.Inches(4.2), OfficeLength.Inches(1.4), OfficeLength.Inches(0.5), "Callout Label")
                    .AddStencilShape(id: "already", stencil: VisioStencils.Timeline.Get("milestone"), x: OfficeLength.Inches(8.5), y: OfficeLength.Inches(5), text: "Already stenciled"))
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
                .AddPage("Legacy Sequence", page => page
                    .AddEllipse("actor", OfficeLength.Inches(1.5), OfficeLength.Inches(6), OfficeLength.Inches(0.8), OfficeLength.Inches(0.8), "Client Actor")
                    .AddRectangle("participant", OfficeLength.Inches(3), OfficeLength.Inches(6), OfficeLength.Inches(1.4), OfficeLength.Inches(0.7), "Payment Service")
                    .AddRectangle("boundary", OfficeLength.Inches(4.8), OfficeLength.Inches(6), OfficeLength.Inches(1.4), OfficeLength.Inches(0.7), "API Boundary")
                    .AddRectangle("control", OfficeLength.Inches(6.6), OfficeLength.Inches(6), OfficeLength.Inches(1.4), OfficeLength.Inches(0.7), "Retry Controller")
                    .AddRectangle("entity", OfficeLength.Inches(8.4), OfficeLength.Inches(6), OfficeLength.Inches(1.4), OfficeLength.Inches(0.7), "Domain Object")
                    .AddData("database", OfficeLength.Inches(3), OfficeLength.Inches(4.6), OfficeLength.Inches(1.4), OfficeLength.Inches(0.7), "Order Store")
                    .AddRectangle("activation", OfficeLength.Inches(4.8), OfficeLength.Inches(4.6), OfficeLength.Inches(0.3), OfficeLength.Inches(1.1), "Activation")
                    .AddRectangle("fragment", OfficeLength.Inches(6.6), OfficeLength.Inches(4.4), OfficeLength.Inches(2.2), OfficeLength.Inches(1.2), "alt retry fragment")
                    .AddRectangle("note", OfficeLength.Inches(8.8), OfficeLength.Inches(4.4), OfficeLength.Inches(1.6), OfficeLength.Inches(0.7), "Sequence Note")
                    .AddStencilShape(id: "already", stencil: VisioStencils.Sequence.Get("participant"), x: OfficeLength.Inches(1.5), y: OfficeLength.Inches(4.4), text: "Already stenciled"))
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
                .AddPage("Legacy Swimlane", page => page
                    .AddRectangle("lane", OfficeLength.Inches(4.5), OfficeLength.Inches(6), OfficeLength.Inches(6.0), OfficeLength.Inches(1.1), "Responsibility Lane")
                    .AddRectangle("phase", OfficeLength.Inches(2), OfficeLength.Inches(7.2), OfficeLength.Inches(2.0), OfficeLength.Inches(0.5), "Phase Stage")
                    .AddRectangle("activity", OfficeLength.Inches(2), OfficeLength.Inches(5.8), OfficeLength.Inches(1.6), OfficeLength.Inches(0.7), "Review Task")
                    .AddDiamond("decision", OfficeLength.Inches(4), OfficeLength.Inches(5.8), OfficeLength.Inches(1.2), OfficeLength.Inches(0.9), "Decision Choice")
                    .AddData("data", OfficeLength.Inches(6), OfficeLength.Inches(5.8), OfficeLength.Inches(1.6), OfficeLength.Inches(0.7), "Input Document")
                    .AddEllipse("start", OfficeLength.Inches(8), OfficeLength.Inches(5.8), OfficeLength.Inches(1.4), OfficeLength.Inches(0.7), "Start")
                    .AddStencilShape(id: "already", stencil: VisioStencils.Swimlane.Get("activity"), x: OfficeLength.Inches(8), y: OfficeLength.Inches(7.1), text: "Already stenciled"))
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
                .AddPage("Legacy Cloud", page => page
                    .AddRectangle("subscription", OfficeLength.Inches(4.5), OfficeLength.Inches(6.4), OfficeLength.Inches(6.0), OfficeLength.Inches(1.1), "Subscription Boundary")
                    .AddRectangle("region", OfficeLength.Inches(2), OfficeLength.Inches(5), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Region Zone")
                    .AddRectangle("service", OfficeLength.Inches(4), OfficeLength.Inches(5), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Cloud Service")
                    .AddHexagon("function", OfficeLength.Inches(6), OfficeLength.Inches(5), OfficeLength.Inches(1.4), OfficeLength.Inches(0.8), "Serverless Function")
                    .AddDiamond("gateway", OfficeLength.Inches(8), OfficeLength.Inches(5), OfficeLength.Inches(1.2), OfficeLength.Inches(0.9), "API Gateway")
                    .AddData("queue", OfficeLength.Inches(2), OfficeLength.Inches(3.8), OfficeLength.Inches(1.6), OfficeLength.Inches(0.7), "Message Queue")
                    .AddData("secret", OfficeLength.Inches(4), OfficeLength.Inches(3.8), OfficeLength.Inches(1.6), OfficeLength.Inches(0.7), "Secret Store")
                    .AddEllipse("monitoring", OfficeLength.Inches(6), OfficeLength.Inches(3.8), OfficeLength.Inches(0.9), OfficeLength.Inches(0.9), "Monitoring Logs")
                    .AddStencilShape(id: "already", stencil: VisioStencils.Cloud.Get("service"), x: OfficeLength.Inches(8), y: OfficeLength.Inches(3.8), text: "Already stenciled"))
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
                .AddPage("Legacy Security", page => page
                    .AddRectangle("idp", OfficeLength.Inches(2), OfficeLength.Inches(6), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Identity Provider")
                    .AddEllipse("user", OfficeLength.Inches(4), OfficeLength.Inches(6), OfficeLength.Inches(0.8), OfficeLength.Inches(0.8), "User Principal")
                    .AddRectangle("group", OfficeLength.Inches(6), OfficeLength.Inches(6), OfficeLength.Inches(1.5), OfficeLength.Inches(0.8), "Security Group")
                    .AddDiamond("policy", OfficeLength.Inches(8), OfficeLength.Inches(6), OfficeLength.Inches(1.3), OfficeLength.Inches(0.9), "Conditional Access Policy")
                    .AddDiamond("key", OfficeLength.Inches(2), OfficeLength.Inches(4.5), OfficeLength.Inches(0.9), OfficeLength.Inches(0.75), "Certificate Secret")
                    .AddDiamond("firewall", OfficeLength.Inches(4), OfficeLength.Inches(4.5), OfficeLength.Inches(1.3), OfficeLength.Inches(0.9), "Firewall Policy")
                    .AddData("audit", OfficeLength.Inches(6), OfficeLength.Inches(4.5), OfficeLength.Inches(1.6), OfficeLength.Inches(0.8), "Audit Log")
                    .AddRectangle("boundary", OfficeLength.Inches(8), OfficeLength.Inches(4.5), OfficeLength.Inches(2.0), OfficeLength.Inches(1.0), "Trust Boundary")
                    .AddTriangle("alert", OfficeLength.Inches(2), OfficeLength.Inches(3.2), OfficeLength.Inches(0.9), OfficeLength.Inches(0.8), "Security Alert")
                    .AddStencilShape(id: "already", stencil: VisioStencils.SecurityIdentity.Get("user"), x: OfficeLength.Inches(4), y: OfficeLength.Inches(3.2), text: "Already stenciled"))
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
