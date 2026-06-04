using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioStencilMigrationArtifactTests {
        [Fact]
        public void SavedMigrationPlanArtifactCanBeLoadedAndAppliedLater() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string planPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".txt");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("First", page => page
                    .Stencil("worker", VisioStencils.Flowchart.Get("process"), 2, 5, "Worker A=B\\C"))
                .Page("Second", page => page
                    .Stencil("queue", VisioStencils.Flowchart.Get("data"), 2, 5, "Queue\nLine 2"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationMap.Create(builder => builder
                .MapStencilId("flow.process", VisioStencils.Infrastructure.Get("server"), resizeToStencil: true)
                .MapStencilId("flow.data", VisioStencils.DataPlatform.Get("stream"), resizeToStencil: true));

            VisioStencilMigrationPlan reviewPlan = VisioDocument.Load(filePath).PlanStencilMigration(map);
            reviewPlan.SaveText(planPath);
            string artifactText = File.ReadAllText(planPath);
            Assert.Contains("migration.artifactVersion=1", artifactText, StringComparison.Ordinal);
            Assert.Contains("migration.replacement[0].text.isNull=false", artifactText, StringComparison.Ordinal);
            Assert.Contains(@"Worker A=B\\C", artifactText, StringComparison.Ordinal);
            Assert.Contains(@"Queue\nLine 2", artifactText, StringComparison.Ordinal);

            VisioStencilMigrationPlan approvedPlan = VisioStencilMigrationPlan.LoadText(planPath);
            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(approvedPlan, map);
            loaded.Save(updatedPath);

            Assert.Equal(2, approvedPlan.Count);
            Assert.Equal(2, result.Count);

            VisioDocument updated = VisioDocument.Load(updatedPath);
            Assert.Equal("infra.server", updated.Pages[0].FindShapeById("worker")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("data.stream", updated.Pages[1].FindShapeById("queue")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void MigrationPlanArtifactPreservesShapeTextNullStateForDriftValidation() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string planPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".txt");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Empty Text", page => page
                    .Rect("empty", 2, 5, 1.6, 0.7, string.Empty))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationMap.Create(builder => builder
                .MapMaster("Rectangle", VisioStencils.Flowchart.Get("process"), resizeToStencil: true));

            VisioStencilMigrationPlan plan = VisioDocument.Load(filePath).PlanStencilMigration(map);
            plan.SaveText(planPath);
            VisioStencilMigrationPlan loadedPlan = VisioStencilMigrationPlan.LoadText(planPath);
            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(loadedPlan, map);

            Assert.Single(loadedPlan.Replacements);
            Assert.Null(loadedPlan.Replacements[0].Text);
            Assert.Equal(1, result.Count);
            Assert.Equal("flow.process", loaded.Pages[0].FindShapeById("empty")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
        }

        [Fact]
        public void MigrationPlanArtifactPreservesExplicitEmptyShapeText() {
            VisioStencilMigrationPlan plan = new(new[] {
                new VisioStencilMigrationPlannedReplacement(
                    pageId: 1,
                    pageName: "Page",
                    pageNameU: "Page",
                    shapeId: "shape",
                    text: string.Empty,
                    matchKind: VisioStencilMigrationMatchKind.MasterNameU,
                    matchValue: "Rectangle",
                    oldMasterNameU: "Rectangle",
                    newMasterNameU: "Process",
                    oldStencilId: null,
                    newStencilId: "flow.process",
                    replacementStencilName: "Process",
                    replacementStencilCategory: "Flowchart",
                    resizeToStencil: true)
            });

            VisioStencilMigrationPlan loaded = VisioStencilMigrationPlan.FromText(plan.ToText());

            Assert.Single(loaded.Replacements);
            Assert.Equal(string.Empty, loaded.Replacements[0].Text);
            Assert.Null(loaded.Replacements[0].OldStencilId);
        }

        [Fact]
        public void MigrationPlanArtifactRejectsMalformedCount() {
            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                VisioStencilMigrationPlan.FromText("migration.count=not-a-number"));

            Assert.Contains("migration.count", exception.Message, StringComparison.Ordinal);
        }
    }
}
