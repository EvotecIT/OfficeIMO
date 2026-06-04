using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioStencilMigrationDomainPresetTests {
        [Fact]
        public void ContainersKubernetesPresetMigratesLegacyPlatformLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Legacy Kubernetes", page => page
                    .Rect("cluster", 5, 6.5, 5.0, 1.1, "AKS Cluster")
                    .Rect("namespace", 2, 5.2, 1.8, 0.8, "Tenant Namespace")
                    .Rect("node", 4, 5.2, 1.8, 0.8, "Worker Node")
                    .Hexagon("pod", 6, 5.2, 1.4, 0.8, "Workload Pod")
                    .Rect("container", 8, 5.2, 1.6, 0.8, "Container Image")
                    .Rect("service", 2, 4.0, 1.8, 0.8, "Kubernetes Service")
                    .Diamond("ingress", 4, 4.0, 1.3, 0.9, "Ingress Route")
                    .Data("config", 6, 4.0, 1.6, 0.7, "Config Map")
                    .Data("secret", 8, 4.0, 1.6, 0.7, "Secret Credential")
                    .Stencil("already", VisioStencils.ContainersKubernetes.Get("pod"), 5, 2.8, "Already stenciled"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationPresets.ContainersKubernetes(VisioStencils.ContainersKubernetes);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationPlan plan = loaded.PlanStencilMigration(map);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(map);
            loaded.Save(updatedPath);

            Assert.Equal(9, plan.Count);
            Assert.Equal(9, result.Count);
            Assert.DoesNotContain(result.Replacements, replacement => replacement.ShapeId == "already");

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            Assert.Equal("k8s.cluster", page.FindShapeById("cluster")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("k8s.namespace", page.FindShapeById("namespace")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("k8s.node", page.FindShapeById("node")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("k8s.pod", page.FindShapeById("pod")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("k8s.container", page.FindShapeById("container")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("k8s.service", page.FindShapeById("service")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("k8s.ingress", page.FindShapeById("ingress")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("k8s.config", page.FindShapeById("config")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("k8s.secret", page.FindShapeById("secret")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("k8s.pod", page.FindShapeById("already")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void DataPlatformPresetMigratesLegacyDataLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Legacy Data Platform", page => page
                    .Data("lake", 2, 6, 1.8, 0.8, "Lake Storage")
                    .Data("warehouse", 4, 6, 1.8, 0.8, "Data Warehouse")
                    .Data("stream", 6, 6, 1.7, 0.75, "Event Stream")
                    .Rect("pipeline", 8, 6, 1.7, 0.75, "ETL Pipeline")
                    .Rect("catalog", 2, 4.6, 1.8, 0.8, "Metadata Catalog")
                    .Rect("api", 4, 4.6, 1.6, 0.8, "Query API")
                    .Diamond("quality", 6, 4.6, 1.3, 0.9, "Quality Gate")
                    .Data("database", 8, 4.6, 1.7, 0.8, "SQL Database")
                    .Stencil("already", VisioStencils.DataPlatform.Get("database"), 5, 3.2, "Already stenciled"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationPresets.DataPlatform(VisioStencils.DataPlatform);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationPlan plan = loaded.PlanStencilMigration(map);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(map);
            loaded.Save(updatedPath);

            Assert.Equal(8, plan.Count);
            Assert.Equal(8, result.Count);
            Assert.DoesNotContain(result.Replacements, replacement => replacement.ShapeId == "already");

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            Assert.Equal("data.lake", page.FindShapeById("lake")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("data.warehouse", page.FindShapeById("warehouse")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("data.stream", page.FindShapeById("stream")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("data.pipeline", page.FindShapeById("pipeline")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("data.catalog", page.FindShapeById("catalog")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("data.api", page.FindShapeById("api")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("data.quality", page.FindShapeById("quality")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("data.database", page.FindShapeById("database")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("data.database", page.FindShapeById("already")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Empty(VisioValidator.Validate(updatedPath));
        }

        [Fact]
        public void CollaborationBusinessPresetMigratesLegacyBusinessLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument.Create(filePath)
                .AsFluent()
                .Page("Legacy Business Process", page => page
                    .Rect("lane", 4.5, 6.4, 6.0, 1.1, "Responsibility Lane")
                    .Diamond("approval", 2, 5, 1.3, 0.9, "Approval Review")
                    .Data("document", 4, 5, 1.6, 0.8, "Policy Document")
                    .Data("message", 6, 5, 1.6, 0.8, "Email Notification")
                    .Rect("meeting", 8, 5, 1.7, 0.8, "Planning Workshop")
                    .Rect("system", 2, 3.8, 1.9, 0.85, "Business Application")
                    .Rect("team", 4, 3.8, 1.7, 0.8, "Finance Team")
                    .Ellipse("person", 6, 3.8, 0.9, 0.9, "Requester Person")
                    .Stencil("already", VisioStencils.CollaborationBusiness.Get("team"), 8, 3.8, "Already stenciled"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationPresets.CollaborationBusiness(VisioStencils.CollaborationBusiness);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationPlan plan = loaded.PlanStencilMigration(map);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(map);
            loaded.Save(updatedPath);

            Assert.Equal(8, plan.Count);
            Assert.Equal(8, result.Count);
            Assert.DoesNotContain(result.Replacements, replacement => replacement.ShapeId == "already");

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            Assert.Equal("collab.lane", page.FindShapeById("lane")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("collab.approval", page.FindShapeById("approval")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("collab.document", page.FindShapeById("document")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("collab.message", page.FindShapeById("message")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("collab.meeting", page.FindShapeById("meeting")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("collab.system", page.FindShapeById("system")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("collab.team", page.FindShapeById("team")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("collab.person", page.FindShapeById("person")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("collab.team", page.FindShapeById("already")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Empty(VisioValidator.Validate(updatedPath));
        }
    }
}
