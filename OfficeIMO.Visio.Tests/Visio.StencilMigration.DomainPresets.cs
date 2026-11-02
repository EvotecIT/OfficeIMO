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
                .AddPage("Legacy Kubernetes", page => page
                    .AddRectangle("cluster", OfficeLength.Inches(5), OfficeLength.Inches(6.5), OfficeLength.Inches(5.0), OfficeLength.Inches(1.1), "AKS Cluster")
                    .AddRectangle("namespace", OfficeLength.Inches(2), OfficeLength.Inches(5.2), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Tenant Namespace")
                    .AddRectangle("node", OfficeLength.Inches(4), OfficeLength.Inches(5.2), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Worker Node")
                    .AddRectangle("cluster-node", OfficeLength.Inches(6), OfficeLength.Inches(6.5), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Cluster Node")
                    .AddHexagon("pod", OfficeLength.Inches(6), OfficeLength.Inches(5.2), OfficeLength.Inches(1.4), OfficeLength.Inches(0.8), "Workload Pod")
                    .AddRectangle("container", OfficeLength.Inches(8), OfficeLength.Inches(5.2), OfficeLength.Inches(1.6), OfficeLength.Inches(0.8), "Container Image")
                    .AddRectangle("service", OfficeLength.Inches(2), OfficeLength.Inches(4.0), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Kubernetes Service")
                    .AddDiamond("ingress", OfficeLength.Inches(4), OfficeLength.Inches(4.0), OfficeLength.Inches(1.3), OfficeLength.Inches(0.9), "Ingress Route")
                    .AddData("config", OfficeLength.Inches(6), OfficeLength.Inches(4.0), OfficeLength.Inches(1.6), OfficeLength.Inches(0.7), "Config Map")
                    .AddData("secret", OfficeLength.Inches(8), OfficeLength.Inches(4.0), OfficeLength.Inches(1.6), OfficeLength.Inches(0.7), "Secret Credential")
                    .AddStencilShape(id: "already", stencil: VisioStencils.ContainersKubernetes.Get("pod"), x: OfficeLength.Inches(5), y: OfficeLength.Inches(2.8), text: "Already stenciled"))
                .End()
                .Save();

            VisioStencilMigrationMap map = VisioStencilMigrationPresets.ContainersKubernetes(VisioStencils.ContainersKubernetes);

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioStencilMigrationPlan plan = loaded.PlanStencilMigration(map);
            VisioStencilMigrationResult result = loaded.ApplyStencilMigration(map);
            loaded.Save(updatedPath);

            Assert.Equal(10, plan.Count);
            Assert.Equal(10, result.Count);
            Assert.DoesNotContain(result.Replacements, replacement => replacement.ShapeId == "already");

            VisioDocument updated = VisioDocument.Load(updatedPath);
            VisioPage page = Assert.Single(updated.Pages);
            Assert.Equal("k8s.cluster", page.FindShapeById("cluster")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("k8s.namespace", page.FindShapeById("namespace")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("k8s.node", page.FindShapeById("node")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
            Assert.Equal("k8s.node", page.FindShapeById("cluster-node")!.GetUserCellValue(VisioSemanticUserCells.StencilId));
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
                .AddPage("Legacy Data Platform", page => page
                    .AddData("lake", OfficeLength.Inches(2), OfficeLength.Inches(6), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Lake Storage")
                    .AddData("warehouse", OfficeLength.Inches(4), OfficeLength.Inches(6), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Data Warehouse")
                    .AddData("stream", OfficeLength.Inches(6), OfficeLength.Inches(6), OfficeLength.Inches(1.7), OfficeLength.Inches(0.75), "Event Stream")
                    .AddRectangle("pipeline", OfficeLength.Inches(8), OfficeLength.Inches(6), OfficeLength.Inches(1.7), OfficeLength.Inches(0.75), "ETL Pipeline")
                    .AddRectangle("catalog", OfficeLength.Inches(2), OfficeLength.Inches(4.6), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Metadata Catalog")
                    .AddRectangle("api", OfficeLength.Inches(4), OfficeLength.Inches(4.6), OfficeLength.Inches(1.6), OfficeLength.Inches(0.8), "Query API")
                    .AddDiamond("quality", OfficeLength.Inches(6), OfficeLength.Inches(4.6), OfficeLength.Inches(1.3), OfficeLength.Inches(0.9), "Quality Gate")
                    .AddData("database", OfficeLength.Inches(8), OfficeLength.Inches(4.6), OfficeLength.Inches(1.7), OfficeLength.Inches(0.8), "SQL Database")
                    .AddStencilShape(id: "already", stencil: VisioStencils.DataPlatform.Get("database"), x: OfficeLength.Inches(5), y: OfficeLength.Inches(3.2), text: "Already stenciled"))
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
                .AddPage("Legacy Business Process", page => page
                    .AddRectangle("lane", OfficeLength.Inches(4.5), OfficeLength.Inches(6.4), OfficeLength.Inches(6.0), OfficeLength.Inches(1.1), "Responsibility Lane")
                    .AddDiamond("approval", OfficeLength.Inches(2), OfficeLength.Inches(5), OfficeLength.Inches(1.3), OfficeLength.Inches(0.9), "Approval Review")
                    .AddData("document", OfficeLength.Inches(4), OfficeLength.Inches(5), OfficeLength.Inches(1.6), OfficeLength.Inches(0.8), "Policy Document")
                    .AddData("message", OfficeLength.Inches(6), OfficeLength.Inches(5), OfficeLength.Inches(1.6), OfficeLength.Inches(0.8), "Email Notification")
                    .AddRectangle("meeting", OfficeLength.Inches(8), OfficeLength.Inches(5), OfficeLength.Inches(1.7), OfficeLength.Inches(0.8), "Planning Workshop")
                    .AddRectangle("system", OfficeLength.Inches(2), OfficeLength.Inches(3.8), OfficeLength.Inches(1.9), OfficeLength.Inches(0.85), "Business Application")
                    .AddRectangle("team", OfficeLength.Inches(4), OfficeLength.Inches(3.8), OfficeLength.Inches(1.7), OfficeLength.Inches(0.8), "Finance Team")
                    .AddEllipse("person", OfficeLength.Inches(6), OfficeLength.Inches(3.8), OfficeLength.Inches(0.9), OfficeLength.Inches(0.9), "Requester Person")
                    .AddStencilShape(id: "already", stencil: VisioStencils.CollaborationBusiness.Get("team"), x: OfficeLength.Inches(8), y: OfficeLength.Inches(3.8), text: "Already stenciled"))
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
