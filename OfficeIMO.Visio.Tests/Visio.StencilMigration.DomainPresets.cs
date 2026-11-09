using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioStencilMigrationDomainPresetTests {
        [Fact]
        public void ContainersKubernetesPresetMigratesLegacyPlatformLabels() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            string updatedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            using var nativeDocument3 = VisioDocument.Create(filePath);
            nativeDocument3.AddPage("Legacy Kubernetes", page => {
                page.AddRectangle("cluster", OfficeLength.Inches(5), OfficeLength.Inches(6.5), OfficeLength.Inches(5.0), OfficeLength.Inches(1.1), "AKS Cluster");
                page.AddRectangle("namespace", OfficeLength.Inches(2), OfficeLength.Inches(5.2), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Tenant Namespace");
                page.AddRectangle("node", OfficeLength.Inches(4), OfficeLength.Inches(5.2), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Worker Node");
                page.AddRectangle("cluster-node", OfficeLength.Inches(6), OfficeLength.Inches(6.5), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Cluster Node");
                page.AddHexagon("pod", OfficeLength.Inches(6), OfficeLength.Inches(5.2), OfficeLength.Inches(1.4), OfficeLength.Inches(0.8), "Workload Pod");
                page.AddRectangle("container", OfficeLength.Inches(8), OfficeLength.Inches(5.2), OfficeLength.Inches(1.6), OfficeLength.Inches(0.8), "Container Image");
                page.AddRectangle("service", OfficeLength.Inches(2), OfficeLength.Inches(4.0), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Kubernetes Service");
                page.AddDiamond("ingress", OfficeLength.Inches(4), OfficeLength.Inches(4.0), OfficeLength.Inches(1.3), OfficeLength.Inches(0.9), "Ingress Route");
                page.AddData("config", OfficeLength.Inches(6), OfficeLength.Inches(4.0), OfficeLength.Inches(1.6), OfficeLength.Inches(0.7), "Config Map");
                page.AddData("secret", OfficeLength.Inches(8), OfficeLength.Inches(4.0), OfficeLength.Inches(1.6), OfficeLength.Inches(0.7), "Secret Credential");
                page.AddStencilShape(id: "already", stencil: VisioStencils.ContainersKubernetes.Get("pod"), x: OfficeLength.Inches(5), y: OfficeLength.Inches(2.8), text: "Already stenciled");
            });
            nativeDocument3.Save();

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

            using var nativeDocument2 = VisioDocument.Create(filePath);
            nativeDocument2.AddPage("Legacy Data Platform", page => {
                page.AddData("lake", OfficeLength.Inches(2), OfficeLength.Inches(6), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Lake Storage");
                page.AddData("warehouse", OfficeLength.Inches(4), OfficeLength.Inches(6), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Data Warehouse");
                page.AddData("stream", OfficeLength.Inches(6), OfficeLength.Inches(6), OfficeLength.Inches(1.7), OfficeLength.Inches(0.75), "Event Stream");
                page.AddRectangle("pipeline", OfficeLength.Inches(8), OfficeLength.Inches(6), OfficeLength.Inches(1.7), OfficeLength.Inches(0.75), "ETL Pipeline");
                page.AddRectangle("catalog", OfficeLength.Inches(2), OfficeLength.Inches(4.6), OfficeLength.Inches(1.8), OfficeLength.Inches(0.8), "Metadata Catalog");
                page.AddRectangle("api", OfficeLength.Inches(4), OfficeLength.Inches(4.6), OfficeLength.Inches(1.6), OfficeLength.Inches(0.8), "Query API");
                page.AddDiamond("quality", OfficeLength.Inches(6), OfficeLength.Inches(4.6), OfficeLength.Inches(1.3), OfficeLength.Inches(0.9), "Quality Gate");
                page.AddData("database", OfficeLength.Inches(8), OfficeLength.Inches(4.6), OfficeLength.Inches(1.7), OfficeLength.Inches(0.8), "SQL Database");
                page.AddStencilShape(id: "already", stencil: VisioStencils.DataPlatform.Get("database"), x: OfficeLength.Inches(5), y: OfficeLength.Inches(3.2), text: "Already stenciled");
            });
            nativeDocument2.Save();

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

            using var nativeDocument1 = VisioDocument.Create(filePath);
            nativeDocument1.AddPage("Legacy Business Process", page => {
                page.AddRectangle("lane", OfficeLength.Inches(4.5), OfficeLength.Inches(6.4), OfficeLength.Inches(6.0), OfficeLength.Inches(1.1), "Responsibility Lane");
                page.AddDiamond("approval", OfficeLength.Inches(2), OfficeLength.Inches(5), OfficeLength.Inches(1.3), OfficeLength.Inches(0.9), "Approval Review");
                page.AddData("document", OfficeLength.Inches(4), OfficeLength.Inches(5), OfficeLength.Inches(1.6), OfficeLength.Inches(0.8), "Policy Document");
                page.AddData("message", OfficeLength.Inches(6), OfficeLength.Inches(5), OfficeLength.Inches(1.6), OfficeLength.Inches(0.8), "Email Notification");
                page.AddRectangle("meeting", OfficeLength.Inches(8), OfficeLength.Inches(5), OfficeLength.Inches(1.7), OfficeLength.Inches(0.8), "Planning Workshop");
                page.AddRectangle("system", OfficeLength.Inches(2), OfficeLength.Inches(3.8), OfficeLength.Inches(1.9), OfficeLength.Inches(0.85), "Business Application");
                page.AddRectangle("team", OfficeLength.Inches(4), OfficeLength.Inches(3.8), OfficeLength.Inches(1.7), OfficeLength.Inches(0.8), "Finance Team");
                page.AddEllipse("person", OfficeLength.Inches(6), OfficeLength.Inches(3.8), OfficeLength.Inches(0.9), OfficeLength.Inches(0.9), "Requester Person");
                page.AddStencilShape(id: "already", stencil: VisioStencils.CollaborationBusiness.Get("team"), x: OfficeLength.Inches(8), y: OfficeLength.Inches(3.8), text: "Already stenciled");
            });
            nativeDocument1.Save();

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
