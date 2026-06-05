using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio {
    public static partial class VisioGallery {
        /// <summary>
        /// Creates a data-driven platform lineage graph that demonstrates lake, warehouse, quality, catalog, API, and analytics flows.
        /// </summary>
        /// <param name="filePath">Target VSDX file path.</param>
        public static VisioDocument CreateDataPlatformLineageGraph(string filePath) {
            VisioGraphNodeRecord source = CreateNode("source-system", "Source System", VisioStencils.CollaborationBusiness, "system", "application");
            source.IsRoot = true;
            source.ShapeData.Add("Domain", "Orders");

            VisioGraphNodeRecord stream = CreateNode("event-stream", "Event Stream", VisioStencils.DataPlatform, "stream", "event-stream", "kafka");
            stream.ShapeData.Add("Topic", "orders.changed");

            VisioGraphNodeRecord batch = CreateNode("batch-pipeline", "Batch Pipeline", VisioStencils.DataPlatform, "pipeline", "etl", "job");
            batch.ShapeData.Add("Schedule", "Hourly");
            batch.HyperlinkAddress = "https://example.org/data/pipelines/orders-batch";
            batch.HyperlinkDescription = "Orders batch pipeline";

            VisioGraphNodeRecord quality = CreateNode("quality-gate", "Quality Gate", VisioStencils.DataPlatform, "quality", "validation", "dq");
            quality.ShapeData.Add("RuleSet", "Completeness, freshness, keys");

            VisioGraphNodeRecord lake = CreateNode("data-lake", "Data Lake", VisioStencils.DataPlatform, "lake", "analytics", "storage");
            lake.ShapeData.Add("Retention", "7 years");

            VisioGraphNodeRecord warehouse = CreateNode("warehouse", "Warehouse", VisioStencils.DataPlatform, "warehouse", "dwh", "mart");
            warehouse.ShapeData.Add("Sla", "99.9%");

            VisioGraphNodeRecord catalog = CreateNode("catalog", "Data Catalog", VisioStencils.DataPlatform, "catalog", "metadata", "lineage");
            catalog.ShapeData.Add("Owner", "Data Governance");
            catalog.HyperlinkAddress = "https://example.org/data/catalog/orders";
            catalog.HyperlinkDescription = "Orders lineage catalog";

            VisioGraphNodeRecord api = CreateNode("query-api", "Query API", VisioStencils.DataPlatform, "api", "query", "endpoint");
            api.ShapeData.Add("Audience", "Internal services");

            VisioGraphNodeRecord analytics = CreateNode("analytics", "Analytics Workspace", VisioStencils.Cloud, "monitoring", "metrics", "analytics");
            analytics.ShapeData.Add("Consumers", "Finance, Operations");

            VisioGraphEdgeRecord capture = new("source-system", "event-stream") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "CDC"
            };
            capture.ShapeData.Add("Format", "CloudEvents");

            VisioGraphEdgeRecord schedule = new("source-system", "batch-pipeline") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "extract"
            };
            schedule.ShapeData.Add("Window", "hourly");

            VisioGraphEdgeRecord validate = new("batch-pipeline", "quality-gate") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "validate"
            };
            validate.ShapeData.Add("Gate", "block on critical");

            VisioGraphEdgeRecord streamToLake = new("event-stream", "data-lake") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "bronze"
            };
            streamToLake.ShapeData.Add("Format", "JSON");

            VisioGraphEdgeRecord batchToLake = new("quality-gate", "data-lake") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "curated"
            };
            batchToLake.ShapeData.Add("Format", "Parquet");

            VisioGraphEdgeRecord model = new("data-lake", "warehouse") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "model"
            };
            model.ShapeData.Add("Layer", "silver/gold");

            VisioGraphEdgeRecord publish = new("warehouse", "query-api") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "publish"
            };
            publish.ShapeData.Add("Contract", "versioned schema");

            VisioGraphEdgeRecord lineageLake = new("data-lake", "catalog") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "lineage"
            };

            VisioGraphEdgeRecord lineageWarehouse = new("warehouse", "catalog") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "metadata"
            };

            VisioGraphEdgeRecord consume = new("query-api", "analytics") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "semantic model"
            };
            consume.ShapeData.Add("Refresh", "15 minutes");

            VisioGraphClusterRecord ingestion = new("ingestion-cluster", "Ingestion", new[] { "source-system", "event-stream", "batch-pipeline" });
            ingestion.ShapeData.Add("Owner", "Data Engineering");
            ingestion.HyperlinkAddress = "https://example.org/runbooks/data-ingestion";
            ingestion.HyperlinkDescription = "Data ingestion runbook";

            VisioGraphClusterRecord serving = new("serving-cluster", "Quality, Serving, and Governance", new[] { "quality-gate", "data-lake", "warehouse", "catalog", "query-api", "analytics" });
            serving.ShapeData.Add("Owner", "Data Platform");

            return VisioDocument.Create(filePath)
                .GraphDiagram("Data Platform Lineage Graph", graph => graph
                    .Title("Orders Data Platform Lineage")
                    .Theme(VisioStyleTheme.Technical())
                    .Layout(VisioGraphLayout.Layered)
                    .Direction(VisioGraphDirection.LeftToRight)
                    .Legend()
                    .PageSize(21.2, 9.2)
                    .Margins(0.8, 0.9, 0.8, 0.8)
                    .NodeSize(1.38, 0.76)
                    .Spacing(1.18, 1.18)
                    .Import(
                        new[] { source, stream, batch, quality, lake, warehouse, catalog, api, analytics },
                        new[] { capture, schedule, validate, streamToLake, batchToLake, model, publish, lineageLake, lineageWarehouse, consume },
                        new[] { ingestion, serving }));
        }
    }
}
