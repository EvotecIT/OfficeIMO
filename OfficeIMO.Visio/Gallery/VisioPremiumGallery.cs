using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Generates premium OfficeIMO.Visio reference diagrams used for showcase and visual baseline proof.
    /// </summary>
    public static class VisioPremiumGallery {
        /// <summary>
        /// Generates premium gallery documents into a folder and optionally validates them.
        /// </summary>
        /// <param name="folderPath">Target folder.</param>
        /// <param name="options">Gallery generation options.</param>
        public static IReadOnlyList<VisioGalleryResult> Create(string folderPath, VisioGalleryOptions? options = null) {
            if (string.IsNullOrWhiteSpace(folderPath)) {
                throw new ArgumentException("Folder path cannot be null or whitespace.", nameof(folderPath));
            }

            Directory.CreateDirectory(folderPath);
            VisioGalleryOptions resolvedOptions = options ?? new VisioGalleryOptions();
            List<VisioGalleryResult> results = new() {
                CreateResult("Premium Cloud Architecture", Path.Combine(folderPath, "Premium - Cloud Architecture.vsdx"), CreateCloudArchitecture, resolvedOptions),
                CreateResult("Premium Network Segmentation", Path.Combine(folderPath, "Premium - Network Segmentation.vsdx"), CreateNetworkSegmentation, resolvedOptions),
                CreateResult("Premium Executive Dependencies", Path.Combine(folderPath, "Premium - Executive Dependencies.vsdx"), CreateExecutiveDependencyGraph, resolvedOptions),
                CreateResult("Premium Technical Topology", Path.Combine(folderPath, "Premium - Technical Topology.vsdx"), CreateTechnicalTopology, resolvedOptions),
                CreateResult("Premium Print Audit Trail", Path.Combine(folderPath, "Premium - Print Audit Trail.vsdx"), CreatePrintAuditTrail, resolvedOptions),
                CreateResult("Premium Incident Sequence", Path.Combine(folderPath, "Premium - Incident Sequence.vsdx"), CreateIncidentSequence, resolvedOptions),
                CreateResult("Premium Release Timeline", Path.Combine(folderPath, "Premium - Release Timeline.vsdx"), CreateReleaseTimeline, resolvedOptions),
                CreateResult("Premium Governed Process", Path.Combine(folderPath, "Premium - Governed Process.vsdx"), CreateGovernedSwimlane, resolvedOptions)
            };

            return results;
        }

        private static VisioGalleryResult CreateResult(string name, string filePath, Func<string, VisioDocument> createDocument, VisioGalleryOptions options) {
            VisioDocument document = createDocument(filePath);
            document.Save();

            IReadOnlyList<string> packageIssues = options.ValidatePackage
                ? VisioValidator.Validate(filePath)
                : Array.Empty<string>();
            IReadOnlyList<VisioDiagramQualityIssue> qualityIssues = options.AnalyzeVisualQuality
                ? document.AnalyzeVisualQuality(options.QualityOptions)
                : Array.Empty<VisioDiagramQualityIssue>();
            VisioDesktopValidationResult? desktopValidation = options.ValidateWithVisioDesktop
                ? VisioDesktopValidator.Validate(filePath, options.DesktopValidationOptions)
                : null;

            return new VisioGalleryResult(name, filePath, packageIssues, qualityIssues, desktopValidation, options.RequireVisioDesktop);
        }

        private static VisioDocument CreateCloudArchitecture(string filePath) {
            return VisioDocument.Create(filePath)
                .ArchitectureDiagram("Premium Cloud Architecture", diagram => diagram
                    .Title("Customer Portal - Production Architecture")
                    .Legend()
                    .Theme(VisioStyleTheme.Cloud())
                    .PageSize(18, 8.5)
                    .Margins(0.85, 0.85)
                    .Spacing(0.95, 0.9)
                    .ComponentSize(1.55, 0.86)
                    .Region("edge", "Edge and Identity", 0, 0, 2, 3)
                    .Region("runtime", "Runtime Services", 2, 0, 3, 3)
                    .Region("data", "Data Platform", 5, 0, 2, 3)
                    .Actor("users", "Customers", 0, 1)
                    .Security("identity", "Identity", 1, 0)
                    .Gateway("front-door", "Front Door", 1, 1)
                    .Service("portal", "Portal", 2, 1)
                    .Service("api", "API", 3, 1)
                    .Queue("events", "Events", 3, 2)
                    .Compute("worker", "Worker", 4, 2)
                    .Database("sql", "SQL", 5, 1)
                    .Storage("archive", "Archive", 6, 2)
                    .DataFlow("users", "front-door", "HTTPS")
                    .Dependency("front-door", "identity", "auth")
                    .DataFlow("front-door", "portal", "route")
                    .ControlFlow("portal", "api", "REST")
                    .ControlFlow("api", "events", "publish")
                    .ControlFlow("events", "worker", "trigger")
                    .DataFlow("api", "sql", "SQL")
                    .DataFlow("worker", "archive", "export")
                    .Callout("sql", "confidential", "PII data boundary", VisioSide.Right));
        }

        private static VisioDocument CreateNetworkSegmentation(string filePath) {
            return VisioDocument.Create(filePath)
                .NetworkDiagram("Premium Network Segmentation", network => network
                    .Title("Branch Network - Segmented Access")
                    .Theme(VisioStyleTheme.Enterprise())
                    .PageSize(18, 10.5)
                    .Margins(0.85, 0.85)
                    .Spacing(1.2, 0.95)
                    .NodeSize(1.35, 0.78)
                    .Zone("perimeter", "Perimeter", 0, 0, 3, 2)
                    .Zone("servers", "Server Zone", 3, 0, 4, 3)
                    .Zone("clients", "Client LAN", 0, 3, 3, 2)
                    .Internet("internet", "Internet", 0, 1)
                    .Firewall("edge", "Edge Firewall", 1, 1)
                    .Switch("core", "Core Switch", 2, 1)
                    .Server("jump", "Jump Host", 4, 0)
                    .Server("app", "App Server", 4, 1)
                    .Database("db", "Database", 5, 1)
                    .Storage("backup", "Backup NAS", 4, 2)
                    .Switch("access", "Access Switch", 1, 3)
                    .Workstation("finance", "Finance PC", 0, 4)
                    .Workstation("support", "Support PC", 1, 4)
                    .Wireless("wifi", "Guest Wi-Fi", 2, 4)
                    .Ethernet("internet", "edge", "WAN")
                    .Trunk("edge", "core", "uplink")
                    .Management("core", "jump", "admin")
                    .Trunk("core", "app", "10Gb")
                    .Ethernet("app", "db", "SQL")
                    .Management("app", "backup", "backup")
                    .Trunk("core", "access", "client trunk")
                    .Ethernet("access", "finance", "wired")
                    .Ethernet("access", "support", "wired")
                    .WirelessLink("access", "wifi", "guest")
                    .Callout("core", "VLAN 20/30 and guest access", VisioSide.Top));
        }

        private static VisioDocument CreateExecutiveDependencyGraph(string filePath) {
            VisioDocument document = VisioDocument.Create(filePath);
            document.UseMastersByDefault = true;

            return document.GraphDiagram("Premium Executive Dependencies", graph => graph
                    .Title("Revenue Platform - Critical Path")
                    .Theme(VisioStyleTheme.Enterprise())
                    .Layout(VisioGraphLayout.Layered)
                    .Direction(VisioGraphDirection.LeftToRight)
                    .PageSize(16.5, 6.8)
                    .Margins(0.8, 0.85, 0.8, 0.9)
                    .NodeSize(1.35, 0.78)
                    .Spacing(0.72, 0.9)
                    .StencilNode("sales", "Sales", VisioStencils.Architecture, "actor", "client")
                    .StencilNode("portal", "Portal", VisioStencils.Architecture, "service", "app")
                    .StencilNode("checkout", "Checkout", VisioStencils.Architecture, "service", "api")
                    .StencilNode("risk", "Risk", VisioStencils.Architecture, "security", "policy")
                    .StencilNode("payments", "Payments", VisioStencils.Architecture, "service", "api")
                    .StencilNode("ledger", "Ledger", VisioStencils.Architecture, "database", "sql")
                    .StencilNode("warehouse", "Warehouse", VisioStencils.Architecture, "storage", "object-store")
                    .Root("sales")
                    .ControlEdge("sales-portal", "sales", "portal", "qualified lead")
                    .ControlEdge("portal-checkout", "portal", "checkout", "order")
                    .ControlEdge("checkout-risk", "checkout", "risk", "score")
                    .ControlEdge("risk-payments", "risk", "payments", "approved")
                    .DataEdge("payments-ledger", "payments", "ledger", "settle")
                    .DataEdge("checkout-warehouse", "checkout", "warehouse", "reserve")
                    .NodeShapeData("portal", "Owner", "Digital", "Owner", VisioShapeDataType.String)
                    .NodeShapeData("payments", "Criticality", "Tier 0", "Criticality", VisioShapeDataType.String)
                    .EdgeShapeData("payments-ledger", "Sla", "15 minutes", "SLA", VisioShapeDataType.String));
        }

        private static VisioDocument CreateTechnicalTopology(string filePath) {
            return VisioDocument.Create(filePath)
                .BlockDiagram("Premium Technical Topology", diagram => diagram
                    .Title("Zero Trust Runtime - Technical Topology")
                    .Theme(VisioStyleTheme.Technical())
                    .PageSize(17.4, 9.2)
                    .Margins(0.75, 0.75)
                    .Region("edge-zone", "Edge", 0, 0, 2, 3)
                    .Region("runtime-zone", "Runtime Mesh", 2, 0, 2, 3)
                    .Region("data-zone", "State", 4, 0, 1, 3)
                    .Block("endpoint", "Endpoint", 0, 1)
                    .Block("ingress", "Ingress", 1, 1, VisioBlockShapeKind.Decision)
                    .EmphasisBlock("api", "API Pod", 2, 1)
                    .Block("policy", "Policy", 3, 0, VisioBlockShapeKind.Decision)
                    .Block("worker", "Worker Pool", 3, 1)
                    .Block("events", "Event Bus", 2, 2, VisioBlockShapeKind.Data)
                    .Block("cache", "Cache", 4, 1, VisioBlockShapeKind.Data)
                    .Block("vault", "Secrets", 4, 0, VisioBlockShapeKind.Decision)
                    .ControlFlow("endpoint", "ingress", "mTLS")
                    .ControlFlow("ingress", "api", "route")
                    .ControlFlow("api", "policy", "authorize")
                    .ControlFlow("policy", "worker", "allow")
                    .ControlFlow("policy", "vault", "secret")
                    .DataFlow("api", "events", "publish")
                    .DataFlow("worker", "cache", "read/write"));
        }

        private static VisioDocument CreatePrintAuditTrail(string filePath) {
            return VisioDocument.Create(filePath)
                .Flowchart("Premium Print Audit Trail", flow => flow
                    .Title("Quarter-End Access Review - Print Audit Trail")
                    .Theme(VisioStyleTheme.Print())
                    .PageSize(8.5, 13.5)
                    .Spacing(0.36)
                    .Start("start", "Review window opens")
                    .Data("snapshot", "Export access snapshot")
                    .Step("owner-review", "Owner review")
                    .Decision("evidence-complete", "Evidence complete?")
                    .Step("remediate", "Remediate gaps")
                    .Step("evidence", "Attach evidence")
                    .End("archive", "Archive signed record"));
        }

        private static VisioDocument CreateIncidentSequence(string filePath) {
            return VisioDocument.Create(filePath)
                .SequenceDiagram("Premium Incident Sequence", sequence => sequence
                    .Title("Payment Timeout - Detection and Recovery")
                    .Theme(VisioStyleTheme.DarkSafe())
                    .PageSize(12, 6.8)
                    .Margins(0.8, 0.75, 0.8, 0.75)
                    .ParticipantSize(1.25, 0.62)
                    .Spacing(1.35, 0.74, 0.68)
                    .Actor("support", "Support")
                    .Participant("monitor", "Monitor")
                    .Control("api", "Payments API")
                    .Participant("queue", "Retry Queue")
                    .Database("ledger", "Ledger")
                    .Call("monitor", "support", "Alert: timeout spike", "alert")
                    .Call("support", "api", "Check health", "check")
                    .Return("api", "support", "Gateway latency")
                    .Async("support", "queue", "Pause retries", "pause")
                    .Call("api", "ledger", "Verify settlement", "verify")
                    .Return("ledger", "api", "Consistent")
                    .Async("queue", "api", "Resume controlled drain", "resume")
                    .SelfMessage("support", "Update incident record", id: "record")
                    .Activation("support", 0, 7, "support-active")
                    .Activation("api", 1, 6, "api-active")
                    .Activation("queue", 3, 6, "queue-active"));
        }

        private static VisioDocument CreateReleaseTimeline(string filePath) {
            return VisioDocument.Create(filePath)
                .TimelineDiagram("Premium Release Timeline", timeline => timeline
                    .Title("Visio Premium Gallery Roadmap")
                    .Theme(VisioStyleTheme.Process())
                    .PageSize(13.5, 5.8)
                    .Margins(0.65, 0.55, 0.65, 0.55)
                    .AxisY(2.35)
                    .MilestoneSize(0.22, 1.45, 0.48)
                    .SpanSize(0.3, 0.14)
                    .Range(new DateTime(2026, 6, 1), new DateTime(2026, 9, 30))
                    .Span("gallery", new DateTime(2026, 6, 3), new DateTime(2026, 7, 12), "Premium gallery", 0)
                    .Span("stencils", new DateTime(2026, 6, 18), new DateTime(2026, 8, 9), "Stencil packs", 1)
                    .Span("baselines", new DateTime(2026, 7, 15), new DateTime(2026, 9, 10), "Visual baselines", 0, VisioTimelinePlacement.Below)
                    .Milestone("brief", new DateTime(2026, 6, 5), "Design brief", VisioTimelinePlacement.Above)
                    .Decision("gate", new DateTime(2026, 7, 19), "Visual gate", VisioTimelinePlacement.Below)
                    .Risk("risk", new DateTime(2026, 8, 14), "Export drift", VisioTimelinePlacement.Above)
                    .Release("preview", new DateTime(2026, 9, 4), "Preview", VisioTimelinePlacement.Below)
                    .Milestone("ship", new DateTime(2026, 9, 24), "Ship", VisioTimelinePlacement.Above));
        }

        private static VisioDocument CreateGovernedSwimlane(string filePath) {
            return VisioDocument.Create(filePath)
                .SwimlaneDiagram("Premium Governed Process", swim => swim
                    .Title("Access Request - Governed Fulfillment")
                    .Theme(VisioStyleTheme.Process())
                    .PageSize(14, 8.5)
                    .Margins(0.7, 0.7, 0.7, 0.7)
                    .GridSize(2.65, 1.55, 1.45, 0.55)
                    .ActivitySize(1.72, 0.74)
                    .Lane("requester", "Requester")
                    .Lane("manager", "Manager")
                    .Lane("identity", "Identity")
                    .Lane("security", "Security")
                    .Phase("submit", "Submit")
                    .Phase("approve", "Approve")
                    .Phase("provision", "Provision")
                    .Phase("review", "Review")
                    .Start("request", "Submit request", "requester", "submit")
                    .Step("manager-review", "Business approval", "manager", "approve")
                    .Decision("sensitive", "Sensitive?", "security", "approve")
                    .Step("provision", "Provision access", "identity", "provision")
                    .Data("evidence", "Record evidence", "identity", "review")
                    .End("notify", "Notify requester", "requester", "review")
                    .Handoff("request", "manager-review", "handoff")
                    .Flow("manager-review", "sensitive")
                    .Exception("sensitive", "manager-review")
                    .Handoff("sensitive", "provision")
                    .Flow("provision", "evidence")
                    .Flow("evidence", "notify")
                    .Callout("evidence", "Retention: 1 year", VisioSide.Bottom));
        }
    }
}
