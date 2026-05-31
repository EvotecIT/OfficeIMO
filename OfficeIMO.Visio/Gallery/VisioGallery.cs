using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Generates a small gallery of polished OfficeIMO.Visio reference diagrams.
    /// </summary>
    public static class VisioGallery {
        /// <summary>
        /// Generates gallery documents into a folder and optionally validates them.
        /// </summary>
        /// <param name="folderPath">Target folder.</param>
        /// <param name="options">Gallery generation options.</param>
        public static IReadOnlyList<VisioGalleryResult> Create(string folderPath, VisioGalleryOptions? options = null) {
            if (string.IsNullOrWhiteSpace(folderPath)) {
                throw new ArgumentException("Folder path cannot be null or whitespace.", nameof(folderPath));
            }

            Directory.CreateDirectory(folderPath);
            VisioGalleryOptions resolvedOptions = options ?? new VisioGalleryOptions();
            List<VisioGalleryResult> results = new();
            results.Add(CreateResult("Approval Flowchart", Path.Combine(folderPath, "OfficeIMO Gallery - Approval Flowchart.vsdx"), CreateApprovalFlowchart, resolvedOptions));
            results.Add(CreateResult("Service Block Diagram", Path.Combine(folderPath, "OfficeIMO Gallery - Service Block Diagram.vsdx"), CreateServiceBlockDiagram, resolvedOptions));
            results.Add(CreateResult("Architecture Diagram", Path.Combine(folderPath, "OfficeIMO Gallery - Architecture Diagram.vsdx"), CreateArchitectureDiagram, resolvedOptions));
            results.Add(CreateResult("Network Diagram", Path.Combine(folderPath, "OfficeIMO Gallery - Network Diagram.vsdx"), CreateNetworkDiagram, resolvedOptions));
            results.Add(CreateResult("Swimlane Process", Path.Combine(folderPath, "OfficeIMO Gallery - Swimlane Process.vsdx"), CreateSwimlaneProcess, resolvedOptions));
            results.Add(CreateResult("Org Chart", Path.Combine(folderPath, "OfficeIMO Gallery - Org Chart.vsdx"), CreateOrgChart, resolvedOptions));
            results.Add(CreateResult("Timeline Roadmap", Path.Combine(folderPath, "OfficeIMO Gallery - Timeline Roadmap.vsdx"), CreateTimelineRoadmap, resolvedOptions));
            results.Add(CreateResult("Routed Decision Flow", Path.Combine(folderPath, "OfficeIMO Gallery - Routed Decision Flow.vsdx"), CreateRoutedDecisionFlow, resolvedOptions));
            results.Add(CreateResult("CI/CD Inventory Graph", Path.Combine(folderPath, "OfficeIMO Gallery - CI-CD Inventory Graph.vsdx"), CreateCiCdInventoryGraph, resolvedOptions));
            results.Add(CreateResult("Identity Authentication Graph", Path.Combine(folderPath, "OfficeIMO Gallery - Identity Authentication Graph.vsdx"), CreateIdentityAuthenticationGraph, resolvedOptions));
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

        private static VisioDocument CreateApprovalFlowchart(string filePath) {
            return VisioDocument.Create(filePath)
                .Flowchart("Approval Flow", flow => flow
                    .Theme(VisioStyleTheme.Modern())
                    .PageSize(8.5, 11)
                    .Start("start", "Request received")
                    .Step("review", "Review request")
                    .Decision("approved", "Approved?")
                    .Step("publish", "Publish decision")
                    .End("done", "Done"));
        }

        private static VisioDocument CreateServiceBlockDiagram(string filePath) {
            return VisioDocument.Create(filePath)
                .BlockDiagram("Service Flow", diagram => diagram
                    .Theme(VisioStyleTheme.Technical())
                    .PageSize(18, 9)
                    .Region("edge", "Edge", 0, 0, 2, 1)
                    .Region("core", "Core Services", 2, 0, 3, 2)
                    .Block("client", "Client", 0, 0)
                    .Block("gateway", "Gateway", 1, 0)
                    .EmphasisBlock("api", "API", 2, 0)
                    .Block("queue", "Queue", 3, 1, VisioBlockShapeKind.Data)
                    .Block("store", "Storage", 4, 0, VisioBlockShapeKind.Data)
                    .DataFlow("client", "gateway", "request")
                    .DataFlow("gateway", "api", "route")
                    .ControlFlow("api", "queue", "enqueue")
                    .DataFlow("api", "store", "persist"));
        }

        private static VisioDocument CreateArchitectureDiagram(string filePath) {
            return VisioDocument.Create(filePath)
                .ArchitectureDiagram("Jenkins on Azure", diagram => diagram
                    .Theme(VisioStyleTheme.Technical())
                    .PageSize(14, 8.5)
                    .Region("vnet", "Virtual Network", 1, 0, 4, 3)
                    .Region("subnet", "Build Subnet", 1, 1, 4, 2)
                    .Actor("users", "Users", 0, 1)
                    .Gateway("public-ip", "Public IP", 1, 1)
                    .Service("jenkins", "Jenkins Server", 2, 1)
                    .Compute("agent", "Build Agent", 3, 1)
                    .Database("data", "Data", 2, 2)
                    .Storage("artifacts", "Artifacts", 4, 2)
                    .Security("vault", "Key Vault", 2, 0)
                    .DataFlow("users", "public-ip", "HTTPS")
                    .DataFlow("public-ip", "jenkins", "route")
                    .ControlFlow("jenkins", "agent", "scale")
                    .Dependency("jenkins", "data", "state")
                    .Dependency("jenkins", "vault", "secrets")
                    .DataFlow("agent", "artifacts", "publish"));
        }

        private static VisioDocument CreateNetworkDiagram(string filePath) {
            return VisioDocument.Create(filePath)
                .NetworkDiagram("Branch Network", network => network
                    .Theme(VisioStyleTheme.Technical())
                    .Zone("perimeter", "Perimeter", 0, 0, 3, 1)
                    .Zone("servers", "Server Zone", 3, 0, 3, 1)
                    .Zone("clients", "Client LAN", 1, 2, 5, 1)
                    .Internet("internet", "Internet", 0, 0)
                    .Firewall("firewall", "Firewall", 1, 0)
                    .Switch("core", "Core Switch", 2, 0)
                    .Server("app", "App Server", 3, 0)
                    .Database("db", "Database", 4, 0)
                    .Storage("backup", "Backup NAS", 5, 0)
                    .Workstation("pc1", "Finance PC", 1, 2)
                    .Workstation("pc2", "Support PC", 2, 2)
                    .Printer("printer", "Printer", 3, 2)
                    .Wireless("wifi", "Wi-Fi", 4, 2)
                    .Legend("legend", "solid: data\ndashed: mgmt", 5, 2)
                    .Ethernet("internet", "firewall", "WAN")
                    .Trunk("firewall", "core", "uplink")
                    .Trunk("core", "app", "10Gb")
                    .Ethernet("app", "db")
                    .Ethernet("db", "backup")
                    .Ethernet("core", "pc2")
                    .Ethernet("pc1", "pc2")
                    .Ethernet("pc2", "printer")
                    .WirelessLink("printer", "wifi", "wireless"));
        }

        private static VisioDocument CreateSwimlaneProcess(string filePath) {
            return VisioDocument.Create(filePath)
                .SwimlaneDiagram("Order Fulfillment", swim => swim
                    .Theme(VisioStyleTheme.Modern())
                    .PageSize(14, 8.5)
                    .Lane("customer", "Customer")
                    .Lane("sales", "Sales")
                    .Lane("ops", "Operations")
                    .Phase("request", "Request")
                    .Phase("review", "Review")
                    .Phase("approval", "Approval")
                    .Phase("fulfill", "Fulfill")
                    .Start("start", "Submit order", "customer", "request")
                    .Step("qualify", "Qualify order", "sales", "review")
                    .Decision("approved", "Approved?", "sales", "approval")
                    .Step("revise", "Revise request", "customer", "approval")
                    .Step("pick", "Pick items", "ops", "approval")
                    .Data("invoice", "Create invoice", "sales", "fulfill")
                    .End("ship", "Ship order", "ops", "fulfill")
                    .Flow("start", "qualify", "handoff")
                    .Flow("qualify", "approved")
                    .Exception("approved", "revise", "no")
                    .Handoff("approved", "pick", "yes")
                    .Flow("pick", "invoice")
                    .Flow("invoice", "ship"));
        }

        private static VisioDocument CreateOrgChart(string filePath) {
            return VisioDocument.Create(filePath)
                .OrgChartDiagram("Leadership", org => org
                    .Theme(VisioStyleTheme.Modern())
                    .PageSize(14, 8.5)
                    .Root("ceo", "Marta Nowak", "Chief Executive Officer")
                    .Assistant("ea", "Eli Green", "Executive Assistant", "ceo")
                    .Manager("cto", "Alex Chen", "Chief Technology Officer", "ceo")
                    .Manager("coo", "Sam Rivera", "Chief Operating Officer", "ceo")
                    .Manager("cfo", "Priya Shah", "Chief Financial Officer", "ceo")
                    .TeamBand("engineering", "Engineering", "cto")
                    .TeamBand("operations", "Operations", "coo")
                    .Position("platform", "Nina Patel", "Platform Lead", "cto", "engineering")
                    .Position("security", "Owen Brooks", "Security Lead", "cto", "engineering")
                    .Vacancy("sre", "Open SRE Role", "coo", "operations")
                    .External("advisor", "Taylor Reed", "Advisor", "cfo"));
        }

        private static VisioDocument CreateTimelineRoadmap(string filePath) {
            return VisioDocument.Create(filePath)
                .TimelineDiagram("Product Roadmap", timeline => timeline
                    .Theme(VisioStyleTheme.Modern())
                    .Range(new DateTime(2026, 1, 1), new DateTime(2026, 6, 30))
                    .Span("discovery", new DateTime(2026, 1, 8), new DateTime(2026, 2, 20), "Discovery", 0)
                    .Span("build", new DateTime(2026, 2, 21), new DateTime(2026, 5, 15), "Build", 1)
                    .Span("enablement", new DateTime(2026, 4, 1), new DateTime(2026, 6, 10), "Enablement", 0, VisioTimelinePlacement.Below)
                    .Milestone("kickoff", new DateTime(2026, 1, 12), "Kickoff", VisioTimelinePlacement.Above)
                    .Decision("gate", new DateTime(2026, 2, 25), "Go / no-go", VisioTimelinePlacement.Below)
                    .Risk("risk", new DateTime(2026, 3, 18), "Security review", VisioTimelinePlacement.Above)
                    .Release("preview", new DateTime(2026, 5, 20), "Public preview", VisioTimelinePlacement.Below)
                    .Milestone("ga", new DateTime(2026, 6, 25), "GA", VisioTimelinePlacement.Above));
        }

        private static VisioDocument CreateRoutedDecisionFlow(string filePath) {
            VisioDocument document = VisioDocument.Create(filePath);
            VisioStyleTheme theme = VisioStyleTheme.Minimal();
            VisioPage page = document.AddPage("Routed Decision", 10.5, 6.5);

            VisioShape intake = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "intake", 1.8, 4.8, "Intake");
            VisioShape review = page.AddStencilShape(VisioStencils.Flowchart.Get("decision"), "review", 4.7, 4.8, "Valid?");
            VisioShape accept = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "accept", 7.6, 4.8, "Accept");
            VisioShape rework = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "rework", 4.7, 2.1, "Rework");
            VisioShape archive = page.AddStencilShape(VisioStencils.Flowchart.Get("data"), "archive", 7.6, 2.1, "Archive");

            page.SelectShapes(shape => shape.MasterNameU != null).Style(theme.Primary);
            review.ApplyStyle(theme.Decision);
            archive.ApplyStyle(theme.Success);

            page.AddConnector(intake, review, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .RouteOrthogonal()
                .ApplyStyle(theme.DataConnector);

            VisioConnector yes = page.AddConnector(review, accept, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .RouteOrthogonal()
                .PlaceLabel(0.55, offsetY: 0.18)
                .ApplyStyle(theme.Connector);
            yes.Label = "yes";

            VisioConnector no = page.AddConnector(review, rework, ConnectorKind.Dynamic, VisioSide.Bottom, VisioSide.Top)
                .RouteOrthogonal(VisioConnectorRouteStyle.VerticalThenHorizontal, -0.2)
                .PlaceLabel(0.55, offsetX: -0.25)
                .ApplyStyle(theme.ControlConnector);
            no.Label = "no";

            page.AddConnector(rework, archive, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .RouteThrough(VisioConnectorWaypoint.At(6.15, 2.1), VisioConnectorWaypoint.At(6.15, 2.8))
                .ApplyStyle(theme.ControlConnector);

            return document;
        }

        /// <summary>
        /// Creates a data-driven CI/CD inventory graph that demonstrates node, edge, cluster, Shape Data, hyperlink, stencil, and legend support.
        /// </summary>
        /// <param name="filePath">Target VSDX file path.</param>
        public static VisioDocument CreateCiCdInventoryGraph(string filePath) {
            VisioGraphNodeRecord engineer = CreateNode("engineer", "Engineer", VisioStencils.CollaborationBusiness, "person", "actor");
            engineer.IsRoot = true;
            engineer.ShapeData.Add("Owner", "Platform Engineering");

            VisioGraphNodeRecord repository = CreateNode("repo", "Source Repo", VisioStencils.CollaborationBusiness, "document", "record");
            repository.ShapeData.Add("System", "Git");

            VisioGraphNodeRecord pipeline = CreateNode("pipeline", "Build Pipeline", VisioStencils.DataPlatform, "pipeline", "job");
            pipeline.ShapeData.Add("Sla", "15 minutes");
            pipeline.HyperlinkAddress = "https://example.org/pipelines/customer-api";
            pipeline.HyperlinkDescription = "Pipeline definition";

            VisioGraphNodeRecord agent = CreateNode("agent", "Build Agent", VisioStencils.Infrastructure, "server", "compute");
            agent.ShapeData.Add("Pool", "Linux");

            VisioGraphNodeRecord registry = CreateNode("registry", "Image Registry", VisioStencils.DataPlatform, "catalog", "metadata");
            registry.ShapeData.Add("Retention", "90 days");

            VisioGraphNodeRecord cluster = CreateNode("cluster", "AKS Cluster", VisioStencils.ContainersKubernetes, "kubernetes", "aks");
            cluster.ShapeData.Add("Environment", "Production");

            VisioGraphNodeRecord secrets = CreateNode("secrets", "Secret Store", VisioStencils.Cloud, "secret", "vault");
            secrets.ShapeData.Add("Rotation", "30 days");

            VisioGraphNodeRecord monitor = CreateNode("monitor", "Observability", VisioStencils.Cloud, "monitoring", "metrics");
            monitor.ShapeData.Add("Signal", "Logs; Metrics; Alerts");

            VisioGraphEdgeRecord commit = new("engineer", "repo") {
                Label = "commit"
            };
            VisioGraphEdgeRecord trigger = new("repo", "pipeline") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "trigger"
            };
            VisioGraphEdgeRecord schedule = new("pipeline", "agent") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "run"
            };
            VisioGraphEdgeRecord publish = new("agent", "registry") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "image"
            };
            publish.ShapeData.Add("Protocol", "OCI");

            VisioGraphEdgeRecord deploy = new("registry", "cluster") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "deploy"
            };
            deploy.ShapeData.Add("Gate", "signed image");

            VisioGraphEdgeRecord secretFlow = new("secrets", "cluster") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "secrets"
            };
            VisioGraphEdgeRecord telemetry = new("cluster", "monitor") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "telemetry"
            };
            VisioGraphEdgeRecord onCall = new("engineer", "monitor") {
                Label = "on-call",
                Directed = false
            };

            VisioGraphClusterRecord delivery = new("delivery-cluster", "Delivery Control Plane", new[] { "repo", "pipeline", "agent", "registry" });
            delivery.ShapeData.Add("Owner", "DevEx");
            delivery.HyperlinkAddress = "https://example.org/runbooks/delivery";
            delivery.HyperlinkDescription = "Delivery runbook";

            VisioGraphClusterRecord runtime = new("runtime-cluster", "Runtime", new[] { "cluster", "secrets", "monitor" });
            runtime.ShapeData.Add("Owner", "SRE");
            runtime.HyperlinkAddress = "https://example.org/runbooks/runtime";
            runtime.HyperlinkDescription = "Runtime runbook";

            return VisioDocument.Create(filePath)
                .GraphDiagram("CI/CD Inventory Graph", graph => graph
                    .Title("CI/CD Pipeline and Runtime Inventory")
                    .Theme(VisioStyleTheme.Technical())
                    .Layout(VisioGraphLayout.Layered)
                    .Direction(VisioGraphDirection.LeftToRight)
                    .Legend()
                    .PageSize(16.5, 7.5)
                    .Margins(0.8, 0.85, 0.8, 0.75)
                    .NodeSize(1.35, 0.74)
                    .Spacing(0.72, 0.86)
                    .Import(
                        new[] { engineer, repository, pipeline, agent, registry, cluster, secrets, monitor },
                        new[] { commit, trigger, schedule, publish, deploy, secretFlow, telemetry, onCall },
                        new[] { delivery, runtime }));
        }

        /// <summary>
        /// Creates a data-driven identity authentication graph that demonstrates trust boundaries, control/data flows, Shape Data, and stencil profiles.
        /// </summary>
        /// <param name="filePath">Target VSDX file path.</param>
        public static VisioDocument CreateIdentityAuthenticationGraph(string filePath) {
            VisioGraphNodeRecord user = CreateNode("user", "User", VisioStencils.SecurityIdentity, "user", "person");
            user.IsRoot = true;
            user.ShapeData.Add("AuthType", "Interactive");

            VisioGraphNodeRecord device = CreateNode("device", "Managed Device", VisioStencils.Network, "workstation", "endpoint");
            device.ShapeData.Add("Compliance", "Required");

            VisioGraphNodeRecord app = CreateNode("app", "SaaS App", VisioStencils.CollaborationBusiness, "system", "application");
            app.ShapeData.Add("Audience", "Employees");

            VisioGraphNodeRecord idp = CreateNode("idp", "Identity Provider", VisioStencils.SecurityIdentity, "idp", "entra", "oidc");
            idp.ShapeData.Add("Protocol", "OIDC");
            idp.HyperlinkAddress = "https://example.org/runbooks/identity-provider";
            idp.HyperlinkDescription = "Identity provider runbook";

            VisioGraphNodeRecord policy = CreateNode("policy", "Conditional Access", VisioStencils.SecurityIdentity, "policy", "conditional-access");
            policy.ShapeData.Add("Decision", "Allow, challenge, or block");

            VisioGraphNodeRecord mfa = CreateNode("mfa", "MFA Challenge", VisioStencils.SecurityIdentity, "key", "credential");
            mfa.ShapeData.Add("Factor", "FIDO2 or app approval");

            VisioGraphNodeRecord groups = CreateNode("groups", "RBAC Groups", VisioStencils.SecurityIdentity, "group", "role");
            groups.ShapeData.Add("Source", "Directory groups");

            VisioGraphNodeRecord audit = CreateNode("audit", "Audit Log", VisioStencils.SecurityIdentity, "audit", "evidence");
            audit.ShapeData.Add("Retention", "365 days");
            audit.HyperlinkAddress = "https://example.org/security/audit";
            audit.HyperlinkDescription = "Audit workspace";

            VisioGraphEdgeRecord launch = new("user", "device") {
                Label = "sign in"
            };
            VisioGraphEdgeRecord request = new("device", "app") {
                Label = "access request"
            };
            VisioGraphEdgeRecord redirect = new("app", "idp") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "redirect"
            };
            redirect.ShapeData.Add("Protocol", "OIDC");

            VisioGraphEdgeRecord evaluate = new("idp", "policy") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "evaluate"
            };
            evaluate.ShapeData.Add("Signals", "user, device, risk");

            VisioGraphEdgeRecord challenge = new("policy", "mfa") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "challenge"
            };
            VisioGraphEdgeRecord claims = new("groups", "idp") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "claims"
            };
            VisioGraphEdgeRecord token = new("idp", "app") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "token"
            };
            token.ShapeData.Add("Lifetime", "60 minutes");

            VisioGraphEdgeRecord evidence = new("idp", "audit") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "sign-in log"
            };
            VisioGraphEdgeRecord review = new("groups", "audit") {
                Label = "access review",
                Directed = false
            };

            VisioGraphClusterRecord userContext = new("user-context", "User Context", new[] { "user", "device" });
            userContext.ShapeData.Add("Boundary", "Managed endpoint");

            VisioGraphClusterRecord trustBoundary = new("trust-boundary", "Identity Trust Boundary", new[] { "idp", "policy", "mfa", "groups", "audit" });
            trustBoundary.ShapeData.Add("Owner", "Identity Security");
            trustBoundary.HyperlinkAddress = "https://example.org/runbooks/conditional-access";
            trustBoundary.HyperlinkDescription = "Conditional access runbook";

            return VisioDocument.Create(filePath)
                .GraphDiagram("Identity Authentication Graph", graph => graph
                    .Title("Active Directory Identity Authentication Flow")
                    .Theme(VisioStyleTheme.Enterprise())
                    .Layout(VisioGraphLayout.Layered)
                    .Direction(VisioGraphDirection.LeftToRight)
                    .Legend()
                    .PageSize(17.2, 7.8)
                    .Margins(0.8, 0.85, 0.8, 0.75)
                    .NodeSize(1.35, 0.74)
                    .Spacing(0.78, 0.86)
                    .Import(
                        new[] { user, device, app, idp, policy, mfa, groups, audit },
                        new[] { launch, request, redirect, evaluate, challenge, claims, token, evidence, review },
                        new[] { userContext, trustBoundary }));
        }

        private static VisioGraphNodeRecord CreateNode(string id, string text, VisioStencilCatalog catalog, params string[] queries) {
            VisioGraphNodeRecord record = new(id, text) {
                StencilCatalog = catalog
            };
            foreach (string query in queries) {
                record.StencilQueries.Add(query);
            }

            return record;
        }
    }
}
