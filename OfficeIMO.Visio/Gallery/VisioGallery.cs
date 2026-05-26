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
    }
}
