using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio {
    public static partial class VisioGallery {
        /// <summary>
        /// Creates a data-driven privileged-access review graph that demonstrates identity, approval, secret, ticket, and audit evidence flows.
        /// </summary>
        /// <param name="filePath">Target VSDX file path.</param>
        public static VisioDocument CreatePrivilegedAccessReviewGraph(string filePath) {
            VisioGraphNodeRecord requester = CreateNode("requester", "Requestor", VisioStencils.SecurityIdentity, "user", "person");
            requester.IsRoot = true;
            requester.ShapeData.Add("Role", "Engineer");

            VisioGraphNodeRecord ticket = CreateNode("ticket", "Access Ticket", VisioStencils.CollaborationBusiness, "document", "record");
            ticket.ShapeData.Add("Workflow", "JIT elevation");
            ticket.HyperlinkAddress = "https://example.org/access/tickets";
            ticket.HyperlinkDescription = "Access ticket queue";

            VisioGraphNodeRecord manager = CreateNode("manager", "Manager Approval", VisioStencils.CollaborationBusiness, "person", "actor");
            manager.ShapeData.Add("Sla", "4 hours");

            VisioGraphNodeRecord policy = CreateNode("policy", "Access Policy", VisioStencils.SecurityIdentity, "policy", "conditional-access");
            policy.ShapeData.Add("Decision", "Risk-based approval");

            VisioGraphNodeRecord pam = CreateNode("pam", "PAM Broker", VisioStencils.SecurityIdentity, "key", "credential");
            pam.ShapeData.Add("Mode", "Just-in-time");
            pam.HyperlinkAddress = "https://example.org/runbooks/privileged-access";
            pam.HyperlinkDescription = "Privileged access runbook";

            VisioGraphNodeRecord vault = CreateNode("vault", "Secret Vault", VisioStencils.Cloud, "secret", "vault");
            vault.ShapeData.Add("Rotation", "After checkout");

            VisioGraphNodeRecord target = CreateNode("target", "Privileged Target", VisioStencils.Infrastructure, "server", "compute");
            target.ShapeData.Add("Environment", "Production");

            VisioGraphNodeRecord audit = CreateNode("audit", "Audit Evidence", VisioStencils.SecurityIdentity, "audit", "evidence");
            audit.ShapeData.Add("Retention", "7 years");

            VisioGraphNodeRecord siem = CreateNode("siem", "SIEM Review", VisioStencils.Cloud, "monitoring", "metrics");
            siem.ShapeData.Add("Signal", "session, command, anomaly");

            VisioGraphEdgeRecord request = new("requester", "ticket") {
                Label = "request"
            };
            request.ShapeData.Add("Channel", "Portal");

            VisioGraphEdgeRecord approve = new("ticket", "manager") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "approve"
            };
            approve.ShapeData.Add("Sla", "4h");

            VisioGraphEdgeRecord evaluate = new("manager", "policy") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "evaluate"
            };
            evaluate.ShapeData.Add("Signals", "owner, risk, scope");

            VisioGraphEdgeRecord grant = new("policy", "pam") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "grant"
            };
            grant.ShapeData.Add("Duration", "60 minutes");

            VisioGraphEdgeRecord checkout = new("pam", "vault") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "checkout"
            };
            checkout.ShapeData.Add("SecretType", "ephemeral credential");

            VisioGraphEdgeRecord session = new("pam", "target") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "broker session"
            };
            session.ShapeData.Add("Protocol", "RDP/SSH");

            VisioGraphEdgeRecord sessionLog = new("pam", "audit") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "session log"
            };
            sessionLog.ShapeData.Add("Evidence", "recording");

            VisioGraphEdgeRecord alert = new("audit", "siem") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "analytics"
            };
            alert.ShapeData.Add("Rule", "Privileged anomaly");

            VisioGraphEdgeRecord review = new("manager", "audit") {
                Label = "recertify",
                Directed = false
            };

            VisioGraphClusterRecord requestCluster = new("request-cluster", "Request and Approval", new[] { "requester", "ticket", "manager" });
            requestCluster.ShapeData.Add("Owner", "Business Application");
            requestCluster.HyperlinkAddress = "https://example.org/runbooks/access-approval";
            requestCluster.HyperlinkDescription = "Approval runbook";

            VisioGraphClusterRecord accessCluster = new("access-control-cluster", "Privileged Access Control", new[] { "policy", "pam", "vault" });
            accessCluster.ShapeData.Add("Owner", "Identity Security");

            VisioGraphClusterRecord evidenceCluster = new("evidence-cluster", "Evidence and Review", new[] { "audit", "siem" });
            evidenceCluster.ShapeData.Add("Retention", "7 years");

            return VisioDocument.Create(filePath)
                .GraphDiagram("Privileged Access Review Graph", graph => graph
                    .Title("Privileged Access Review and Evidence Flow")
                    .Theme(VisioStyleTheme.Enterprise())
                    .Layout(VisioGraphLayout.Layered)
                    .Direction(VisioGraphDirection.LeftToRight)
                    .Legend()
                    .PageSize(18.5, 9.6)
                    .Margins(0.8, 0.95, 0.8, 0.85)
                    .NodeSize(1.38, 0.76)
                    .Spacing(0.78, 1.5)
                    .Import(
                        new[] { requester, ticket, manager, policy, pam, vault, target, audit, siem },
                        new[] { request, approve, evaluate, grant, checkout, session, sessionLog, alert, review },
                        new[] { requestCluster, accessCluster, evidenceCluster }));
        }
    }
}
