using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio {
    public static partial class VisioGallery {
        /// <summary>
        /// Creates a data-driven process governance review graph that demonstrates approval, evidence, policy, exception, and audit flows.
        /// </summary>
        /// <param name="filePath">Target VSDX file path.</param>
        public static VisioDocument CreateProcessGovernanceReviewGraph(string filePath) {
            VisioGraphNodeRecord requester = CreateNode("requester", "Business Request", VisioStencils.CollaborationBusiness, "document", "record");
            requester.IsRoot = true;
            requester.ShapeData.Add("Intake", "ServiceNow");

            VisioGraphNodeRecord triage = CreateNode("triage", "Triage Meeting", VisioStencils.CollaborationBusiness, "meeting", "workshop");
            triage.ShapeData.Add("Owner", "Process Office");

            VisioGraphNodeRecord policy = CreateNode("policy", "Policy Check", VisioStencils.SecurityIdentity, "policy", "control");
            policy.ShapeData.Add("Framework", "ISO 27001");

            VisioGraphNodeRecord riskDecision = CreateNode("risk-decision", "Risk Accepted?", VisioStencils.Flowchart, "decision", "branch");
            riskDecision.ShapeData.Add("Gate", "Risk acceptance");

            VisioGraphNodeRecord approval = CreateNode("approval", "CAB Approval", VisioStencils.CollaborationBusiness, "approval", "sign-off");
            approval.ShapeData.Add("Sla", "2 business days");
            approval.HyperlinkAddress = "https://example.org/process/cab";
            approval.HyperlinkDescription = "CAB review board";

            VisioGraphNodeRecord implementation = CreateNode("implementation", "Implement Change", VisioStencils.Flowchart, "process", "task");
            implementation.ShapeData.Add("Window", "Approved maintenance");

            VisioGraphNodeRecord exception = CreateNode("exception", "Exception Path", VisioStencils.Flowchart, "off-page", "connector");
            exception.ShapeData.Add("Escalation", "Executive sponsor");

            VisioGraphNodeRecord evidence = CreateNode("evidence", "Evidence Pack", VisioStencils.CollaborationBusiness, "document", "record");
            evidence.ShapeData.Add("Required", "Rollback, tests, approvals");

            VisioGraphNodeRecord audit = CreateNode("audit", "Audit Trail", VisioStencils.SecurityIdentity, "audit", "evidence");
            audit.ShapeData.Add("Retention", "7 years");

            VisioGraphNodeRecord notification = CreateNode("notification", "Stakeholder Notice", VisioStencils.CollaborationBusiness, "message", "notification");
            notification.ShapeData.Add("Audience", "Service owners");

            VisioGraphEdgeRecord submit = new("submit-request", "requester", "triage") {
                Label = "submit"
            };
            submit.ShapeData.Add("Channel", "ticket");

            VisioGraphEdgeRecord assess = new("triage-policy-assess", "triage", "policy") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "assess"
            };
            assess.ShapeData.Add("Checklist", "impact, owner, scope");

            VisioGraphEdgeRecord evaluate = new("policy-risk-evaluate", "policy", "risk-decision") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "evaluate"
            };
            evaluate.ShapeData.Add("Criterion", "policy fit");

            VisioGraphEdgeRecord approve = new("risk-approval-yes", "risk-decision", "approval") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "yes"
            };
            approve.ShapeData.Add("Decision", "approved");

            VisioGraphEdgeRecord implement = new("approval-implementation-authorize", "approval", "implementation") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "authorize"
            };
            implement.ShapeData.Add("Control", "CAB minutes");

            VisioGraphEdgeRecord exceptionPath = new("risk-exception-no", "risk-decision", "exception") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "no"
            };
            exceptionPath.ShapeData.Add("Decision", "exception required");

            VisioGraphEdgeRecord exceptionReview = new("exception-approval-sponsor", "exception", "approval") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "sponsor"
            };
            exceptionReview.ShapeData.Add("Required", "risk sign-off");

            VisioGraphEdgeRecord evidencePack = new("implementation-evidence-pack", "implementation", "evidence") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "evidence"
            };
            evidencePack.ShapeData.Add("Artifacts", "test log, backout proof");

            VisioGraphEdgeRecord archive = new("evidence-audit-archive", "evidence", "audit") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "archive"
            };
            archive.ShapeData.Add("Retention", "7 years");

            VisioGraphEdgeRecord notify = new("implementation-notification", "implementation", "notification") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "notify"
            };
            notify.ShapeData.Add("Timing", "pre/post change");

            VisioGraphClusterRecord intakeCluster = new("intake-cluster", "Intake and Assessment", new[] { "requester", "triage", "policy", "risk-decision" });
            intakeCluster.ShapeData.Add("Owner", "Process Office");
            intakeCluster.HyperlinkAddress = "https://example.org/runbooks/process-intake";
            intakeCluster.HyperlinkDescription = "Process intake runbook";

            VisioGraphClusterRecord governanceCluster = new("governance-cluster", "Governance and Execution", new[] { "approval", "implementation", "exception" });
            governanceCluster.ShapeData.Add("Owner", "Change Advisory Board");

            VisioGraphClusterRecord evidenceCluster = new("process-evidence-cluster", "Evidence, Notice, and Audit", new[] { "notification", "evidence", "audit" });
            evidenceCluster.ShapeData.Add("Retention", "7 years");

            return VisioDocument.Create(filePath)
                .GraphDiagram("Process Governance Review Graph", graph => graph
                    .Title("Change Governance and Evidence Flow")
                    .Theme(VisioStyleTheme.Process())
                    .Layout(VisioGraphLayout.Layered)
                    .Direction(VisioGraphDirection.LeftToRight)
                    .Legend()
                    .PageSize(20.6, 9.8)
                    .Margins(0.8, 0.9, 0.8, 0.85)
                    .NodeSize(1.42, 0.76)
                    .Spacing(0.92, 1.3)
                    .Import(
                        new[] { requester, triage, policy, riskDecision, approval, implementation, exception, evidence, audit, notification },
                        new[] { submit, assess, evaluate, approve, implement, exceptionPath, exceptionReview, evidencePack, archive, notify },
                        new[] { intakeCluster, governanceCluster, evidenceCluster }));
        }
    }
}
