using System.Text.Json.Nodes;
using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows;

/// <summary>Exports bounded evidence and execution failures using SARIF 2.1.0.</summary>
public static class OfficeProvenanceSarif {
    /// <summary>Creates a SARIF run. Structural presence is evidence, not an authenticity or AI verdict.</summary>
    public static string Serialize(IReadOnlyList<OfficeProvenanceWorkflowResult> reports) {
        ArgumentNullException.ThrowIfNull(reports);
        var results = new JsonArray(); var artifacts = new JsonArray();
        foreach (OfficeProvenanceWorkflowResult report in reports) {
            int artifactIndex = artifacts.Count;
            string uri = ToUri(report.InputPath ?? report.RequestId);
            var artifact = new JsonObject { ["location"] = new JsonObject { ["uri"] = uri },
                ["hashes"] = report.InputSha256 == null ? null : new JsonObject { ["sha-256"] = report.InputSha256 },
                ["properties"] = new JsonObject {
                    ["workflowStatus"] = report.Status.ToString(),
                    ["textIntegrityStatus"] = report.Assessment?.TextIntegrityStatus.ToString() ?? "NotRequested",
                    ["verificationStatus"] = report.Assessment?.VerificationStatus.ToString() ?? "NotRequested",
                    ["providerSignalsStatus"] = report.Assessment?.ProviderSignalsStatus.ToString() ?? "NotRequested"
                } };
            if (report.InputSha256 == null) artifact.Remove("hashes");
            artifacts.Add(artifact);
            if (!report.Succeeded) Add("officeimo.execution", "error", report.Summary, null, null);
            foreach (OfficeProvenanceEvidence evidence in (report.Assessment?.Structural ?? report.Inspection)?.Evidence ?? Array.Empty<OfficeProvenanceEvidence>())
                Add("officeimo.carrier." + evidence.Carrier, "note", evidence.Carrier + " at " + evidence.Location + "; structural evidence only.", null, null);
            foreach (OfficeTextIntegrityFinding finding in report.Assessment?.TextIntegrity?.Findings ?? Array.Empty<OfficeTextIntegrityFinding>())
                Add("officeimo.text." + finding.Kind, finding.Risk == OfficeTextIntegrityRisk.PotentiallyDangerous ? "error" :
                    finding.Risk == OfficeTextIntegrityRisk.ContextDependent ? "warning" : "note",
                    finding.UnicodeNotation + " (" + finding.Kind + ", " + finding.Risk + ")", finding.TextOffset, finding.TextLength);
            if (report.Assessment?.VerificationStatus == OfficeProvenanceCheckStatus.Failed || report.Assessment?.ProviderSignalsStatus == OfficeProvenanceCheckStatus.Failed)
                Add("officeimo.provider", "error", "An optional provider check failed. Review the provider-specific result.", null, null);
            void Add(string rule, string level, string message, int? offset, int? length) {
                var location = new JsonObject { ["artifactLocation"] = new JsonObject { ["uri"] = uri, ["index"] = artifactIndex } };
                if (offset.HasValue) location["region"] = new JsonObject { ["charOffset"] = offset.Value, ["charLength"] = length!.Value };
                results.Add(new JsonObject { ["ruleId"] = rule, ["level"] = level,
                    ["message"] = new JsonObject { ["text"] = message },
                    ["locations"] = new JsonArray(new JsonObject { ["physicalLocation"] = location }) });
            }
        }
        return new JsonObject { ["$schema"] = "https://json.schemastore.org/sarif-2.1.0.json", ["version"] = "2.1.0",
            ["runs"] = new JsonArray(new JsonObject {
                ["tool"] = new JsonObject { ["driver"] = new JsonObject { ["name"] = "OfficeIMO", ["informationUri"] = "https://officeimo.com" } },
                ["columnKind"] = "utf16CodeUnits", ["artifacts"] = artifacts, ["results"] = results,
                ["invocations"] = new JsonArray(new JsonObject { ["executionSuccessful"] = reports.All(item => item.Succeeded &&
                    item.Assessment?.VerificationStatus != OfficeProvenanceCheckStatus.Failed && item.Assessment?.ProviderSignalsStatus != OfficeProvenanceCheckStatus.Failed) })
            }) }.ToJsonString();
    }
    private static string ToUri(string path) => Path.IsPathFullyQualified(path)
        ? new Uri(path).AbsoluteUri : string.Join('/', path.Replace('\\', '/').Split('/').Select(Uri.EscapeDataString));
}
