namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    internal const int DefaultAccessibilityFindings = 200;

    internal Task<AgentAccessibilityResult> InspectAccessibilityAsync(string path, bool strict = false, bool includeSarif = false,
        int maxFindings = DefaultAccessibilityFindings, CancellationToken cancellationToken = default) {
        if (maxFindings < 1 || maxFindings > 5000) throw new AgentUsageException("maxFindings must be between 1 and 5000.");
        string input = _pathPolicy.ResolveInput(path);
        if (!File.Exists(input)) throw new FileNotFoundException("Input file was not found.", input);
        if (!OfficeDocument.SupportsAccessibility(input)) {
            throw new AgentUsageException("Accessibility inspection supports Word, Excel and PowerPoint files.");
        }
        cancellationToken.ThrowIfCancellationRequested();
        OfficeAccessibilityReport report = OfficeDocument.InspectAccessibility(input,
            strict ? OfficeAccessibilityPolicyProfile.Strict : OfficeAccessibilityPolicyProfile.Default, cancellationToken: cancellationToken);
        var findings = new List<AgentAccessibilityFinding>(Math.Min(report.Findings.Count, maxFindings));
        foreach (OfficeAccessibilityFinding finding in report.Findings.Take(maxFindings)) {
            findings.Add(new AgentAccessibilityFinding {
                Severity = finding.Severity.ToString(),
                Code = finding.Code,
                Message = finding.Message,
                Location = finding.Location,
                LocationIndex = finding.LocationIndex,
                ObjectId = finding.ObjectId,
                ObjectName = finding.ObjectName,
                IsHeuristic = finding.IsHeuristic,
                MeasuredValue = finding.MeasuredValue,
                RequiredValue = finding.RequiredValue
            });
        }
        return Task.FromResult(new AgentAccessibilityResult {
            Path = input,
            Format = report.Format,
            Profile = report.Profile.ToString(),
            SchemaVersion = report.SchemaVersion,
            ErrorCount = report.ErrorCount,
            WarningCount = report.WarningCount,
            IsSuccessful = report.IsSuccessful,
            FindingCount = report.Findings.Count,
            Findings = findings,
            Truncated = report.Findings.Count > findings.Count,
            Sarif = includeSarif ? report.ToSarif(input) : null
        });
    }
}
