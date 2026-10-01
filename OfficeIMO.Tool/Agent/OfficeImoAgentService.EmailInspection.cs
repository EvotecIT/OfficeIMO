using OfficeIMO.Email;
using OfficeIMO.Email.Data;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    internal Task<AgentEmailInspectResult> InspectEmailDataAsync(string path,
        int maxOutputCharacters = DefaultInspectOutputCharacters, CancellationToken cancellationToken = default) {
        maxOutputCharacters = ValidateOutputBudget(maxOutputCharacters);
        string inputPath = _pathPolicy.ResolveInput(path);
        var policy = new EmailDataInspectionOptions();
        AgentSourceRegistration source = _registry.RegisterEmailData(inputPath, policy.OpenOptions, cancellationToken);
        EmailHtmlDataInspectionReport inspection = EmailHtmlDataInspector.Inspect(inputPath, policy, cancellationToken: cancellationToken);
        // Bind the reported source id to the bytes/catalog that remain after inspection.
        _registry.Resolve(source.SourceId, inputPath, cancellationToken);
        var metadata = inspection.Metadata;
        var result = new AgentEmailInspectResult {
            SourceId = source.SourceId, Path = inputPath, Kind = metadata.Kind.ToString(), Format = metadata.Format,
            ProtectionKind = metadata.ProtectionKind, SignatureStatus = metadata.MessageInspected ? "Unverified" : null,
            BodyCount = metadata.MessageInspected ? metadata.Bodies.Count : null,
            AttachmentCount = metadata.MessageInspected ? metadata.AttachmentCount : null,
            ContainerCount = metadata.ContainerCount > 0 ? metadata.ContainerCount : null,
            DeclaredItemCount = metadata.DeclaredItemCount,
            ContentLineRootCount = metadata.ContentLineRootCount > 0 ? metadata.ContentLineRootCount : null,
            DiagnosticCount = metadata.DiagnosticCount, HtmlInspectionStatus = inspection.HtmlInspectionStatus,
            FindingSampleCount = inspection.HtmlInspectionStatus == "Completed" ? inspection.Findings.Count : null,
            FindingLimitMayHaveBeenReached = inspection.HtmlInspectionStatus == "Completed" ? inspection.FindingLimitMayHaveBeenReached : null,
            BlockedElementCount = inspection.HtmlInspectionStatus == "Completed" ? inspection.BlockedElementCount : null,
            EventHandlerAttributeCount = inspection.HtmlInspectionStatus == "Completed" ? inspection.EventHandlerAttributeCount : null,
            Details = inspection, Truncated = metadata.AttachmentsTruncated || metadata.ContainersTruncated ||
                metadata.DiagnosticsTruncated || metadata.HeaderScanTruncated || inspection.FindingLimitMayHaveBeenReached ||
                inspection.HtmlInspectionStatus == "BodyLimitExceeded" || inspection.HtmlInspectionStatus == "InspectionUnavailable"
        };
        if (AgentJson.Measure(result) > maxOutputCharacters) { result.Details = null; result.Truncated = true; }
        if (AgentJson.Measure(result) > maxOutputCharacters) { result.Path = null; result.Truncated = true; }
        EnsureWithinBudget(result, maxOutputCharacters);
        return Task.FromResult(result);
    }
}
