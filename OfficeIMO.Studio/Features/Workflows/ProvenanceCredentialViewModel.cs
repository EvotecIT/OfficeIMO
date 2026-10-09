using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>Presentation of one active credential using the canonical report transport.</summary>
public sealed class ProvenanceCredentialViewModel {
    /// <summary>Creates a view of the statements recorded in one manifest.</summary>
    public ProvenanceCredentialViewModel(string location, ProvenanceManifestDto manifest) {
        Location = location;
        Generator = manifest.ClaimGenerator ?? "Not recorded";
        Asset = string.Join(" · ", new[] { manifest.Title, manifest.Format }.Where(value => !string.IsNullOrWhiteSpace(value)));
        CertificateSubject = manifest.SignedBy ?? "Not recorded";
        CertificateIssuer = manifest.CertificateIssuer ?? "Not recorded";
        ManifestCount = manifest.ManifestCount;
        Label = manifest.Label ?? "Not recorded";
        Ingredients = manifest.Ingredients;
        Actions = manifest.Actions.Select((action, index) => new ProvenanceCredentialActionViewModel(
            index + 1, ActionName(action.Action), action.SoftwareAgent ?? "Software not recorded",
            action.When is { Length: > 0 } when ? "Recorded time: " + when : "Time not recorded",
            SourceType(action),
            IsWatermark(action.Action))).ToArray();
    }
    /// <summary>Gets the location of the credential carrier.</summary>
    public string Location { get; }
    /// <summary>Gets the application recorded by the claim.</summary>
    public string Generator { get; }
    /// <summary>Gets the recorded asset title and media type.</summary>
    public string Asset { get; }
    /// <summary>Gets the certificate subject, without asserting trust or signature validity.</summary>
    public string CertificateSubject { get; }
    /// <summary>Gets the certificate issuer, as recorded.</summary>
    public string CertificateIssuer { get; }
    /// <summary>Gets the active manifest label.</summary>
    public string Label { get; }
    /// <summary>Gets the number of manifests present, including earlier records not expanded here.</summary>
    public int ManifestCount { get; }
    /// <summary>Gets the source asset titles recorded as ingredients.</summary>
    public IReadOnlyList<string> Ingredients { get; }
    /// <summary>Gets the active claim's actions in their recorded order.</summary>
    public IReadOnlyList<ProvenanceCredentialActionViewModel> Actions { get; }

    private static string ActionName(string action) => action switch {
        "c2pa.created" => "Created", "c2pa.opened" => "Opened", "c2pa.edited" => "Edited",
        "c2pa.converted" or "c2pa.transcoded" => "Converted", "c2pa.cropped" => "Cropped",
        "c2pa.resized" => "Resized", "c2pa.placed" => "Placed other content", "c2pa.published" => "Published",
        _ when IsWatermark(action) => "Added a watermark to the content",
        _ => action.StartsWith("c2pa.", StringComparison.Ordinal) ? action[5..].Replace('_', ' ') : action
    };
    private static bool IsWatermark(string action) => action == "c2pa.watermarked" || action.StartsWith("c2pa.watermarked.", StringComparison.Ordinal);
    private static string SourceType(ProvenanceActionDto action) => action.DigitalSourceKind switch {
        "TrainedAlgorithmicMedia" => "Generative AI declared",
        "CompositeWithTrainedAlgorithmicMedia" => "Mix of AI-generated and other content declared",
        "AlgorithmicMedia" => "Software-generated content declared",
        "DigitalCapture" => "Camera or scanner capture declared",
        "CompositeCapture" => "Composite of captured media declared",
        _ => action.DigitalSourceType is { Length: > 0 } value ? "Recorded source type: " + value : "Source type not recorded"
    };
}

/// <summary>One recorded action in a credential timeline; its position does not assert a verified chronology.</summary>
public sealed record ProvenanceCredentialActionViewModel(int Position, string Action, string Software, string Time, string SourceType, bool HasWatermark);
