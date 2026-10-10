using OfficeIMO.Workflows;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>Presentation of one active credential using the canonical report transport.</summary>
public sealed class ProvenanceCredentialViewModel {
    /// <summary>Creates a view of the statements recorded in one manifest.</summary>
    public ProvenanceCredentialViewModel(string location, ProvenanceManifestDto manifest) {
        IStudioLocalizer localizer = StudioLocalization.Current;
        Location = location;
        Generator = manifest.ClaimGenerator ?? localizer.Get("Provenance.NotRecorded");
        Asset = string.Join(" · ", new[] { manifest.Title, manifest.Format }.Where(value => !string.IsNullOrWhiteSpace(value)));
        CertificateSubject = manifest.SignedBy ?? localizer.Get("Provenance.NotRecorded");
        CertificateIssuer = manifest.CertificateIssuer ?? localizer.Get("Provenance.NotRecorded");
        CertificateSubjectDescription = localizer.Format("Provenance.CertificateSubjectFormat", CertificateSubject);
        CertificateIssuerDescription = localizer.Format("Provenance.CertificateIssuerFormat", CertificateIssuer);
        ManifestCount = manifest.ManifestCount;
        DeclaresGenerativeAi = manifest.DeclaresGenerativeAi;
        ManifestCountDescription = localizer.Format("Provenance.ManifestCountFormat", ManifestCount);
        Label = manifest.Label ?? localizer.Get("Provenance.NotRecorded");
        Ingredients = manifest.Ingredients;
        IngredientDescriptions = manifest.Ingredients.Select(value => localizer.Format("Provenance.SourceAssetFormat", value)).ToArray();
        Actions = manifest.Actions.Select((action, index) => new ProvenanceCredentialActionViewModel(
            index + 1, ActionName(action.Action, localizer), action.SoftwareAgent ?? localizer.Get("Provenance.SoftwareNotRecorded"),
            action.When is { Length: > 0 } when ? localizer.Format("Provenance.RecordedTime", when) : localizer.Get("Provenance.TimeNotRecorded"),
            SourceType(action, localizer),
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
    /// <summary>Gets the localized certificate subject caption without asserting verification.</summary>
    public string CertificateSubjectDescription { get; }
    /// <summary>Gets the localized certificate issuer caption without asserting verification.</summary>
    public string CertificateIssuerDescription { get; }
    /// <summary>Gets the active manifest label.</summary>
    public string Label { get; }
    /// <summary>Gets the number of manifests present, including earlier records not expanded here.</summary>
    public int ManifestCount { get; }
    /// <summary>Gets the claim's aggregate AI declaration, including actions beyond the displayed timeline.</summary>
    public bool DeclaresGenerativeAi { get; }
    /// <summary>Gets the localized count of manifests in the store.</summary>
    public string ManifestCountDescription { get; }
    /// <summary>Gets the source asset titles recorded as ingredients.</summary>
    public IReadOnlyList<string> Ingredients { get; }
    /// <summary>Gets localized captions for the recorded source assets.</summary>
    public IReadOnlyList<string> IngredientDescriptions { get; }
    /// <summary>Gets the active claim's actions in their recorded order.</summary>
    public IReadOnlyList<ProvenanceCredentialActionViewModel> Actions { get; }

    private static string ActionName(string action, IStudioLocalizer localizer) => action switch {
        "c2pa.created" => localizer.Get("Provenance.ActionCreated"), "c2pa.opened" => localizer.Get("Provenance.ActionOpened"), "c2pa.edited" => localizer.Get("Provenance.ActionEdited"),
        "c2pa.converted" or "c2pa.transcoded" => localizer.Get("Provenance.ActionConverted"), "c2pa.cropped" => localizer.Get("Provenance.ActionCropped"),
        "c2pa.resized" => localizer.Get("Provenance.ActionResized"), "c2pa.placed" => localizer.Get("Provenance.ActionPlaced"), "c2pa.published" => localizer.Get("Provenance.ActionPublished"),
        _ when IsWatermark(action) => localizer.Get("Provenance.ActionWatermarked"),
        _ => action.StartsWith("c2pa.", StringComparison.Ordinal) ? action[5..].Replace('_', ' ') : action
    };
    private static bool IsWatermark(string action) => action == "c2pa.watermarked" || action.StartsWith("c2pa.watermarked.", StringComparison.Ordinal);
    private static string SourceType(ProvenanceActionDto action, IStudioLocalizer localizer) => action.DigitalSourceKind switch {
        "TrainedAlgorithmicMedia" => localizer.Get("Provenance.SourceGenerativeAi"),
        "CompositeWithTrainedAlgorithmicMedia" => localizer.Get("Provenance.SourceCompositeAi"),
        "AlgorithmicMedia" => localizer.Get("Provenance.SourceAlgorithmic"),
        "DigitalCapture" => localizer.Get("Provenance.SourceDigitalCapture"),
        "CompositeCapture" => localizer.Get("Provenance.SourceCompositeCapture"),
        _ => action.DigitalSourceType is { Length: > 0 } value ? localizer.Format("Provenance.SourceRecorded", value) : localizer.Get("Provenance.SourceNotRecorded")
    };
}

/// <summary>One recorded action in a credential timeline; its position does not assert a verified chronology.</summary>
public sealed record ProvenanceCredentialActionViewModel(int Position, string Action, string Software, string Time, string SourceType, bool HasWatermark);
