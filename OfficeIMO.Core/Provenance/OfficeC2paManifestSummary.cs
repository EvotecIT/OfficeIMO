using System;
using System.Collections.Generic;

namespace OfficeIMO.Provenance;

/// <summary>
/// What the active C2PA manifest in a Content Credentials store says about the asset: the tool that wrote it,
/// the recorded actions, ingredients, and the certificate it was signed with.
/// </summary>
/// <remarks>
/// These values are read from the manifest as written. They are not verified: the signature, certificate trust,
/// and content binding are not checked. Use a verification provider before treating them as proof.
/// </remarks>
public sealed class OfficeC2paManifestSummary {
    /// <summary>Creates a manifest summary.</summary>
    public OfficeC2paManifestSummary(
        string? label,
        string? claimGenerator,
        string? title,
        string? format,
        IReadOnlyList<OfficeC2paAction>? actions,
        IReadOnlyList<string>? ingredients,
        string? signedBy,
        string? certificateIssuer,
        int manifestCount) {
        Label = label;
        ClaimGenerator = claimGenerator;
        Title = title;
        Format = format;
        Actions = new List<OfficeC2paAction>(actions ?? Array.Empty<OfficeC2paAction>()).AsReadOnly();
        Ingredients = new List<string>(ingredients ?? Array.Empty<string>()).AsReadOnly();
        SignedBy = signedBy;
        CertificateIssuer = certificateIssuer;
        ManifestCount = manifestCount;
    }

    /// <summary>Gets the active manifest label, usually a URN.</summary>
    public string? Label { get; }
    /// <summary>Gets the application that wrote the claim, with its version when recorded (for example "ChatGPT").</summary>
    public string? ClaimGenerator { get; }
    /// <summary>Gets the asset title recorded in the claim.</summary>
    public string? Title { get; }
    /// <summary>Gets the media type recorded in the claim.</summary>
    public string? Format { get; }
    /// <summary>Gets the actions recorded by the active manifest, in order.</summary>
    public IReadOnlyList<OfficeC2paAction> Actions { get; }
    /// <summary>Gets the titles of ingredients (source assets) recorded by the active manifest.</summary>
    public IReadOnlyList<string> Ingredients { get; }
    /// <summary>Gets the organization or common name of the signing certificate's subject.</summary>
    public string? SignedBy { get; }
    /// <summary>Gets the organization or common name of the signing certificate's issuer.</summary>
    public string? CertificateIssuer { get; }
    /// <summary>Gets the number of manifests in the store; earlier manifests describe earlier editing steps.</summary>
    public int ManifestCount { get; }

    /// <summary>Gets whether any recorded action declares content produced by a trained generative model.</summary>
    public bool DeclaresGenerativeAi {
        get {
            foreach (OfficeC2paAction action in Actions) {
                if (action.DigitalSourceKind is OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia
                    or OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia) return true;
            }
            return false;
        }
    }
}

/// <summary>One action recorded in a C2PA actions assertion, such as <c>c2pa.created</c> or <c>c2pa.edited</c>.</summary>
public sealed class OfficeC2paAction {
    /// <summary>Creates an action record.</summary>
    public OfficeC2paAction(string action, string? softwareAgent, string? digitalSourceType, string? when) {
        Action = action ?? throw new ArgumentNullException(nameof(action));
        SoftwareAgent = softwareAgent;
        DigitalSourceType = digitalSourceType;
        DigitalSourceKind = digitalSourceType == null ? OfficeProvenanceDigitalSourceKind.Unknown : OfficeProvenanceXmp.Classify(digitalSourceType);
        When = when;
    }

    /// <summary>Gets the action identifier, for example <c>c2pa.created</c>.</summary>
    public string Action { get; }
    /// <summary>Gets the software that performed the action, with its version when recorded.</summary>
    public string? SoftwareAgent { get; }
    /// <summary>Gets the IPTC digital source type URI, when recorded.</summary>
    public string? DigitalSourceType { get; }
    /// <summary>Gets the classified digital source type.</summary>
    public OfficeProvenanceDigitalSourceKind DigitalSourceKind { get; }
    /// <summary>Gets the recorded time of the action, when present.</summary>
    public string? When { get; }
}
