using System;
using System.Collections.Generic;

namespace OfficeIMO.Provenance;

/// <summary>Identifies a standards-defined C2PA action emitted during signing.</summary>
public enum OfficeProvenanceActionKind {
    /// <summary>The asset was first created.</summary>
    Created,
    /// <summary>An existing parent asset was opened for editing.</summary>
    Opened,
    /// <summary>The asset received a generalized editorial change.</summary>
    Edited,
    /// <summary>Only asset metadata changed.</summary>
    EditedMetadata,
    /// <summary>The asset file format changed.</summary>
    Converted,
    /// <summary>The asset moved to another package or container without editorial change.</summary>
    Repackaged,
    /// <summary>The asset encoding changed without editorial change.</summary>
    Transcoded,
    /// <summary>Visible text was added.</summary>
    AddedText,
    /// <summary>Content was cropped.</summary>
    Cropped,
    /// <summary>Content dimensions changed.</summary>
    Resized,
    /// <summary>The asset was published to a wider audience.</summary>
    Published,
    /// <summary>An action occurred but cannot be classified more precisely.</summary>
    Unknown
}

/// <summary>One action and optional declared digital source type for a new C2PA claim.</summary>
public sealed class OfficeProvenanceAction {
    /// <summary>Creates a claim action.</summary>
    public OfficeProvenanceAction(
        OfficeProvenanceActionKind kind,
        OfficeProvenanceDigitalSourceKind digitalSourceKind = OfficeProvenanceDigitalSourceKind.Unknown) {
        if (!Enum.IsDefined(typeof(OfficeProvenanceActionKind), kind)) {
            throw new ArgumentOutOfRangeException(nameof(kind));
        }
        if (!Enum.IsDefined(typeof(OfficeProvenanceDigitalSourceKind), digitalSourceKind)) {
            throw new ArgumentOutOfRangeException(nameof(digitalSourceKind));
        }
        Kind = kind;
        DigitalSourceKind = digitalSourceKind;
    }

    /// <summary>Gets the standards-defined action.</summary>
    public OfficeProvenanceActionKind Kind { get; }
    /// <summary>Gets the optional declared source type.</summary>
    public OfficeProvenanceDigitalSourceKind DigitalSourceKind { get; }
}

/// <summary>Application-controlled values for a new C2PA claim.</summary>
public sealed class OfficeProvenanceClaim {
    /// <summary>Creates a claim with one or more ordered actions.</summary>
    public OfficeProvenanceClaim(
        string claimGenerator,
        IReadOnlyList<OfficeProvenanceAction> actions,
        string? title = null) {
        if (string.IsNullOrWhiteSpace(claimGenerator)) throw new ArgumentException("A claim generator is required.", nameof(claimGenerator));
        if (claimGenerator.Length > 256 || ContainsControlCharacter(claimGenerator)) {
            throw new ArgumentException("The claim generator must be at most 256 characters and contain no controls.", nameof(claimGenerator));
        }
        if (title != null && (title.Length > 1024 || ContainsControlCharacter(title))) {
            throw new ArgumentException("The title must be at most 1,024 characters and contain no controls.", nameof(title));
        }
        if (actions == null) throw new ArgumentNullException(nameof(actions));
        if (actions.Count == 0) throw new ArgumentException("At least one provenance action is required.", nameof(actions));
        for (int index = 0; index < actions.Count; index++) {
            if (actions[index] == null) throw new ArgumentException("Provenance actions cannot contain null entries.", nameof(actions));
        }
        ClaimGenerator = claimGenerator;
        Actions = new List<OfficeProvenanceAction>(actions).AsReadOnly();
        Title = title;
    }

    /// <summary>Gets the claim-generator user agent.</summary>
    public string ClaimGenerator { get; }
    /// <summary>Gets the optional human-readable asset title.</summary>
    public string? Title { get; }
    /// <summary>Gets actions in the order supplied by the application.</summary>
    public IReadOnlyList<OfficeProvenanceAction> Actions { get; }

    private static bool ContainsControlCharacter(string value) {
        foreach (char character in value) if (char.IsControl(character)) return true;
        return false;
    }
}

/// <summary>One file-based C2PA signing request.</summary>
public sealed class OfficeProvenanceSigningRequest {
    /// <summary>Creates a signing request for an input asset and a separate output path.</summary>
    public OfficeProvenanceSigningRequest(
        string inputPath,
        string outputPath,
        OfficeProvenanceClaim claim,
        string? parentPath = null) {
        if (string.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("An input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        InputPath = inputPath;
        OutputPath = outputPath;
        Claim = claim ?? throw new ArgumentNullException(nameof(claim));
        ParentPath = parentPath;
    }

    /// <summary>Gets the unsigned source asset.</summary>
    public string InputPath { get; }
    /// <summary>Gets the destination for the signed asset.</summary>
    public string OutputPath { get; }
    /// <summary>Gets claim values controlled by the application.</summary>
    public OfficeProvenanceClaim Claim { get; }
    /// <summary>Gets the optional parent asset recorded as a parent ingredient.</summary>
    public string? ParentPath { get; }
}

/// <summary>Bounds provider-backed C2PA signing.</summary>
public sealed class OfficeProvenanceSigningOptions {
    /// <summary>Maximum provider execution time. Defaults to two minutes.</summary>
    public TimeSpan Timeout { get; set; } = TimeSpan.FromMinutes(2);
    /// <summary>Maximum provider output captured. Defaults to 8 MiB.</summary>
    public long MaxReportBytes { get; set; } = 8L * 1024L * 1024L;
    /// <summary>Whether provider output is retained in the result. Defaults to false.</summary>
    public bool IncludeRawReport { get; set; }
    /// <summary>Whether an existing destination may be atomically replaced. Defaults to true.</summary>
    public bool ReplaceExistingOutput { get; set; } = true;
}

/// <summary>Describes the normalized result of provider-backed signing.</summary>
public enum OfficeProvenanceSigningStatus {
    /// <summary>The asset was signed and committed to the requested destination.</summary>
    Signed,
    /// <summary>The signing provider or configured signer was unavailable.</summary>
    ProviderUnavailable,
    /// <summary>The provider rejected the input, claim, or output format.</summary>
    Rejected,
    /// <summary>The provider failed to complete the operation.</summary>
    Error
}

/// <summary>Normalized evidence from one C2PA signing operation.</summary>
public sealed class OfficeProvenanceSigningResult {
    /// <summary>Creates a signing result.</summary>
    public OfficeProvenanceSigningResult(
        OfficeProvenanceSigningStatus status,
        string providerName,
        IReadOnlyList<string> findings,
        string? outputPath = null,
        OfficeProvenanceReport? structuralReport = null,
        string? rawReport = null) {
        if (string.IsNullOrWhiteSpace(providerName)) throw new ArgumentException("A provider name is required.", nameof(providerName));
        Status = status;
        ProviderName = providerName;
        Findings = new List<string>(findings ?? throw new ArgumentNullException(nameof(findings))).AsReadOnly();
        OutputPath = outputPath;
        StructuralReport = structuralReport;
        RawReport = rawReport;
    }

    /// <summary>Gets the normalized signing outcome.</summary>
    public OfficeProvenanceSigningStatus Status { get; }
    /// <summary>Gets the signing provider name.</summary>
    public string ProviderName { get; }
    /// <summary>Gets provider findings.</summary>
    public IReadOnlyList<string> Findings { get; }
    /// <summary>Gets the committed output path on success.</summary>
    public string? OutputPath { get; }
    /// <summary>Gets structural evidence from the completed output on success.</summary>
    public OfficeProvenanceReport? StructuralReport { get; }
    /// <summary>Gets bounded raw provider output when explicitly requested.</summary>
    public string? RawReport { get; }
    /// <summary>Gets whether signing completed successfully.</summary>
    public bool Succeeded => Status == OfficeProvenanceSigningStatus.Signed;
}

/// <summary>Optional provider contract for creating and signing Content Credentials.</summary>
public interface IOfficeProvenanceSigner {
    /// <summary>Gets the signing provider name.</summary>
    string Name { get; }
    /// <summary>Signs an asset and atomically commits the requested output.</summary>
    OfficeProvenanceSigningResult Sign(
        OfficeProvenanceSigningRequest request,
        OfficeProvenanceSigningOptions? options = null);
}

/// <summary>Maps normalized source declarations to current IPTC Digital Source Type URIs.</summary>
public static class OfficeProvenanceDigitalSourceTypes {
    private const string Prefix = "http://cv.iptc.org/newscodes/digitalsourcetype/";

    /// <summary>Returns the standard URI when the normalized kind has one unambiguous mapping.</summary>
    public static bool TryGetUri(OfficeProvenanceDigitalSourceKind kind, out string? uri) {
        switch (kind) {
            case OfficeProvenanceDigitalSourceKind.DigitalCapture:
                uri = Prefix + "digitalCapture";
                return true;
            case OfficeProvenanceDigitalSourceKind.AlgorithmicMedia:
                uri = Prefix + "algorithmicMedia";
                return true;
            case OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia:
                uri = Prefix + "trainedAlgorithmicMedia";
                return true;
            case OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia:
                uri = Prefix + "compositeWithTrainedAlgorithmicMedia";
                return true;
            case OfficeProvenanceDigitalSourceKind.CompositeCapture:
                uri = Prefix + "compositeCapture";
                return true;
            default:
                uri = null;
                return false;
        }
    }
}
