using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Provenance;

/// <summary>Identifies a provider-specific provenance or watermark signal.</summary>
public enum OfficeProvenanceSignalKind {
    /// <summary>A durable watermark embedded in image, audio, or video content.</summary>
    DurableMediaWatermark,
    /// <summary>A statistical watermark encoded through generated text choices.</summary>
    StatisticalTextWatermark,
    /// <summary>A visible disclosure or label identified by a provider.</summary>
    VisibleDisclosure,
    /// <summary>A deterministic byte or markup artifact that does not establish authorship.</summary>
    DeterministicArtifact
}

/// <summary>Describes one provider's bounded detection outcome.</summary>
public enum OfficeProvenanceSignalStatus {
    /// <summary>The provider detected its signal.</summary>
    Detected,
    /// <summary>The provider completed detection and did not find its signal.</summary>
    NotDetected,
    /// <summary>The available content was insufficient for a reliable provider result.</summary>
    Inconclusive,
    /// <summary>The provider or required credential was unavailable.</summary>
    ProviderUnavailable,
    /// <summary>The provider failed to complete detection.</summary>
    Error
}

/// <summary>Normalized evidence returned by a vendor or watermark detector.</summary>
public sealed class OfficeProvenanceSignalResult {
    /// <summary>Creates a normalized provider signal result.</summary>
    public OfficeProvenanceSignalResult(
        string providerName,
        OfficeProvenanceSignalKind signalKind,
        OfficeProvenanceSignalStatus status,
        IReadOnlyList<string>? findings = null) {
        if (string.IsNullOrWhiteSpace(providerName)) throw new ArgumentException("A provider name is required.", nameof(providerName));
        ProviderName = providerName;
        SignalKind = signalKind;
        Status = status;
        Findings = new List<string>(findings ?? Array.Empty<string>()).AsReadOnly();
    }

    /// <summary>Gets the provider identity.</summary>
    public string ProviderName { get; }
    /// <summary>Gets the kind of signal this provider detects.</summary>
    public OfficeProvenanceSignalKind SignalKind { get; }
    /// <summary>Gets the normalized detection status.</summary>
    public OfficeProvenanceSignalStatus Status { get; }
    /// <summary>Gets provider findings without implying a universal AI verdict.</summary>
    public IReadOnlyList<string> Findings { get; }
}

/// <summary>Contract for provider-specific watermark or disclosure detection.</summary>
public interface IOfficeProvenanceSignalDetector {
    /// <summary>Gets the detector provider name.</summary>
    string Name { get; }
    /// <summary>Gets the signal kind detected by this provider.</summary>
    OfficeProvenanceSignalKind SignalKind { get; }
    /// <summary>Inspects one asset and returns a normalized provider result.</summary>
    OfficeProvenanceSignalResult Detect(string filePath);
}

/// <summary>Optional cancellation-aware extension for provider-specific signal detectors.</summary>
public interface ICancellableOfficeProvenanceSignalDetector : IOfficeProvenanceSignalDetector {
    /// <summary>Inspects one asset while observing cancellation and returns a normalized provider result.</summary>
    OfficeProvenanceSignalResult Detect(string filePath, CancellationToken cancellationToken);
}

/// <summary>Configures a combined provenance assessment.</summary>
public sealed class OfficeProvenanceAssessmentOptions {
    /// <summary>Gets structural provenance limits.</summary>
    public OfficeProvenanceOptions Structural { get; } = new OfficeProvenanceOptions();
    /// <summary>Gets Unicode text-integrity limits.</summary>
    public OfficeTextIntegrityOptions TextIntegrity { get; } = new OfficeTextIntegrityOptions();
    /// <summary>Gets optional cryptographic verification policy.</summary>
    public OfficeProvenanceVerificationOptions Verification { get; } = new OfficeProvenanceVerificationOptions();
    /// <summary>Whether text-like files receive Unicode integrity inspection. Defaults to true.</summary>
    public bool InspectTextIntegrity { get; set; } = true;
}

/// <summary>Combined structural, cryptographic, Unicode, and provider-specific evidence.</summary>
public sealed class OfficeProvenanceAssessmentReport {
    /// <summary>Creates a combined assessment without collapsing evidence into an authorship verdict.</summary>
    public OfficeProvenanceAssessmentReport(
        OfficeProvenanceReport structural,
        OfficeProvenanceVerificationResult? verification,
        OfficeTextIntegrityReport? textIntegrity,
        IReadOnlyList<OfficeProvenanceSignalResult>? providerSignals = null) {
        Structural = structural ?? throw new ArgumentNullException(nameof(structural));
        Verification = verification;
        TextIntegrity = textIntegrity;
        ProviderSignals = new List<OfficeProvenanceSignalResult>(providerSignals ?? Array.Empty<OfficeProvenanceSignalResult>()).AsReadOnly();
    }

    /// <summary>Gets dependency-free structural carrier evidence.</summary>
    public OfficeProvenanceReport Structural { get; }
    /// <summary>Gets optional cryptographic content-binding and trust evidence.</summary>
    public OfficeProvenanceVerificationResult? Verification { get; }
    /// <summary>Gets exact Unicode findings for text-like inputs.</summary>
    public OfficeTextIntegrityReport? TextIntegrity { get; }
    /// <summary>Gets provider-specific watermark or disclosure results.</summary>
    public IReadOnlyList<OfficeProvenanceSignalResult> ProviderSignals { get; }
    /// <summary>Gets whether the configured verifier validated a content credential.</summary>
    public bool HasVerifiedContentCredential => Verification?.Status == OfficeProvenanceVerificationStatus.Valid;
    /// <summary>Gets whether any configured provider detected its own signal.</summary>
    public bool HasDetectedProviderSignal => ProviderSignals.Any(item => item.Status == OfficeProvenanceSignalStatus.Detected);
}

/// <summary>Builds one evidence report while preserving the meaning and uncertainty of each signal.</summary>
public static class OfficeProvenanceAssessment {
    /// <summary>Assesses one file with optional cryptographic and vendor-specific providers.</summary>
    public static OfficeProvenanceAssessmentReport InspectFile(
        string filePath,
        OfficeProvenanceAssessmentOptions? options = null,
        IOfficeProvenanceVerifier? verifier = null,
        IEnumerable<IOfficeProvenanceSignalDetector>? signalDetectors = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        string fullPath = Path.GetFullPath(filePath);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("The asset to assess was not found.", fullPath);
        options ??= new OfficeProvenanceAssessmentOptions();
        IOfficeProvenanceSignalDetector[] detectors = (signalDetectors ?? Array.Empty<IOfficeProvenanceSignalDetector>())
            .Select(detector => detector ?? throw new ArgumentException(
                "Signal detector collections cannot contain null entries.", nameof(signalDetectors)))
            .ToArray();
        bool hasExternalProviders = verifier != null || detectors.Length != 0;

        using (OfficeProvenanceFileSnapshot snapshot = OfficeProvenanceFileSnapshot.Capture(
                   fullPath,
                   options.Structural.MaxAssetBytes)) {
            snapshot.SealForProviderAccess();
            OfficeProvenanceReport structural = OfficeProvenanceInspector.InspectFile(snapshot.FilePath, options.Structural);
            if (hasExternalProviders) {
                snapshot.CaptureExternalManifestDependencies(
                    fullPath,
                    structural,
                    options.Structural.MaxManifestBytes,
                    options.Structural.MaxExpandedContainerBytes);
            }
            OfficeProvenanceAssessmentReport assessment = AssessSnapshotFile(
                snapshot.FilePath,
                fullPath,
                structural,
                options,
                verifier,
                detectors);
            snapshot.VerifyPrimaryFile();
            if (hasExternalProviders) snapshot.VerifyExternalManifestDependencies();
            return assessment;
        }
    }

    /// <summary>Combines an existing structural report with optional text, cryptographic, and provider evidence.</summary>
    /// <remarks>
    /// This overload lets workflow hosts preserve format-owner structural inspection while keeping evidence
    /// composition and provider-result validation in the canonical provenance owner. The caller must ensure
    /// <paramref name="filePath"/> identifies the same immutable bytes used to create <paramref name="structural"/>.
    /// </remarks>
    public static OfficeProvenanceAssessmentReport AssessFile(
        string filePath,
        OfficeProvenanceReport structural,
        OfficeProvenanceAssessmentOptions? options = null,
        IOfficeProvenanceVerifier? verifier = null,
        IEnumerable<IOfficeProvenanceSignalDetector>? signalDetectors = null,
        CancellationToken cancellationToken = default) => AssessFileCore(
            filePath,
            filePath,
            structural,
            options,
            verifier,
            signalDetectors,
            cancellationToken);

    internal static OfficeProvenanceAssessmentReport AssessSnapshotFile(
        string snapshotFilePath,
        string logicalFilePath,
        OfficeProvenanceReport structural,
        OfficeProvenanceAssessmentOptions? options = null,
        IOfficeProvenanceVerifier? verifier = null,
        IEnumerable<IOfficeProvenanceSignalDetector>? signalDetectors = null,
        CancellationToken cancellationToken = default,
        Encoding? textEncoding = null) => AssessFileCore(
            snapshotFilePath,
            logicalFilePath,
            structural,
            options,
            verifier,
            signalDetectors,
            cancellationToken,
            textEncoding);

    private static OfficeProvenanceAssessmentReport AssessFileCore(
        string filePath,
        string logicalFilePath,
        OfficeProvenanceReport structural,
        OfficeProvenanceAssessmentOptions? options,
        IOfficeProvenanceVerifier? verifier,
        IEnumerable<IOfficeProvenanceSignalDetector>? signalDetectors,
        CancellationToken cancellationToken,
        Encoding? textEncoding = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A file path is required.", nameof(filePath));
        if (string.IsNullOrWhiteSpace(logicalFilePath)) throw new ArgumentException("A logical file path is required.", nameof(logicalFilePath));
        string fullPath = Path.GetFullPath(filePath);
        string logicalFullPath = Path.GetFullPath(logicalFilePath);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("The asset to assess was not found.", fullPath);
        if (structural == null) throw new ArgumentNullException(nameof(structural));
        options ??= new OfficeProvenanceAssessmentOptions();

        cancellationToken.ThrowIfCancellationRequested();
        OfficeTextIntegrityReport? textIntegrity = null;
        if (options.InspectTextIntegrity && IsTextLike(structural.Format)) {
            textIntegrity = OfficeTextIntegrityInspector.InspectFile(
                fullPath,
                options.TextIntegrity,
                logicalFullPath,
                textEncoding,
                cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
        }
        OfficeProvenanceVerificationResult? verification = verifier == null
            ? null
            : Verify(verifier, fullPath, options.Verification, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        if (verifier != null && verification == null) {
            throw OfficeProvenanceProviderContractException.Create(
                $"The '{verifier.Name}' provenance verifier returned no result.");
        }
        if (verifier != null && !string.Equals(verification!.ProviderName, verifier.Name, StringComparison.Ordinal)) {
            throw OfficeProvenanceProviderContractException.Create(
                $"The '{verifier.Name}' provenance verifier returned inconsistent provider metadata.");
        }
        var signals = new List<OfficeProvenanceSignalResult>();
        if (signalDetectors != null) {
            foreach (IOfficeProvenanceSignalDetector detector in signalDetectors) {
                cancellationToken.ThrowIfCancellationRequested();
                if (detector == null) throw new ArgumentException("Signal detector collections cannot contain null entries.", nameof(signalDetectors));
                OfficeProvenanceSignalResult result = Detect(detector, fullPath, cancellationToken) ??
                    throw OfficeProvenanceProviderContractException.Create(
                        $"The '{detector.Name}' signal detector returned no result.");
                cancellationToken.ThrowIfCancellationRequested();
                if (!string.Equals(result.ProviderName, detector.Name, StringComparison.Ordinal) || result.SignalKind != detector.SignalKind) {
                    throw OfficeProvenanceProviderContractException.Create(
                        $"The '{detector.Name}' signal detector returned inconsistent provider metadata.");
                }
                signals.Add(result);
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new OfficeProvenanceAssessmentReport(structural, verification, textIntegrity, signals.AsReadOnly());
    }

    private static bool IsTextLike(OfficeProvenanceAssetFormat format) =>
        format is OfficeProvenanceAssetFormat.StructuredText or OfficeProvenanceAssetFormat.UnstructuredText or
            OfficeProvenanceAssetFormat.Html or OfficeProvenanceAssetFormat.Svg;

    private static OfficeProvenanceVerificationResult Verify(
        IOfficeProvenanceVerifier verifier,
        string filePath,
        OfficeProvenanceVerificationOptions options,
        CancellationToken cancellationToken) => verifier is ICancellableOfficeProvenanceVerifier cancellable
            ? cancellable.Verify(filePath, options, cancellationToken)
            : verifier.Verify(filePath, options);

    private static OfficeProvenanceSignalResult Detect(
        IOfficeProvenanceSignalDetector detector,
        string filePath,
        CancellationToken cancellationToken) => detector is ICancellableOfficeProvenanceSignalDetector cancellable
            ? cancellable.Detect(filePath, cancellationToken)
            : detector.Detect(filePath);
}
