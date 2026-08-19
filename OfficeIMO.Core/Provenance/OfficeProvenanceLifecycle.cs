using System;
using System.IO;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Provenance;

/// <summary>Controls how provenance is handled when an asset is transformed.</summary>
public enum OfficeProvenanceTransformationPolicy {
    /// <summary>Preserves a credentialed source only when candidate bytes are unchanged; otherwise blocks.</summary>
    PreserveIfUnchanged,
    /// <summary>Removes carried credentials from changed output and records before/after evidence.</summary>
    RemoveInvalidated,
    /// <summary>Creates a new signed output with the source recorded as a parent ingredient.</summary>
    SignAsDerived
}

/// <summary>Configures provenance disposition for one transformation.</summary>
public sealed class OfficeProvenanceTransformationOptions {
    private readonly OfficeProvenanceRemovalOptions _removal = new OfficeProvenanceRemovalOptions {
        RemoveAiSourceMetadata = false
    };

    /// <summary>Gets or sets the disposition policy. Defaults to fail-closed preservation.</summary>
    public OfficeProvenanceTransformationPolicy Policy { get; set; } = OfficeProvenanceTransformationPolicy.PreserveIfUnchanged;
    /// <summary>Gets structural inspection limits.</summary>
    public OfficeProvenanceOptions Inspection { get; } = new OfficeProvenanceOptions();
    /// <summary>Gets removal policy used by <see cref="OfficeProvenanceTransformationPolicy.RemoveInvalidated"/>.</summary>
    public OfficeProvenanceRemovalOptions Removal => _removal;
    /// <summary>Gets signing limits used by <see cref="OfficeProvenanceTransformationPolicy.SignAsDerived"/>.</summary>
    public OfficeProvenanceSigningOptions Signing { get; } = new OfficeProvenanceSigningOptions();
    /// <summary>Gets or sets the claim required by derived signing.</summary>
    public OfficeProvenanceClaim? Claim { get; set; }
}

/// <summary>Audit evidence for one completed provenance-aware transformation.</summary>
public sealed class OfficeProvenanceTransformationResult {
    /// <summary>Creates transformation audit evidence.</summary>
    public OfficeProvenanceTransformationResult(
        OfficeProvenanceTransformationPolicy policy,
        bool contentChanged,
        OfficeProvenanceReport source,
        OfficeProvenanceReport candidate,
        OfficeProvenanceReport output,
        OfficeProvenanceRemovalResult? removal = null,
        OfficeProvenanceSigningResult? signing = null) {
        Policy = policy;
        ContentChanged = contentChanged;
        Source = source ?? throw new ArgumentNullException(nameof(source));
        Candidate = candidate ?? throw new ArgumentNullException(nameof(candidate));
        Output = output ?? throw new ArgumentNullException(nameof(output));
        Removal = removal;
        Signing = signing;
    }

    /// <summary>Gets the applied policy.</summary>
    public OfficeProvenanceTransformationPolicy Policy { get; }
    /// <summary>Gets whether source and candidate bytes differed before provenance disposition.</summary>
    public bool ContentChanged { get; }
    /// <summary>Gets source structural evidence.</summary>
    public OfficeProvenanceReport Source { get; }
    /// <summary>Gets candidate structural evidence before disposition.</summary>
    public OfficeProvenanceReport Candidate { get; }
    /// <summary>Gets committed output structural evidence.</summary>
    public OfficeProvenanceReport Output { get; }
    /// <summary>Gets selective removal evidence when that policy was used.</summary>
    public OfficeProvenanceRemovalResult? Removal { get; }
    /// <summary>Gets provider signing evidence when derived signing was used.</summary>
    public OfficeProvenanceSigningResult? Signing { get; }
}

/// <summary>Applies explicit, auditable provenance disposition after conversion or editing.</summary>
public static class OfficeProvenanceLifecycle {
    /// <summary>
    /// Finalizes a candidate asset. A changed credentialed source is never silently copied with the
    /// default policy, and a derived credential always records the source as its parent.
    /// </summary>
    public static OfficeProvenanceTransformationResult FinalizeFile(
        string sourcePath,
        string candidatePath,
        string outputPath,
        OfficeProvenanceTransformationOptions? options = null,
        IOfficeProvenanceSigner? signer = null) {
        string source = RequireFile(sourcePath, nameof(sourcePath));
        string candidate = RequireFile(candidatePath, nameof(candidatePath));
        if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("An output path is required.", nameof(outputPath));
        string output = Path.GetFullPath(outputPath);
        options ??= new OfficeProvenanceTransformationOptions();

        byte[] sourceBytes = ReadBounded(source, options.Inspection.MaxAssetBytes);
        byte[] candidateBytes = ReadBounded(candidate, options.Inspection.MaxAssetBytes);
        bool changed = !BytesEqual(sourceBytes, candidateBytes);
        if (changed && PathsEqual(source, output)) {
            throw new InvalidOperationException("A transformed candidate cannot overwrite its provenance source.");
        }

        OfficeProvenanceReport sourceReport = OfficeProvenanceInspector.Inspect(sourceBytes, source, options.Inspection);
        OfficeProvenanceReport candidateReport = OfficeProvenanceInspector.Inspect(candidateBytes, candidate, options.Inspection);
        switch (options.Policy) {
            case OfficeProvenanceTransformationPolicy.PreserveIfUnchanged:
                if (changed && HasContentCredentialCarrier(sourceReport)) {
                    throw new InvalidOperationException(
                        "The source carries a Content Credential. Select removal or derived signing before committing changed output.");
                }
                OfficeFileCommit.WriteAllBytes(output, candidateBytes);
                return new OfficeProvenanceTransformationResult(
                    options.Policy,
                    changed,
                    sourceReport,
                    candidateReport,
                    OfficeProvenanceInspector.Inspect(candidateBytes, output, options.Inspection));

            case OfficeProvenanceTransformationPolicy.RemoveInvalidated:
                OfficeProvenanceRemovalResult removal = OfficeProvenanceRemover.Remove(candidateBytes, candidate, options.Removal);
                OfficeFileCommit.WriteAllBytes(output, removal.ToArray());
                return new OfficeProvenanceTransformationResult(
                    options.Policy,
                    changed,
                    sourceReport,
                    candidateReport,
                    removal.After,
                    removal: removal);

            case OfficeProvenanceTransformationPolicy.SignAsDerived:
                if (signer == null) throw new InvalidOperationException("Derived signing requires an IOfficeProvenanceSigner.");
                if (options.Claim == null) throw new InvalidOperationException("Derived signing requires an OfficeProvenanceClaim.");
                return SignDerived(
                    source,
                    candidate,
                    output,
                    sourceBytes,
                    candidateBytes,
                    changed,
                    sourceReport,
                    candidateReport,
                    options,
                    signer);

            default:
                throw new ArgumentOutOfRangeException(nameof(options), "Unsupported provenance transformation policy.");
        }
    }

    private static string RequireFile(string path, string parameterName) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A file path is required.", parameterName);
        string fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("The provenance lifecycle input was not found.", fullPath);
        return fullPath;
    }

    private static byte[] ReadBounded(string path, long maximumBytes) {
        using var stream = File.OpenRead(path);
        return OfficeProvenanceBinary.ReadBounded(stream, maximumBytes);
    }

    private static bool BytesEqual(byte[] left, byte[] right) {
        if (left.Length != right.Length) return false;
        for (int index = 0; index < left.Length; index++) if (left[index] != right[index]) return false;
        return true;
    }

    private static bool PathsEqual(string left, string right) => string.Equals(
        Path.GetFullPath(left).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
        Path.GetFullPath(right).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
        Environment.OSVersion.Platform == PlatformID.Win32NT ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);

    private static bool HasContentCredentialCarrier(OfficeProvenanceReport report) =>
        report.HasC2paManifest || report.HasExternalC2paManifest;

    private static bool HasStructurallyValidContentCredential(OfficeProvenanceReport report) {
        foreach (OfficeProvenanceEvidence evidence in report.Evidence) {
            if (evidence.IsStructurallyValid &&
                (evidence.Carrier == OfficeProvenanceCarrierKind.C2paManifest ||
                 evidence.Carrier == OfficeProvenanceCarrierKind.C2paExternalManifest)) {
                return true;
            }
        }
        return false;
    }

    private static OfficeProvenanceTransformationResult SignDerived(
        string source,
        string candidate,
        string output,
        byte[] sourceBytes,
        byte[] candidateBytes,
        bool changed,
        OfficeProvenanceReport sourceReport,
        OfficeProvenanceReport candidateReport,
        OfficeProvenanceTransformationOptions options,
        IOfficeProvenanceSigner signer) {
        if (!options.Signing.ReplaceExistingOutput && File.Exists(output)) {
            throw new InvalidOperationException("The signing destination already exists and replacement is disabled.");
        }
        string snapshotRoot = Path.Combine(Path.GetTempPath(), "OfficeIMO-Provenance-" + Guid.NewGuid().ToString("N"));
        string sourceDirectory = Path.Combine(snapshotRoot, "source");
        string candidateDirectory = Path.Combine(snapshotRoot, "candidate");
        string sourceSnapshot = Path.Combine(sourceDirectory, Path.GetFileName(source));
        string candidateSnapshot = Path.Combine(candidateDirectory, Path.GetFileName(candidate));
        string providerOutputPath = string.Empty;
        string commitStagingPath = string.Empty;
        try {
            Directory.CreateDirectory(sourceDirectory);
            Directory.CreateDirectory(candidateDirectory);
            OfficeFileCommit.WriteAllBytes(sourceSnapshot, sourceBytes);
            OfficeFileCommit.WriteAllBytes(candidateSnapshot, candidateBytes);
            OfficeFileCommit.EnsureTargetDirectory(output);
            providerOutputPath = OfficeFileCommit.CreateStagingPath(output);
            string providerName = signer.Name;
            if (string.IsNullOrWhiteSpace(providerName)) {
                throw new InvalidOperationException("The provenance signer has no provider identity.");
            }

            OfficeProvenanceSigningResult providerResult = signer.Sign(
                new OfficeProvenanceSigningRequest(candidateSnapshot, providerOutputPath, options.Claim!, sourceSnapshot),
                options.Signing);
            if (!providerResult.Succeeded) {
                string detail = providerResult.Findings.Count == 0
                    ? providerResult.Status.ToString()
                    : string.Join("; ", providerResult.Findings);
                throw new InvalidOperationException($"The provenance signer did not produce a credentialed output: {detail}");
            }
            if (!string.Equals(providerResult.ProviderName, providerName, StringComparison.Ordinal)) {
                throw new InvalidOperationException("The provenance signer returned evidence for a different provider identity.");
            }
            if (string.IsNullOrWhiteSpace(providerResult.OutputPath) || !PathsEqual(providerResult.OutputPath!, providerOutputPath)) {
                throw new InvalidOperationException("The provenance signer did not identify the requested staging output.");
            }
            if (!File.Exists(providerOutputPath)) {
                throw new InvalidOperationException("The provenance signer reported success without creating the requested staging output.");
            }

            byte[] signedBytes = ReadBounded(providerOutputPath, options.Inspection.MaxAssetBytes);
            OfficeProvenanceReport stagedReport = OfficeProvenanceInspector.Inspect(
                signedBytes,
                providerOutputPath,
                options.Inspection);
            if (!HasStructurallyValidContentCredential(stagedReport)) {
                throw new InvalidOperationException("The provenance signer output does not contain a structurally valid Content Credential.");
            }

            commitStagingPath = OfficeFileCommit.CreateStagingPath(output);
            OfficeFileCommit.WriteAllBytes(commitStagingPath, signedBytes);
            OfficeFileCommit.CommitTemporaryFileAtomically(
                commitStagingPath,
                output,
                options.Signing.ReplaceExistingOutput
                    ? OfficeFileCommit.ConflictPolicy.Replace
                    : OfficeFileCommit.ConflictPolicy.FailIfExists);
            commitStagingPath = string.Empty;
            OfficeProvenanceReport outputReport = OfficeProvenanceInspector.InspectFile(output, options.Inspection);
            if (!HasStructurallyValidContentCredential(outputReport)) {
                throw new InvalidOperationException("The committed output no longer contains the validated Content Credential.");
            }
            var signing = new OfficeProvenanceSigningResult(
                OfficeProvenanceSigningStatus.Signed,
                providerName,
                providerResult.Findings,
                output,
                outputReport,
                providerResult.RawReport);
            return new OfficeProvenanceTransformationResult(
                options.Policy,
                changed,
                sourceReport,
                candidateReport,
                outputReport,
                signing: signing);
        } finally {
            OfficeFileCommit.DeleteIfExists(providerOutputPath);
            OfficeFileCommit.DeleteIfExists(commitStagingPath);
            OfficeFileCommit.DeleteIfExists(sourceSnapshot);
            OfficeFileCommit.DeleteIfExists(candidateSnapshot);
            TryDeleteDirectory(sourceDirectory);
            TryDeleteDirectory(candidateDirectory);
            TryDeleteDirectory(snapshotRoot);
        }
    }

    private static void TryDeleteDirectory(string path) {
        try {
            if (Directory.Exists(path)) Directory.Delete(path, recursive: false);
        } catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
