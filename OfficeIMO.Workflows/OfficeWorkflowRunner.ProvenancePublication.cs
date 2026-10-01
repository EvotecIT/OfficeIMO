using System.Security.Cryptography;
using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    internal sealed class StagedArtifactFingerprint : IDisposable {
        private readonly string? _physicalIdentity;
        private readonly bool _usesPhysicalIdentity;
        private FileStream? _lease;
        private FileStream? _publishedLease;

        private StagedArtifactFingerprint(
            long length,
            byte[] sha256,
            string? physicalIdentity,
            bool usesPhysicalIdentity,
            FileStream lease) {
            Length = length;
            Sha256 = sha256;
            _physicalIdentity = physicalIdentity;
            _usesPhysicalIdentity = usesPhysicalIdentity;
            _lease = lease;
        }

        internal long Length { get; }
        private byte[] Sha256 { get; }
        internal string Sha256Hex => System.Convert.ToHexString(Sha256).ToLowerInvariant();

        internal static StagedArtifactFingerprint Capture(
            string path,
            long maximumBytes,
            CancellationToken cancellationToken,
            string artifactDescription = "staged provenance artifact") => CaptureCore(
                path,
                maximumBytes,
                expectedLength: null,
                expectedSha256: null,
                cancellationToken,
                artifactDescription,
                OfficeWorkflowPathIdentity.SupportsPhysicalIdentity);

        /// <summary>Exercises the portable length-and-hash fingerprint path when filesystem identity is unavailable.</summary>
        internal static StagedArtifactFingerprint CapturePortable(
            string path,
            long maximumBytes,
            CancellationToken cancellationToken = default) => CaptureCore(
                path,
                maximumBytes,
                expectedLength: null,
                expectedSha256: null,
                cancellationToken,
                "staged provenance artifact",
                usesPhysicalIdentity: false);

        internal static StagedArtifactFingerprint CaptureExpected(
            string path,
            long maximumBytes,
            long expectedLength,
            byte[] expectedSha256,
            CancellationToken cancellationToken) => CaptureCore(
                path,
                maximumBytes,
                expectedLength,
                expectedSha256 ?? throw new ArgumentNullException(nameof(expectedSha256)),
                cancellationToken,
                "staged provenance artifact",
                OfficeWorkflowPathIdentity.SupportsPhysicalIdentity);

        private static StagedArtifactFingerprint CaptureCore(
            string path,
            long maximumBytes,
            long? expectedLength,
            byte[]? expectedSha256,
            CancellationToken cancellationToken,
            string artifactDescription,
            bool usesPhysicalIdentity) {
            var stream = OpenForIdentity(path);
            try {
                if (stream.Length > maximumBytes) {
                    throw OfficeProvenanceLimitException.CreateOutput(
                        $"The {artifactDescription} exceeds the configured output limit of {maximumBytes} bytes.");
                }
                byte[] sha256 = ComputeHash(stream, cancellationToken);
                if (expectedLength.HasValue &&
                    (stream.Length != expectedLength.Value ||
                     !CryptographicOperations.FixedTimeEquals(sha256, expectedSha256!))) {
                    throw new InvalidDataException(
                        "The staged provenance artifact did not match the bytes returned by its format owner.");
                }
                stream.Position = 0;
                string? physicalIdentity = usesPhysicalIdentity
                    ? OfficeWorkflowPathIdentity.GetPhysicalIdentityKey(path, stream)
                    : null;
                return new StagedArtifactFingerprint(stream.Length, sha256, physicalIdentity, usesPhysicalIdentity, stream);
            } catch {
                stream.Dispose();
                throw;
            }
        }

        internal void VerifyStagingPath(string path, long maximumBytes, CancellationToken cancellationToken) {
            if (!MatchesPath(path, maximumBytes, cancellationToken)) {
                throw new InvalidDataException(
                    "The staged provenance artifact changed after output validation; publication was blocked.");
            }
        }

        internal bool TryPinPublishedPath(string path, long maximumBytes, CancellationToken cancellationToken) {
            FileStream stream = OpenForIdentity(path);
            try {
                if (!MatchesStream(path, stream, maximumBytes, cancellationToken)) return false;
                _publishedLease?.Dispose();
                _publishedLease = stream;
                return true;
            } catch {
                stream.Dispose();
                throw;
            } finally {
                if (!ReferenceEquals(_publishedLease, stream)) stream.Dispose();
            }
        }

        internal void VerifyPublishedPath(string path, long maximumBytes, CancellationToken cancellationToken) {
            if (_publishedLease is null || !MatchesPath(path, maximumBytes, cancellationToken)) {
                throw new InvalidDataException(
                    "The published provenance artifact changed before publication was finalized.");
            }
        }

        internal void ReleasePublishedLease() {
            _publishedLease?.Dispose();
            _publishedLease = null;
        }

        internal void ReleaseStagingLease() {
            _lease?.Dispose();
            _lease = null;
        }

        internal void TryDeleteMatchingPath(string path, long maximumBytes, CancellationToken cancellationToken) {
            string quarantinePath = Path.Combine(
                Path.GetDirectoryName(Path.GetFullPath(path))!,
                ".officeimo-provenance-rollback-" + Guid.NewGuid().ToString("N") + ".tmp");
            bool moved = false;
            try {
                File.Move(path, quarantinePath, overwrite: false);
                moved = true;
                bool matches;
                using (FileStream stream = OpenForIdentity(quarantinePath)) {
                    matches = MatchesStream(quarantinePath, stream, maximumBytes, cancellationToken);
                }
                if (matches) {
                    File.Delete(quarantinePath);
                    moved = false;
                    return;
                }

                File.Move(quarantinePath, path, overwrite: false);
                moved = false;
            } catch (Exception exception) when (exception is FileNotFoundException or DirectoryNotFoundException or IOException or UnauthorizedAccessException) {
                // Never delete a known destination pathname after a failed identity check. If a
                // different writer claimed it, retain the random quarantine rather than losing data.
            } finally {
                if (moved && File.Exists(quarantinePath)) {
                    try {
                        if (!File.Exists(path)) {
                            File.Move(quarantinePath, path, overwrite: false);
                            moved = false;
                        }
                    } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }
                }
            }
        }

        public void Dispose() {
            ReleasePublishedLease();
            ReleaseStagingLease();
        }

        internal bool MatchesPath(string path, long maximumBytes, CancellationToken cancellationToken) {
            try {
                using FileStream stream = OpenForIdentity(path);
                return MatchesStream(path, stream, maximumBytes, cancellationToken);
            } catch (Exception exception) when (exception is FileNotFoundException or DirectoryNotFoundException) {
                return false;
            }
        }

        private bool MatchesStream(
            string path,
            FileStream stream,
            long maximumBytes,
            CancellationToken cancellationToken) {
            if (stream.Length > maximumBytes || stream.Length != Length) return false;
            if (_usesPhysicalIdentity) {
                string physicalIdentity = OfficeWorkflowPathIdentity.GetPhysicalIdentityKey(path, stream);
                if (!string.Equals(physicalIdentity, _physicalIdentity, StringComparison.Ordinal)) return false;
            }
            byte[] currentHash = ComputeHash(stream, cancellationToken);
            return CryptographicOperations.FixedTimeEquals(currentHash, Sha256);
        }

        private static FileStream OpenForIdentity(string path) => new(
            path,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete,
            81920,
            FileOptions.SequentialScan);

        private static byte[] ComputeHash(Stream stream, CancellationToken cancellationToken) {
            stream.Position = 0;
            using IncrementalHash algorithm = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            var buffer = new byte[81920];
            int read;
            while ((read = stream.Read(buffer, 0, buffer.Length)) != 0) {
                cancellationToken.ThrowIfCancellationRequested();
                algorithm.AppendData(buffer, 0, read);
            }
            return algorithm.GetHashAndReset();
        }
    }

    private static async Task<string> PublishVerifiedAsync(
        string stagingPath,
        string requestedPath,
        OfficeWorkflowConflictPolicy policy,
        SortedSet<string>? blockedOutputIdentities,
        string? ownReservedOutputIdentity,
        StagedArtifactFingerprint staged,
        IOfficeWorkflowPublicationGuard? guard,
        OfficeProvenanceFileSnapshot? expectedDisplacedInput,
        long maximumBytes,
        CancellationToken cancellationToken,
        Action beforePublish,
        Action<string> beforeCommitFinalized,
        Action<string, Exception> backupCleanupFailed) {
        bool published = false;
        try {
            beforePublish();
            cancellationToken.ThrowIfCancellationRequested();
            staged.VerifyStagingPath(stagingPath, maximumBytes, cancellationToken);
            staged.ReleaseStagingLease();

            string publishedPath;
            switch (policy) {
                case OfficeWorkflowConflictPolicy.Fail:
                    EnsureBatchCandidateDoesNotOverlapAnotherRequest(requestedPath);
                    await EnsurePublicationAllowedAsync(guard, requestedPath, false, cancellationToken).ConfigureAwait(false);
                    File.Move(stagingPath, requestedPath, overwrite: false);
                    publishedPath = requestedPath;
                    PinAndFinalize(publishedPath);
                    break;
                case OfficeWorkflowConflictPolicy.Rename:
                    publishedPath = await PublishRenamedAsync().ConfigureAwait(false);
                    break;
                case OfficeWorkflowConflictPolicy.Replace:
                    publishedPath = await PublishReplacementAsync().ConfigureAwait(false);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(policy), policy, "Unsupported conflict policy.");
            }

            published = true;
            return publishedPath;
        } finally {
            if (!published) {
                staged.ReleasePublishedLease();
                staged.TryDeleteMatchingPath(stagingPath, maximumBytes, CancellationToken.None);
            }
        }

        void PinAndFinalize(string path) {
            try {
                if (!staged.TryPinPublishedPath(path, maximumBytes, cancellationToken)) {
                    throw new InvalidDataException(
                        "The staged provenance artifact changed while it was being published.");
                }
                beforeCommitFinalized(path);
            } catch {
                staged.ReleasePublishedLease();
                staged.TryDeleteMatchingPath(path, maximumBytes, CancellationToken.None);
                throw;
            }
        }

        void EnsureBatchCandidateDoesNotOverlapAnotherRequest(string path) {
            if (blockedOutputIdentities is null) return;
            string identity = OfficeWorkflowPathIdentity.NormalizeWithPortableFallback(path);
            bool isOwnReservation = string.Equals(identity, ownReservedOutputIdentity, StringComparison.Ordinal);
            bool hasHierarchyCollision = TryFindAncestorOrDescendant(
                identity,
                blockedOutputIdentities,
                out string? collisionIdentity) &&
                !string.Equals(collisionIdentity, ownReservedOutputIdentity, StringComparison.Ordinal);
            if ((!isOwnReservation && blockedOutputIdentities.Contains(identity)) || hasHierarchyCollision) {
                throw new IOException(
                    "The provenance output now overlaps another batch request path and cannot be published safely.");
            }
        }

        async Task<string> PublishRenamedAsync() {
            for (int suffix = 0; suffix < 10_000; suffix++) {
                cancellationToken.ThrowIfCancellationRequested();
                string candidate = suffix == 0 ? requestedPath : AddSuffix(requestedPath, suffix);
                if (blockedOutputIdentities is not null) {
                    string identity = OfficeWorkflowPathIdentity.NormalizeWithPortableFallback(candidate);
                    bool isOwnReservation = string.Equals(
                        identity,
                        ownReservedOutputIdentity,
                        StringComparison.Ordinal);
                    bool hasHierarchyCollision = TryFindAncestorOrDescendant(
                        identity,
                        blockedOutputIdentities,
                        out string? collisionIdentity) &&
                        !string.Equals(collisionIdentity, ownReservedOutputIdentity, StringComparison.Ordinal);
                    if ((!isOwnReservation && blockedOutputIdentities.Contains(identity)) ||
                        hasHierarchyCollision) continue;
                }
                if (!await CanPublishAsync(guard, candidate, false, cancellationToken).ConfigureAwait(false)) continue;
                try {
                    File.Move(stagingPath, candidate, overwrite: false);
                } catch (IOException) when (File.Exists(candidate) || Directory.Exists(candidate)) {
                    // Another request owns this candidate. Try the next deterministic suffix.
                    continue;
                }
                PinAndFinalize(candidate);
                return candidate;
            }
            throw new IOException("No available numbered output path could be reserved.");
        }

        async Task<string> PublishReplacementAsync() {
            await EnsurePublicationAllowedAsync(guard, requestedPath, false, cancellationToken).ConfigureAwait(false);
            EnsureBatchCandidateDoesNotOverlapAnotherRequest(requestedPath);
            if (!File.Exists(requestedPath)) {
                if (expectedDisplacedInput != null) {
                    throw new IOException(
                        "The provenance input changed while its verified replacement was being published.");
                }
                bool destinationAppeared = false;
                try {
                    File.Move(stagingPath, requestedPath, overwrite: false);
                } catch (IOException) when (File.Exists(requestedPath)) {
                    // The destination appeared during the claim. Validate and replace it below.
                    destinationAppeared = true;
                }
                if (!destinationAppeared) {
                    PinAndFinalize(requestedPath);
                    return requestedPath;
                }
            }

            EnsureBatchCandidateDoesNotOverlapAnotherRequest(requestedPath);
            if (expectedDisplacedInput != null) {
                bool inputCommitted = OfficeFileCommit.TryCommitTemporaryFileAtomicallyIfDestinationUnchangedAndFinalize(
                    stagingPath,
                    requestedPath,
                    backupPath => expectedDisplacedInput.MatchesCapturedSource(backupPath, cancellationToken),
                    installedPath => staged.TryPinPublishedPath(installedPath, maximumBytes, cancellationToken),
                    beforeCommitFinalized,
                    backupCleanupFailed);
                if (!inputCommitted) {
                    staged.ReleasePublishedLease();
                    throw new IOException(
                        "The provenance input changed while its verified replacement was being published.");
                }
                return requestedPath;
            }

            using StagedArtifactFingerprint destination = StagedArtifactFingerprint.Capture(
                requestedPath,
                maximumBytes,
                cancellationToken,
                "existing provenance destination");
            destination.ReleaseStagingLease();
            bool committed = OfficeFileCommit.TryCommitTemporaryFileAtomicallyIfDestinationUnchangedAndFinalize(
                stagingPath,
                requestedPath,
                backupPath => destination.MatchesPath(backupPath, maximumBytes, cancellationToken),
                installedPath => staged.TryPinPublishedPath(installedPath, maximumBytes, cancellationToken),
                beforeCommitFinalized,
                backupCleanupFailed);
            if (!committed) {
                staged.ReleasePublishedLease();
                throw new IOException("The provenance destination changed while the verified artifact was being published.");
            }
            return requestedPath;
        }
    }

}
