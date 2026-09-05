using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed partial class PdfWorkspaceRecoveryStoreTests {
    [Fact]
    public async Task ExplicitDeletionWaitsForAnotherDocumentsWriteLease() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-recovery-delete-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            var store = new PdfWorkspaceRecoveryStore(root);
            string source = Path.Combine(root, "source.pdf");
            byte[] bytes = [1, 2, 3];
            string fingerprint = PdfWorkspaceRecoveryStore.Fingerprint(bytes);
            await store.WriteAsync(source, fingerprint, bytes, 1, CancellationToken.None);
            Task deletion;
            using (var lease = new FileStream(Path.Combine(root, PdfWorkspaceRecoveryStore.LockFileName),
                FileMode.Open, FileAccess.ReadWrite, FileShare.None)) {
                deletion = store.DeleteAsync(source);
                Assert.False(deletion.IsCompleted);
                Assert.Equal(bytes, store.ReadVerifiedSnapshot(source, fingerprint));
            }
            await deletion;
            Assert.Null(store.Find(source, fingerprint));
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task IncompleteMetadataUsesFileAgeRatherThanTreatingMissingDatesAsExpired(bool legacy) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-recovery-malformed-age-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            var store = new PdfWorkspaceRecoveryStore(root);
            byte[] bytes = [1, 2, 3];
            string snapshot = await store.WriteAsync(Path.Combine(root, "source.pdf"), PdfWorkspaceRecoveryStore.Fingerprint(bytes), bytes, 1, CancellationToken.None);
            if (legacy) {
                File.Delete(snapshot);
                snapshot = Path.ChangeExtension(snapshot, ".pdf");
            }
            var metadata = new System.Text.Json.Nodes.JsonObject { ["SchemaVersion"] = legacy ? 1 : 2 };
            WriteFixture(snapshot, metadata, bytes, legacy);
            RecoveryCleanupResult recent = await store.CleanupExpiredAsync();
            Assert.Equal(0, recent.RemovedFiles);
            Assert.True(File.Exists(snapshot));
            File.SetLastWriteTimeUtc(snapshot, DateTime.UtcNow.AddDays(-31));
            if (legacy) File.SetLastWriteTimeUtc(Path.ChangeExtension(snapshot, ".json"), DateTime.UtcNow.AddDays(-31));
            RecoveryCleanupResult old = await store.CleanupExpiredAsync();
            Assert.Equal(legacy ? 2 : 1, old.RemovedFiles);
            Assert.False(File.Exists(snapshot));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ExpiryRemovesKnownOldDataAndAbandonedStagingButPreservesCurrentAndUnknownRecords() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-recovery-expiry-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            var store = new PdfWorkspaceRecoveryStore(root);
            byte[] bytes = [1, 2, 3];
            string fingerprint = PdfWorkspaceRecoveryStore.Fingerprint(bytes);
            string current = await store.WriteAsync(Path.Combine(root, "current.pdf"), fingerprint, bytes, 1, CancellationToken.None);
            string expired = await store.WriteAsync(Path.Combine(root, "expired.pdf"), fingerprint, bytes, 1, CancellationToken.None);
            var metadata = ReadMetadata(expired);
            metadata["UpdatedAt"] = DateTimeOffset.UtcNow.AddDays(-31);
            WriteFixture(expired, metadata, bytes);
            string legacy = Path.ChangeExtension(expired, ".pdf");
            metadata["SchemaVersion"] = 1;
            WriteFixture(legacy, metadata, bytes, legacy: true);
            string unknown = await store.WriteAsync(Path.Combine(root, "future.pdf"), fingerprint, bytes, 1, CancellationToken.None);
            metadata = ReadMetadata(unknown);
            metadata["SchemaVersion"] = 999;
            metadata["UpdatedAt"] = DateTimeOffset.UtcNow.AddDays(-40);
            WriteFixture(unknown, metadata, bytes);
            string staging = Path.Combine(root, ".officeimo-" + Guid.NewGuid().ToString("N") + ".tmp");
            await File.WriteAllBytesAsync(staging, bytes);
            File.SetLastWriteTimeUtc(staging, DateTime.UtcNow.AddDays(-2));
            string recentStaging = Path.Combine(root, ".officeimo-" + Guid.NewGuid().ToString("N") + ".tmp");
            await File.WriteAllBytesAsync(recentStaging, bytes);
            string unrelated = Path.Combine(root, "personal.pdf");
            await File.WriteAllBytesAsync(unrelated, bytes);
            File.SetLastWriteTimeUtc(unrelated, DateTime.UtcNow.AddDays(-40));

            RecoveryCleanupResult result = await store.CleanupExpiredAsync();
            Assert.Equal(4, result.RemovedFiles); // Current-format snapshot, legacy pair, abandoned staging.
            Assert.Equal(0, result.FailedFiles);
            Assert.True(result.RemovedBytes > 0);
            Assert.False(File.Exists(expired));
            Assert.False(File.Exists(legacy));
            Assert.False(File.Exists(Path.ChangeExtension(legacy, ".json")));
            Assert.False(File.Exists(staging));
            Assert.True(File.Exists(current));
            Assert.True(File.Exists(unknown));
            Assert.True(File.Exists(recentStaging));
            Assert.Equal(bytes, await File.ReadAllBytesAsync(unrelated));

            result = await store.ClearAllAsync();
            Assert.Equal(3, result.RemovedFiles);
            Assert.Equal(0, result.FailedFiles);
            Assert.Equal([unrelated], RecoveryDataFiles(root));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task CanceledMaintenanceAndWritesWaitForTheSameStorageLeaseWithoutChangingData() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-recovery-lease-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            var first = new PdfWorkspaceRecoveryStore(root);
            var second = new PdfWorkspaceRecoveryStore(root);
            string source = Path.Combine(root, "source.pdf");
            byte[] original = [1, 2, 3];
            string fingerprint = PdfWorkspaceRecoveryStore.Fingerprint(original);
            await first.WriteAsync(source, fingerprint, original, 1, CancellationToken.None);
            using (var lease = new FileStream(Path.Combine(root, PdfWorkspaceRecoveryStore.LockFileName),
                FileMode.Open, FileAccess.ReadWrite, FileShare.None)) {
                using var cancellation = new CancellationTokenSource();
                Task write = second.WriteAsync(source, fingerprint, [4, 5, 6], 2, cancellation.Token);
                Task clear = first.ClearAllAsync(cancellation.Token);
                Assert.False(write.IsCompleted);
                Assert.False(clear.IsCompleted);
                cancellation.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => write);
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => clear);
                Assert.Equal(original, first.ReadVerifiedSnapshot(source, fingerprint));
            }
            RecoveryCleanupResult result = await second.ClearAllAsync();
            Assert.Equal(1, result.RemovedFiles);
            Assert.Equal(0, result.FailedFiles);
            Assert.Null(first.Find(source, fingerprint));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ConcurrentExpiryAndPublicationRetainACompleteCurrentSnapshot() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-recovery-concurrent-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf");
            string fingerprint = PdfWorkspaceRecoveryStore.Fingerprint([0]);
            var first = new PdfWorkspaceRecoveryStore(root);
            var second = new PdfWorkspaceRecoveryStore(root);
            var operations = new List<Task>();
            for (int index = 1; index <= 20; index++) {
                operations.Add(first.WriteAsync(source, fingerprint, [(byte)index], index, CancellationToken.None));
                operations.Add(second.CleanupExpiredAsync());
            }
            await Task.WhenAll(operations);
            byte[] saved = Assert.IsType<byte[]>(first.ReadVerifiedSnapshot(source, fingerprint));
            Assert.Single(saved);
            Assert.InRange(saved[0], (byte)1, (byte)20);
            Assert.Single(RecoveryDataFiles(root));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task CleanupReportsLockedFilesAndPreservesLinksAndTheirTargets() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-recovery-clear-failure-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string recoveryRoot = Path.Combine(root, "recovery");
            var store = new PdfWorkspaceRecoveryStore(recoveryRoot);
            byte[] bytes = [1, 2, 3];
            string snapshot = await store.WriteAsync(Path.Combine(root, "source.pdf"), PdfWorkspaceRecoveryStore.Fingerprint(bytes), bytes, 1, CancellationToken.None);
            if (OperatingSystem.IsWindows()) {
                using var locked = new FileStream(snapshot, FileMode.Open, FileAccess.Read, FileShare.Read);
                RecoveryCleanupResult failed = await store.ClearAllAsync();
                Assert.Equal(1, failed.FailedFiles);
                Assert.Equal(0, failed.RemovedFiles);
                Assert.True(File.Exists(snapshot));
            } else {
                string target = Path.Combine(root, "private-source.pdf");
                await File.WriteAllBytesAsync(target, bytes);
                string link = Path.Combine(recoveryRoot, new string('a', 32) + ".recovery");
                File.CreateSymbolicLink(link, target);
                RecoveryCleanupResult result = await store.ClearAllAsync();
                Assert.Equal(1, result.FailedFiles);
                Assert.True(File.Exists(link));
                Assert.Equal(bytes, await File.ReadAllBytesAsync(target));
            }
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }
}
