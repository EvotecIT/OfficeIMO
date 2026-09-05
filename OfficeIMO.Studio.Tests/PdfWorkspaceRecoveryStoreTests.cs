using System.Text.Json.Nodes;
using System.Buffers.Binary;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed partial class PdfWorkspaceRecoveryStoreTests {
    [Theory]
    [InlineData("expired", false)]
    [InlineData("future", false)]
    [InlineData("schema", false)]
    [InlineData("negative-revision", false)]
    [InlineData("oversize-metadata", false)]
    [InlineData("truncated-pdf", false)]
    [InlineData("oversize-pdf", false)]
    [InlineData("expired", true)]
    [InlineData("future", true)]
    [InlineData("schema", true)]
    [InlineData("negative-revision", true)]
    [InlineData("oversize-metadata", true)]
    [InlineData("truncated-pdf", true)]
    [InlineData("oversize-pdf", true)]
    public async Task InvalidSnapshotsAreRejectedWithoutChangingTheSource(string corruption, bool legacy) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-recovery-validation-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf");
            PdfDocument.Create(compose => compose.Page(page => page.Size(600D, 800D))).Save(source);
            byte[] original = await File.ReadAllBytesAsync(source);
            string fingerprint = PdfWorkspaceRecoveryStore.Fingerprint(original);
            var store = new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery"));
            string snapshot = await store.WriteAsync(source, fingerprint, original, 1, CancellationToken.None);
            Assert.Equal(original, store.ReadVerifiedSnapshot(source, fingerprint));
            JsonObject metadata = ReadMetadata(snapshot);
            if (legacy) {
                File.Delete(snapshot);
                snapshot = Path.ChangeExtension(snapshot, ".pdf");
                metadata["SchemaVersion"] = 1;
            }
            switch (corruption) {
                case "expired": metadata["UpdatedAt"] = DateTimeOffset.UtcNow.AddDays(-31); break;
                case "future": metadata["UpdatedAt"] = DateTimeOffset.UtcNow.AddDays(2); break;
                case "schema": metadata["SchemaVersion"] = 99; break;
                case "negative-revision": metadata["Revision"] = -1; break;
                case "oversize-metadata": metadata["Padding"] = new string('x', 65 * 1024); break;
            }
            WriteFixture(snapshot, metadata, original, legacy);
            switch (corruption) {
                case "truncated-pdf":
                    using (var stream = new FileStream(snapshot, FileMode.Open, FileAccess.Write)) {
                        stream.SetLength(stream.Length - 1);
                    }
                    break;
                case "oversize-pdf":
                    using (var stream = new FileStream(snapshot, FileMode.Open, FileAccess.Write)) {
                        stream.SetLength(PdfWorkspaceRecoveryStore.MaximumSnapshotBytes + 128 * 1024);
                    }
                    break;
            }
            Assert.Null(store.ReadVerifiedSnapshot(source, fingerprint));
            Assert.Null(store.Find(source, fingerprint));
            Assert.Equal(original, await File.ReadAllBytesAsync(source));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task RestoreRevalidatesSnapshotAfterTheDocumentWasOpened() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-recovery-restore-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf");
            string replacement = Path.Combine(root, "replacement.pdf");
            PdfDocument.Create(compose => compose.Page(page => page.Size(600D, 800D))).Save(source);
            PdfDocument.Create(compose => {
                compose.Page(page => page.Size(600D, 800D));
                compose.Page(page => page.Size(600D, 800D));
            }).Save(replacement);
            byte[] original = await File.ReadAllBytesAsync(source);
            byte[] edited = await File.ReadAllBytesAsync(replacement);
            string fingerprint = PdfWorkspaceRecoveryStore.Fingerprint(original);
            var store = new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery"));
            string snapshot = await store.WriteAsync(source, fingerprint, edited, 1, CancellationToken.None);
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None, recoveryStore: store);
            Assert.True(workspace.HasRecovery);
            await File.WriteAllBytesAsync(snapshot, original);

            await Assert.ThrowsAsync<InvalidDataException>(() => workspace.RestoreRecoveryAsync(CancellationToken.None));
            Assert.False(workspace.IsDirty);
            Assert.Equal(original, workspace.CopyBytes());
            Assert.Empty(workspace.Journal);

            await store.WriteAsync(source, fingerprint, edited, 1, CancellationToken.None);
            await workspace.RestoreRecoveryAsync(CancellationToken.None);
            Assert.True(workspace.IsDirty);
            Assert.Equal(edited, workspace.CopyBytes());
            Assert.Equal(original, await File.ReadAllBytesAsync(source));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task FailedPublicationAndCancellationPreserveTheLastCompleteSnapshot() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-recovery-atomic-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf");
            byte[] first = [1, 2, 3], second = [4, 5, 6];
            string fingerprint = PdfWorkspaceRecoveryStore.Fingerprint(first);
            var store = new PdfWorkspaceRecoveryStore(root);
            string snapshot = await store.WriteAsync(source, fingerprint, first, 1, CancellationToken.None);
            if (!OperatingSystem.IsWindows()) {
                Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite, File.GetUnixFileMode(snapshot));
                File.SetUnixFileMode(snapshot, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.GroupRead);
            }
            using var canceled = new CancellationTokenSource();
            canceled.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                store.WriteAsync(source, fingerprint, second, 2, canceled.Token));
            Assert.Equal(first, store.ReadVerifiedSnapshot(source, fingerprint));
            Assert.Equal([snapshot], RecoveryDataFiles(root));

            if (OperatingSystem.IsWindows()) {
                using (var locked = new FileStream(snapshot, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                    Exception? error = await Record.ExceptionAsync(() =>
                        store.WriteAsync(source, fingerprint, second, 2, CancellationToken.None));
                    Assert.True(error is IOException or UnauthorizedAccessException);
                }
                Assert.Equal(first, store.ReadVerifiedSnapshot(source, fingerprint));
                Assert.Equal([snapshot], RecoveryDataFiles(root));
            }

            // An interrupted, unpublished staging file cannot hide the committed snapshot.
            string interrupted = snapshot + ".tmp-" + Guid.NewGuid().ToString("N");
            await File.WriteAllBytesAsync(interrupted, [0]);
            Assert.Equal(first, store.ReadVerifiedSnapshot(source, fingerprint));
            await store.WriteAsync(source, fingerprint, second, 2, CancellationToken.None);
            Assert.Equal(second, store.ReadVerifiedSnapshot(source, fingerprint));
            if (!OperatingSystem.IsWindows()) Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite, File.GetUnixFileMode(snapshot));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task LegacySnapshotsRemainReadableAndMigrateOnlyAfterSuccessfulPublication() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-recovery-migration-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf");
            byte[] first = [1, 2, 3], second = [4, 5, 6];
            string fingerprint = PdfWorkspaceRecoveryStore.Fingerprint(first);
            var store = new PdfWorkspaceRecoveryStore(root);
            string snapshot = await store.WriteAsync(source, fingerprint, first, 1, CancellationToken.None);
            JsonObject metadata = ReadMetadata(snapshot);
            metadata.Remove("SchemaVersion"); // Original records predate the schema field.
            File.Delete(snapshot);
            string legacy = Path.ChangeExtension(snapshot, ".pdf");
            WriteFixture(legacy, metadata, first, legacy: true);
            Assert.Equal(first, store.ReadVerifiedSnapshot(source, fingerprint));

            using var canceled = new CancellationTokenSource();
            canceled.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                store.WriteAsync(source, fingerprint, second, 2, canceled.Token));
            Assert.Equal(first, store.ReadVerifiedSnapshot(source, fingerprint));

            await store.WriteAsync(source, fingerprint, second, 2, CancellationToken.None);
            Assert.Equal(second, store.ReadVerifiedSnapshot(source, fingerprint));
            Assert.False(File.Exists(legacy));
            Assert.False(File.Exists(Path.ChangeExtension(legacy, ".json")));

            // A leftover legacy pair must not replace a corrupt newer generation silently.
            WriteFixture(legacy, metadata, first, legacy: true);
            await File.WriteAllBytesAsync(snapshot, [0]);
            Assert.Null(store.ReadVerifiedSnapshot(source, fingerprint));
            store.Delete(source);
            Assert.Null(store.Find(source, fingerprint));
            Assert.Empty(RecoveryDataFiles(root));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    private static JsonObject ReadMetadata(string snapshot) {
        using var stream = File.OpenRead(snapshot);
        byte[] header = new byte[20];
        stream.ReadExactly(header);
        byte[] metadata = new byte[BinaryPrimitives.ReadInt32LittleEndian(header.AsSpan(8, 4))];
        stream.ReadExactly(metadata);
        return JsonNode.Parse(metadata)!.AsObject();
    }

    private static string[] RecoveryDataFiles(string root) => Directory.GetFiles(root)
        .Where(path => Path.GetFileName(path) != PdfWorkspaceRecoveryStore.LockFileName).ToArray();

    private static void WriteFixture(string path, JsonObject metadata, byte[] pdf, bool legacy = false) {
        byte[] json = System.Text.Encoding.UTF8.GetBytes(metadata.ToJsonString());
        if (legacy) {
            File.WriteAllBytes(path, pdf);
            File.WriteAllBytes(Path.ChangeExtension(path, ".json"), json);
            return;
        }
        using var stream = File.Create(path);
        byte[] header = new byte[20];
        "OIMORCV2"u8.CopyTo(header);
        BinaryPrimitives.WriteInt32LittleEndian(header.AsSpan(8, 4), json.Length);
        BinaryPrimitives.WriteInt64LittleEndian(header.AsSpan(12, 8), pdf.LongLength);
        stream.Write(header);
        stream.Write(json);
        stream.Write(pdf);
    }
}
