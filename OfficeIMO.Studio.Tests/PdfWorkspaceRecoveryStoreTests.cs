using System.Text.Json.Nodes;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfWorkspaceRecoveryStoreTests {
    [Theory]
    [InlineData("expired")]
    [InlineData("future")]
    [InlineData("schema")]
    [InlineData("negative-revision")]
    [InlineData("oversize-metadata")]
    [InlineData("truncated-pdf")]
    [InlineData("oversize-pdf")]
    public async Task InvalidSnapshotsAreRejectedWithoutChangingTheSource(string corruption) {
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
            string metadataPath = Path.ChangeExtension(snapshot, ".json");
            JsonObject metadata = JsonNode.Parse(await File.ReadAllTextAsync(metadataPath))!.AsObject();
            switch (corruption) {
                case "expired": metadata["UpdatedAt"] = DateTimeOffset.UtcNow.AddDays(-31); break;
                case "future": metadata["UpdatedAt"] = DateTimeOffset.UtcNow.AddDays(2); break;
                case "schema": metadata["SchemaVersion"] = 2; break;
                case "negative-revision": metadata["Revision"] = -1; break;
                case "oversize-metadata": metadata["Padding"] = new string('x', 65 * 1024); break;
                case "truncated-pdf": await File.WriteAllBytesAsync(snapshot, original[..^1]); break;
                case "oversize-pdf":
                    using (var stream = new FileStream(snapshot, FileMode.Open, FileAccess.Write)) {
                        stream.SetLength(PdfWorkspaceRecoveryStore.MaximumSnapshotBytes + 1);
                    }
                    break;
            }
            await File.WriteAllTextAsync(metadataPath, metadata.ToJsonString());
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
}
