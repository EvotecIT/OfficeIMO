using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed partial class PdfWorkspaceTests {
    [Fact]
    public async Task EncryptedStampsPreserveProtectionAcrossRecoveryHistoryAndSave() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-encrypted-edit-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        string protectedCopy = Path.Combine(root, "protected.pdf");
        string saved = Path.Combine(root, "saved.pdf");
        var recovery = new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery"));
        CreateTextSource(source, "Keep encrypted");
        try {
            using (PdfWorkspace plain = await PdfWorkspace.OpenAsync(source, CancellationToken.None, recovery)) {
                await plain.SaveProtectedCopyAsync(protectedCopy, new PdfStandardEncryptionOptions("open") {
                    OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.All
                }, null, CancellationToken.None);
            }
            using PdfWorkspace encrypted = await PdfWorkspace.OpenAsync(protectedCopy, CancellationToken.None, recovery, password: "open");
            byte[] original = encrypted.CopyBytes();
            Assert.True(encrypted.CanEditPageContent);
            await encrypted.ApplyPageNumbersAsync(CancellationToken.None);
            byte[] numbered = encrypted.CopyBytes();
            Assert.Contains("1 / 1", ReadProtectedWorkspaceText(numbered, PdfStandardPermissions.All), StringComparison.Ordinal);
            await encrypted.ApplyWatermarkAsync("draft", CancellationToken.None);
            byte[] edited = encrypted.CopyBytes();
            Assert.Contains("draft", ReadProtectedWorkspaceText(edited, PdfStandardPermissions.All), StringComparison.Ordinal);
            Assert.True(encrypted.IsEncrypted);
            Assert.True(encrypted.IsDirty);
            Assert.Equal([PdfWorkspaceOperationKind.PageNumbers, PdfWorkspaceOperationKind.Watermark],
                encrypted.Journal.Select(operation => operation.Kind).ToArray());
            Assert.Equal(edited, recovery.ReadVerifiedSnapshot(protectedCopy, PdfWorkspaceRecoveryStore.Fingerprint(original)));

            using (PdfWorkspace reopened = await PdfWorkspace.OpenAsync(protectedCopy, CancellationToken.None, recovery, password: "open")) {
                Assert.True(reopened.HasRecovery);
                await reopened.RestoreRecoveryAsync(CancellationToken.None);
                Assert.Equal(edited, reopened.CopyBytes());
                Assert.True(reopened.IsEncrypted);
            }
            await encrypted.UndoAsync(CancellationToken.None);
            Assert.Equal(numbered, encrypted.CopyBytes());
            Assert.True(encrypted.IsEncrypted);
            await encrypted.RedoAsync(CancellationToken.None);
            Assert.Equal(edited, encrypted.CopyBytes());
            Assert.True(encrypted.IsEncrypted);
            await encrypted.SaveAsync(saved, CancellationToken.None);

            byte[] savedBytes = await File.ReadAllBytesAsync(saved);
            string text = ReadProtectedWorkspaceText(savedBytes, PdfStandardPermissions.All);
            Assert.Contains("Keep encrypted", text, StringComparison.Ordinal);
            Assert.Contains("1 / 1", text, StringComparison.Ordinal);
            Assert.Contains("draft", text, StringComparison.Ordinal);
            Assert.Equal(original, await File.ReadAllBytesAsync(protectedCopy));
            Assert.False(encrypted.IsDirty);
            Assert.False(encrypted.HasRecovery);
            Assert.Null(recovery.ReadVerifiedSnapshot(protectedCopy, PdfWorkspaceRecoveryStore.Fingerprint(original)));
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task EncryptedStampsWithoutContentPermissionLeaveStateAndRecoveryUnchanged() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-restricted-edit-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "protected.pdf");
        var recovery = new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery"));
        const PdfStandardPermissions permissions = PdfStandardPermissions.Print | PdfStandardPermissions.CopyContents;
        try {
            CreateTextSource(source, "Restricted content");
            PdfDocument.Load(source).Security.Encrypt(new PdfStandardEncryptionOptions("open") {
                OwnerPassword = "owner", AllowedPermissions = permissions
            }).ToDocument().Save(source);
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None, recovery, password: "open");
            byte[] original = workspace.CopyBytes();

            Assert.False(workspace.CanEditPageContent);
            await Assert.ThrowsAsync<PdfMutationBlockedException>(() => workspace.ApplyPageNumbersAsync(CancellationToken.None));
            await Assert.ThrowsAsync<PdfMutationBlockedException>(() => workspace.ApplyWatermarkAsync("draft", CancellationToken.None));

            Assert.Equal(original, workspace.CopyBytes());
            Assert.Equal(original, await File.ReadAllBytesAsync(source));
            Assert.True(workspace.IsEncrypted);
            Assert.False(workspace.IsDirty);
            Assert.False(workspace.HasRecovery);
            Assert.Empty(workspace.Journal);
            Assert.Null(recovery.ReadVerifiedSnapshot(source, PdfWorkspaceRecoveryStore.Fingerprint(original)));
            Assert.Contains("Restricted content", ReadProtectedWorkspaceText(original, permissions), StringComparison.Ordinal);
        } finally { Directory.Delete(root, recursive: true); }
    }

    private static string ReadProtectedWorkspaceText(byte[] bytes, PdfStandardPermissions permissions) {
        Assert.Throws<PdfPasswordRequiredException>(() => PdfDocument.Load(bytes).Inspect());
        Assert.Throws<PdfInvalidPasswordException>(() => PdfDocument.Load(bytes, new PdfLoadOptions { Password = "wrong" }).Inspect());
        foreach (string password in new[] { "open", "owner" }) {
            var security = PdfDocument.Load(bytes, new PdfLoadOptions { Password = password }).Inspect().Security;
            Assert.True(security.HasEncryption);
            Assert.Equal(permissions, security.AllowedStandardPermissions);
        }
        return PdfDocument.Load(bytes, new PdfLoadOptions { Password = "open" }).Read().Text;
    }
}
