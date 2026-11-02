using System.Runtime.InteropServices;
using System.Text.Json;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Native;

namespace OfficeIMO.Studio.Core.Tests;

/// <summary>Crash recovery follows the original document across disposable native working copies.</summary>
public sealed class StudioNativeRecoveryTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-native-recovery-" + Guid.NewGuid().ToString("N"));
    private const string Identity = "studio-document://recovery/original";

    public StudioNativeRecoveryTests() {
        Directory.CreateDirectory(_root);
        StudioNativeEngine.Initialize(Path.Combine(_root, "data"));
    }
    public void Dispose() => Directory.Delete(_root, true);

    private byte[] Source() => PdfDocument.Create(document => document.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(400)))).ToBytes();

    private async Task<string> AfterCrashAsync(byte[] source) {
        string first = Path.Combine(_root, "first-copy.pdf"), next = Path.Combine(_root, "second-copy.pdf");
        await File.WriteAllBytesAsync(first, source);
        using (PdfWorkspace workspace = await PdfWorkspace.OpenAsync(first, default, StudioNativeEngine.CreateRecoveryStore(), recoveryIdentity: Identity)) {
            await workspace.RotateAsync([1], 90, default);
        }
        File.Delete(first);
        await File.WriteAllBytesAsync(next, source);
        return next;
    }

    [Fact]
    public async Task Recovery_survives_a_new_working_copy_and_remains_undoable() {
        byte[] original = Source();
        string next = await AfterCrashAsync(original);
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(next, null, default, Identity);
        Assert.True((await document.GetInfoAsync(default)).HasRecovery);
        await document.RestoreRecoveryAsync(default);
        StudioNativeDocumentInfo restored = await document.GetInfoAsync(default);
        Assert.True(restored.IsDirty);
        Assert.True(restored.CanUndo);
        Assert.False(restored.HasRecovery);
        Assert.Equal(90, Assert.Single(restored.Pages).Rotation);
        Assert.Equal(original, await File.ReadAllBytesAsync(next));
        await document.UndoAsync(default);
        Assert.Equal(0, Assert.Single((await document.GetInfoAsync(default)).Pages).Rotation);
        await document.RedoAsync(default);
        Assert.Equal(90, Assert.Single((await document.GetInfoAsync(default)).Pages).Rotation);
    }

    [Fact]
    public async Task A_different_original_revision_cannot_receive_the_old_recovery_snapshot() {
        string next = await AfterCrashAsync(Source());
        PdfDocument.Create(document => document.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(500), global::OfficeIMO.OfficeLength.Points(600)))).Save(next);
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(next, null, default, Identity);
        Assert.False((await document.GetInfoAsync(default)).HasRecovery);
        await document.RestoreRecoveryAsync(default);
        Assert.Equal(500D, Assert.Single((await document.GetInfoAsync(default)).Pages).Width);
        Assert.False((await document.GetInfoAsync(default)).IsDirty);
    }

    [Fact]
    public async Task Discard_and_identity_changes_cleanup_only_their_own_snapshots() {
        byte[] source = Source();
        string next = await AfterCrashAsync(source);
        var store = StudioNativeEngine.CreateRecoveryStore();
        using (StudioNativeDocument document = await StudioNativeDocument.OpenAsync(next, null, default, Identity)) {
            await document.DiscardRecoveryAsync(default);
            Assert.False(store.HasSnapshotFiles(Identity));
            await document.RotatePagesAsync([1], 90, default);
            const string moved = "studio-document://recovery/moved";
            await document.SetRecoveryIdentityAsync(moved, default);
            Assert.False(store.HasSnapshotFiles(Identity));
            Assert.True(store.HasSnapshotFiles(moved));
            await document.SaveAsync(null, default);
            Assert.False(store.HasSnapshotFiles(moved));
        }
    }

    [Fact]
    public async Task The_unmanaged_recovery_entry_points_restore_the_original_documents_snapshot() {
        string next = await AfterCrashAsync(Source());
        AssertUnmanagedRecovery(next);
    }

    private static void AssertUnmanagedRecovery(string next) {
        unsafe {
            nint path = Marshal.StringToCoTaskMemUTF8(next), identity = Marshal.StringToCoTaskMemUTF8(Identity);
            nint document = 0;
            byte* error = null;
            try {
                delegate* unmanaged<byte*, byte*, byte*, nint, nint*, byte**, int> open = &StudioNativeExports.DocumentOpenWithRecoveryIdentity;
                Assert.Equal(0, open((byte*)path, null, (byte*)identity, 0, &document, &error));
                delegate* unmanaged<nint, nint, byte**, int> restore = &StudioNativeExports.DocumentRestoreRecovery;
                Assert.Equal(0, restore(document, 0, &error));
                delegate* unmanaged<nint, byte**, byte**, int> info = &StudioNativeExports.DocumentInfoJson;
                byte* json = null;
                Assert.Equal(0, info(document, &json, &error));
                try {
                    using JsonDocument facts = JsonDocument.Parse(StudioNativeAbi.FromUtf8(json)!);
                    Assert.True(facts.RootElement.GetProperty("isDirty").GetBoolean());
                    Assert.Equal(90, facts.RootElement.GetProperty("pages")[0].GetProperty("rotation").GetInt32());
                } finally { NativeMemory.Free(json); }
            } finally {
                if (document != 0) { delegate* unmanaged<nint, void> close = &StudioNativeExports.DocumentClose; close(document); }
                NativeMemory.Free(error);
                Marshal.FreeCoTaskMem(path); Marshal.FreeCoTaskMem(identity);
            }
        }
    }
}
