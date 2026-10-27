using System.Security.Cryptography;
using System.Runtime.InteropServices;
using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Native;

namespace OfficeIMO.Studio.Core.Tests;

public sealed class StudioNativeExternalSaveTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-external-save-" + Guid.NewGuid().ToString("N"));
    private const string Identity = "studio-document://external-save/original";
    public StudioNativeExternalSaveTests() {
        Directory.CreateDirectory(_root);
        StudioNativeEngine.Initialize(Path.Combine(_root, "data"));
    }
    public void Dispose() => Directory.Delete(_root, true);
    private string Copy(string name, byte[] bytes) {
        string path = Path.Combine(_root, name + ".pdf");
        File.WriteAllBytes(path, bytes);
        return path;
    }
    private static byte[] Source() => PdfDocument.Create(document => document.Page(page => page.Size(300, 400))).ToBytes();

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Failed_or_unacknowledged_publication_retains_recovery_for_the_actual_original(bool published, bool acknowledge) {
        byte[] original = Source(), prepared;
        string first = Copy("first", original);
        using (var document = await StudioNativeDocument.OpenAsync(first, null, default, Identity)) {
            await document.RotatePagesAsync([1], 90, default);
            prepared = await document.PrepareSaveAsync(default);
            Assert.Equal(original, File.ReadAllBytes(first));
            Assert.True((await document.GetInfoAsync(default)).IsDirty);
            if (acknowledge) {
                await document.AcknowledgeSaveAsync(Convert.ToHexString(SHA256.HashData(prepared)), default);
                Assert.False((await document.GetInfoAsync(default)).IsDirty);
            }
            await document.RotatePagesAsync([1], 90, default);
            Assert.Equal(90, PdfDocument.Load(new System.IO.MemoryStream(prepared)).GetPageLayouts()[0].RotationDegrees);
        }
        string reopened = Copy("next", published ? prepared : original);
        using var next = await StudioNativeDocument.OpenAsync(reopened, null, default, Identity);
        Assert.True((await next.GetInfoAsync(default)).HasRecovery);
        await next.RestoreRecoveryAsync(default);
        Assert.Equal(180, (await next.GetInfoAsync(default)).Pages[0].Rotation);
        await next.UndoAsync(default);
        Assert.Equal(published ? 90 : 0, (await next.GetInfoAsync(default)).Pages[0].Rotation);
    }

    [Fact]
    public async Task Only_an_observed_prepared_snapshot_can_advance_saved_state() {
        using var document = await StudioNativeDocument.OpenAsync(Copy("working", Source()), null, default, Identity);
        await document.RotatePagesAsync([1], 90, default);
        byte[] first = await document.PrepareSaveAsync(default);
        await document.RotatePagesAsync([1], 90, default);
        byte[] second = await document.PrepareSaveAsync(default);
        await document.AcknowledgeSaveAsync(new string('0', 64), default);
        Assert.True((await document.GetInfoAsync(default)).IsDirty);
        await document.AcknowledgeSaveAsync(Convert.ToHexString(SHA256.HashData(second)), default, Identity + "/old");
        Assert.True((await document.GetInfoAsync(default)).IsDirty);
        await document.AcknowledgeSaveAsync(Convert.ToHexString(SHA256.HashData(first)), default);
        Assert.True((await document.GetInfoAsync(default)).IsDirty);
        await document.AcknowledgeSaveAsync(Convert.ToHexString(SHA256.HashData(second)), default);
        Assert.False((await document.GetInfoAsync(default)).IsDirty);
        Assert.False(StudioNativeEngine.CreateRecoveryStore().HasSnapshotFiles(Identity));
    }

    [Fact]
    public async Task A_saved_snapshot_without_an_acknowledgement_does_not_offer_identical_recovery() {
        byte[] prepared;
        using (var document = await StudioNativeDocument.OpenAsync(Copy("first", Source()), null, default, Identity)) {
            await document.RotatePagesAsync([1], 90, default);
            prepared = await document.PrepareSaveAsync(default);
        }
        using var next = await StudioNativeDocument.OpenAsync(Copy("next", prepared), null, default, Identity);
        Assert.False((await next.GetInfoAsync(default)).HasRecovery);
        Assert.Equal(90, (await next.GetInfoAsync(default)).Pages[0].Rotation);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Undo_redo_and_a_move_keep_recovery_while_an_older_save_is_in_flight(bool moved) {
        byte[] earlier;
        string identity = moved ? Identity + "/moved" : Identity;
        using (var document = await StudioNativeDocument.OpenAsync(Copy("working", Source()), null, default, Identity)) {
            await document.RotatePagesAsync([1], 90, default);
            earlier = await document.PrepareSaveAsync(default);
            await document.RotatePagesAsync([1], 90, default);
            byte[] latest = await document.PrepareSaveAsync(default);
            await document.AcknowledgeSaveAsync(Convert.ToHexString(SHA256.HashData(latest)), default);
            await document.UndoAsync(default);
            await document.RedoAsync(default);
            if (moved) await document.SetRecoveryIdentityAsync(identity, default);
            Assert.True(StudioNativeEngine.CreateRecoveryStore().HasSnapshotFiles(identity));
        }
        // The older host save publishes immediately before the process exits, before a presenter can acknowledge it.
        using var next = await StudioNativeDocument.OpenAsync(Copy("next", earlier), null, default, identity);
        Assert.True((await next.GetInfoAsync(default)).HasRecovery);
        await next.RestoreRecoveryAsync(default);
        Assert.Equal(180, (await next.GetInfoAsync(default)).Pages[0].Rotation);
    }

    [Fact]
    public async Task A_selection_with_more_links_than_the_history_limit_is_one_undoable_change() {
        using var document = await StudioNativeDocument.OpenAsync(Copy("links", Source()), null, default);
        var requests = Enumerable.Range(0, 30).Select(index => new StudioNativeAnnotationRequest("Link", 1,
            new(20, 10 + index * 10, 80, 8), LinkUri: "https://evotec.xyz")).ToArray();
        await document.AddAnnotationsAsync(requests, default);
        Assert.Equal(30, (await document.GetAnnotationsAsync(1, default)).Annotations.Count);
        await document.UndoAsync(default);
        Assert.Empty((await document.GetAnnotationsAsync(1, default)).Annotations);
        Assert.False((await document.GetInfoAsync(default)).CanUndo);
    }

    [Fact]
    public async Task The_unmanaged_save_protocol_returns_owned_immutable_bytes_and_acknowledges_the_observed_revision() {
        using var document = await StudioNativeDocument.OpenAsync(Copy("native", Source()), null, default, Identity);
        await document.RotatePagesAsync([1], 90, default);
        AssertUnmanagedSave(document);
        Assert.False((await document.GetInfoAsync(default)).IsDirty);
    }

    private static unsafe void AssertUnmanagedSave(StudioNativeDocument document) {
        GCHandle handle = GCHandle.Alloc(document);
        byte* first = null, second = null, error = null;
        try {
            delegate* unmanaged<nint, nint, byte**, nuint*, byte**, int> prepare = &StudioNativeExports.DocumentPrepareSave;
            delegate* unmanaged<nint, int*, int, int, nint, byte**, int> rotate = &StudioNativeExports.DocumentRotatePages;
            delegate* unmanaged<nint, byte*, byte*, nint, byte**, int> acknowledge = &StudioNativeExports.DocumentAcknowledgeSave;
            nuint length = 0;
            Assert.Equal(0, prepare(GCHandle.ToIntPtr(handle), 0, &first, &length, &error));
            int page = 1;
            Assert.Equal(0, rotate(GCHandle.ToIntPtr(handle), &page, 1, 90, 0, &error));
            byte[] captured = new ReadOnlySpan<byte>(first, checked((int)length)).ToArray();
            Assert.Equal(90, PdfDocument.Load(new System.IO.MemoryStream(captured)).GetPageLayouts()[0].RotationDegrees);
            Assert.Equal(0, prepare(GCHandle.ToIntPtr(handle), 0, &second, &length, &error));
            byte[] current = new ReadOnlySpan<byte>(second, checked((int)length)).ToArray();
            byte[] digest = Encoding.UTF8.GetBytes(Convert.ToHexString(SHA256.HashData(current)) + '\0');
            byte[] identity = Encoding.UTF8.GetBytes(Identity + '\0');
            fixed (byte* hash = digest, source = identity)
                Assert.Equal(0, acknowledge(GCHandle.ToIntPtr(handle), hash, source, 0, &error));
            Assert.Equal((nint)0, (nint)error);
        } finally {
            NativeMemory.Free(first); NativeMemory.Free(second); NativeMemory.Free(error);
            handle.Free();
        }
    }
}
