using System.Security.Cryptography;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Native;

namespace OfficeIMO.Studio.Core.Tests;

public sealed class StudioPublicationResourceTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "studio-publication-resource-" + Guid.NewGuid().ToString("N"));
    private const string Identity = "studio-document://resource-closure/original";
    public StudioPublicationResourceTests() {
        Directory.CreateDirectory(_root);
        StudioNativeEngine.Initialize(Path.Combine(_root, "data"));
    }
    public void Dispose() => Directory.Delete(_root, true);
    private string Source() {
        string source = Path.Combine(_root, "Report.pdf");
        PdfDocument.Create(document => {
            foreach (string text in new[] { "First", "Second" })
                document.Page(page => page.Size(400, 300).Content(content => content.Paragraph(p => p.Text(text))));
        }).Save(source);
        return source;
    }
    private static string Fingerprint(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));

    [Fact]
    public unsafe void C_ordered_publication_keeps_recovery_until_observed_then_retires_superseded_saves() {
        delegate* unmanaged<byte*, byte*, byte*, nint, nint*, byte**, int> open = &StudioNativeExports.DocumentOpenWithRecoveryIdentity;
        delegate* unmanaged<nint, int*, int, int, nint, byte**, int> rotate = &StudioNativeExports.DocumentRotatePages;
        delegate* unmanaged<nint, nint, byte**, nuint*, byte**, int> prepare = &StudioNativeExports.DocumentPrepareSave;
        delegate* unmanaged<nint, byte*, byte*, nint, byte**, int> acknowledge = &StudioNativeExports.DocumentAcknowledgeOrderedSave;
        delegate* unmanaged<nint, void> close = &StudioNativeExports.DocumentClose;
        delegate* unmanaged<void*, void> free = &StudioNativeExports.Free;
        byte* source = StudioNativeAbi.ToUtf8(Source());
        byte* identity = StudioNativeAbi.ToUtf8(Identity);
        byte* mismatch = StudioNativeAbi.ToUtf8(Identity + "-another");
        byte* fingerprint = null;
        byte* error = null;
        nint handle = 0;
        try {
            Assert.Equal(StudioNativeAbi.Ok, open(source, null, identity, 0, &handle, &error));
            int page = 1;
            for (int index = 0; index < 3; index++) {
                Assert.Equal(StudioNativeAbi.Ok, rotate(handle, &page, 1, 90, 0, &error));
                byte* bytes = null;
                nuint length = 0;
                try {
                    Assert.Equal(StudioNativeAbi.Ok, prepare(handle, 0, &bytes, &length, &error));
                    free(fingerprint);
                    fingerprint = StudioNativeAbi.ToUtf8(Convert.ToHexString(SHA256.HashData(new ReadOnlySpan<byte>(bytes, checked((int)length)))));
                } finally { free(bytes); }
            }
            var recovery = StudioNativeEngine.CreateRecoveryStore();
            Assert.True(recovery.HasSnapshotFiles(Identity));
            Assert.Equal(StudioNativeAbi.Ok, acknowledge(handle, fingerprint, mismatch, 0, &error));
            Assert.True(recovery.HasSnapshotFiles(Identity));
            Assert.Equal(StudioNativeAbi.Ok, acknowledge(handle, fingerprint, identity, 0, &error));
            Assert.False(recovery.HasSnapshotFiles(Identity));
            Assert.Null(StudioNativeAbi.FromUtf8(error));
        } finally {
            if (handle != 0) close(handle);
            free(source); free(identity); free(mismatch); free(fingerprint); free(error);
        }
    }

    [Fact]
    public async Task Ordered_acknowledgement_retires_coalesced_preparations_and_recovery() {
        using var document = await StudioNativeDocument.OpenAsync(Source(), null, default, Identity);
        byte[] latest = [];
        for (int index = 0; index < 3; index++) {
            await document.RotatePagesAsync([1], 90, default);
            latest = await document.PrepareSaveAsync(default);
        }
        await document.AcknowledgeSaveAsync(Fingerprint(latest), default, Identity, supersedesEarlierPreparations: true);
        Assert.False((await document.GetInfoAsync(default)).IsDirty);
        Assert.False(StudioNativeEngine.CreateRecoveryStore().HasSnapshotFiles(Identity));
    }

    [Fact]
    public async Task Failed_recovery_retirement_keeps_saved_state_and_preparations_retryable() {
        using var document = await StudioNativeDocument.OpenAsync(Source(), null, default, Identity);
        await document.RotatePagesAsync([1], 90, default);
        byte[] snapshot = await document.PrepareSaveAsync(default);
        string fingerprint = Fingerprint(snapshot);
        string recovery = Directory.GetFiles(Path.Combine(_root, "data", "Recovery"), "*.recovery").Single();
        string backup = recovery + ".held";
        File.Move(recovery, backup);
        Directory.CreateDirectory(recovery);
        try {
            await Assert.ThrowsAsync<UnauthorizedAccessException>(() =>
                document.AcknowledgeSaveAsync(fingerprint, default, Identity, supersedesEarlierPreparations: true));
            Assert.True((await document.GetInfoAsync(default)).IsDirty);
        } finally {
            Directory.Delete(recovery);
            File.Move(backup, recovery);
        }
        await document.AcknowledgeSaveAsync(fingerprint, default, Identity, supersedesEarlierPreparations: true);
        Assert.False((await document.GetInfoAsync(default)).IsDirty);
        Assert.False(StudioNativeEngine.CreateRecoveryStore().HasSnapshotFiles(Identity));
    }

    [Fact]
    public async Task Ordered_acknowledgement_keeps_newer_pending_publications_only() {
        using var document = await StudioNativeDocument.OpenAsync(Source(), null, default, Identity);
        await document.RotatePagesAsync([1], 90, default);
        byte[] first = await document.PrepareSaveAsync(default);
        await document.RotatePagesAsync([1], 90, default);
        byte[] middle = await document.PrepareSaveAsync(default);
        await document.RotatePagesAsync([1], 90, default);
        byte[] latest = await document.PrepareSaveAsync(default);
        await document.AcknowledgeSaveAsync(Fingerprint(middle), default, Identity, supersedesEarlierPreparations: true);
        Assert.True((await document.GetInfoAsync(default)).IsDirty);
        var recovery = StudioNativeEngine.CreateRecoveryStore();
        Assert.Null(recovery.ReadVerifiedSnapshot(Identity, Fingerprint(first)));
        Assert.Equal(latest, recovery.ReadVerifiedSnapshot(Identity, Fingerprint(latest)));
        await document.AcknowledgeSaveAsync(Fingerprint(latest), default, Identity, supersedesEarlierPreparations: true);
        Assert.False(recovery.HasSnapshotFiles(Identity));
    }

    [Fact]
    public async Task Preparation_order_survives_undo_and_repeated_snapshot_bytes() {
        using var document = await StudioNativeDocument.OpenAsync(Source(), null, default, Identity);
        await document.RotatePagesAsync([1], 90, default);
        byte[] first = await document.PrepareSaveAsync(default);
        await document.RotatePagesAsync([1], 90, default);
        byte[] middle = await document.PrepareSaveAsync(default);
        await document.UndoAsync(default);
        byte[] newest = await document.PrepareSaveAsync(default);
        Assert.Equal(first, newest);
        await document.AcknowledgeSaveAsync(Fingerprint(middle), default, Identity, supersedesEarlierPreparations: true);
        Assert.True((await document.GetInfoAsync(default)).IsDirty);
        Assert.Equal(newest, StudioNativeEngine.CreateRecoveryStore().ReadVerifiedSnapshot(Identity, Fingerprint(newest)));
        await document.AcknowledgeSaveAsync(Fingerprint(newest), default, Identity, supersedesEarlierPreparations: true);
        Assert.False((await document.GetInfoAsync(default)).IsDirty);
        Assert.False(StudioNativeEngine.CreateRecoveryStore().HasSnapshotFiles(Identity));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task Repeated_bytes_keep_intervening_publications_recoverable_until_distinct_readback(bool completes, bool observedBeforeNewPreparations) {
        byte[] middle, newest;
        using (var document = await StudioNativeDocument.OpenAsync(Source(), null, default, Identity)) {
            await document.RotatePagesAsync([1], 90, default);
            byte[] first = await document.PrepareSaveAsync(default);
            if (observedBeforeNewPreparations)
                await document.AcknowledgeSaveAsync(Fingerprint(first), default, Identity, supersedesEarlierPreparations: true);
            await document.RotatePagesAsync([1], 90, default);
            middle = await document.PrepareSaveAsync(default);
            await document.UndoAsync(default);
            newest = await document.PrepareSaveAsync(default);
            Assert.Equal(first, newest);
            await document.AcknowledgeSaveAsync(Fingerprint(first), default, Identity, supersedesEarlierPreparations: true);
            Assert.True(StudioNativeEngine.CreateRecoveryStore().HasSnapshotFiles(Identity));
            // Coalesced duplicate callbacks still cannot identify the later occurrence of these same bytes.
            await document.AcknowledgeSaveAsync(Fingerprint(first), default, Identity, supersedesEarlierPreparations: true);
            Assert.Equal(newest, StudioNativeEngine.CreateRecoveryStore().ReadVerifiedSnapshot(Identity, Fingerprint(middle)));
            if (completes) {
                await document.AcknowledgeSaveAsync(Fingerprint(middle), default, Identity, supersedesEarlierPreparations: true);
                await document.AcknowledgeSaveAsync(Fingerprint(newest), default, Identity, supersedesEarlierPreparations: true);
                Assert.False((await document.GetInfoAsync(default)).IsDirty);
                Assert.False(StudioNativeEngine.CreateRecoveryStore().HasSnapshotFiles(Identity));
                return;
            }
        }
        // B publishes, then the later A fails or the process exits before its publication.
        string published = Path.Combine(_root, "published.pdf");
        File.WriteAllBytes(published, middle);
        using var reopened = await StudioNativeDocument.OpenAsync(published, null, default, Identity);
        Assert.True((await reopened.GetInfoAsync(default)).HasRecovery);
        await reopened.RestoreRecoveryAsync(default);
        Assert.Equal(90, (await reopened.GetInfoAsync(default)).Pages[0].Rotation);
        Assert.Equal(newest, await reopened.PrepareSaveAsync(default));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Split_preserves_occupied_files_or_directories_and_uses_next_name(bool directory) {
        string source = Source();
        byte[] original = File.ReadAllBytes(source);
        string folder = Path.Combine(_root, "parts");
        Directory.CreateDirectory(folder);
        string occupied = Path.Combine(folder, "Report p1.pdf");
        if (directory) Directory.CreateDirectory(occupied);
        string sentinel = directory ? Path.Combine(occupied, "keep.txt") : occupied;
        byte[] existing = [1, 3, 5, 7];
        File.WriteAllBytes(sentinel, existing);
        using var document = await StudioNativeDocument.OpenAsync(source, null, default);
        var result = await document.WriteAsync(new(Split: new(folder, PagesPerPart: 1)), default);
        Assert.Equal(new[] { "Report p1 2.pdf", "Report p2.pdf" }, result.Files.Select(Path.GetFileName));
        Assert.Equal(existing, File.ReadAllBytes(sentinel));
        Assert.Equal(original, File.ReadAllBytes(source));
        for (int index = 0; index < result.Files.Count; index++) {
            using var part = await StudioNativeDocument.OpenAsync(result.Files[index], null, default);
            Assert.Single((await part.GetInfoAsync(default)).Pages);
            Assert.Contains(index == 0 ? "First" : "Second", string.Concat((await part.GetPageTextAsync(1, default)).Texts));
        }
    }
}
