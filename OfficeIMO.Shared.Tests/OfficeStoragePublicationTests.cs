using System;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class OfficeStoragePublicationTests {
    [Fact]
    public async Task StreamSnapshotIsSeekablePrivateAndRemovedAfterUse() {
        string path;
        string directory;
        byte[] bytes = new byte[] { 1, 2, 3, 4 };
        using (var snapshot = await OfficeStreamFileSnapshot.CaptureAsync(_ => Task.FromResult<Stream>(new MemoryStream(bytes)),
                   ".pdf", 20, null, default)) {
            path = snapshot.FilePath;
            directory = Path.GetDirectoryName(path)!;
            Assert.Equal(bytes, File.ReadAllBytes(path));
            Assert.Equal(4, snapshot.Length);
            Assert.Equal(64, snapshot.Fingerprint.Length);
#if NET6_0_OR_GREATER
            if (!OperatingSystem.IsWindows()) {
                Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute, File.GetUnixFileMode(directory));
                Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite, File.GetUnixFileMode(path));
            }
#endif
            await snapshot.VerifySourceAsync(_ => Task.FromResult<Stream>(new MemoryStream(bytes)), 20, default);
        }
        Assert.False(File.Exists(path));
        Assert.False(Directory.Exists(directory));
    }

    [Theory]
    [InlineData(".pdf/../../escape")]
    [InlineData(".pdf\\escape")]
    [InlineData(".pdf:stream")]
    public async Task StreamSnapshotRejectsUnsafeExtensionsBeforeOpeningSource(string extension) {
        bool opened = false;
        await Assert.ThrowsAsync<ArgumentException>(() => OfficeStreamFileSnapshot.CaptureAsync(_ => {
            opened = true;
            return Task.FromResult<Stream>(new MemoryStream());
        }, extension, 20, null, default));
        Assert.False(opened);
    }

    [Fact]
    public async Task SerializationLimitAppliesBeforeBufferGrowthForEveryWriteSurface() {
        using var stream = new OfficeBoundedMemoryStream(5);
        await stream.WriteAsync(new byte[] { 1, 2, 3, 4 }, 0, 4, default);
        stream.WriteByte(5);
        Assert.Equal(5, stream.Length);
        Assert.True(stream.Capacity <= 5);
        Assert.Throws<InvalidDataException>(() => stream.WriteByte(6));
        await Assert.ThrowsAsync<InvalidDataException>(() => stream.WriteAsync(new byte[] { 6 }, 0, 1, default));
#if NET8_0_OR_GREATER
        await Assert.ThrowsAsync<InvalidDataException>(() => stream.WriteAsync(new byte[] { 6 }.AsMemory()).AsTask());
#endif
        Assert.Throws<InvalidDataException>(() => stream.SetLength(6));
        Assert.Throws<InvalidDataException>(() => stream.Capacity = 6);
        Assert.Equal(new byte[] { 1, 2, 3, 4, 5 }, stream.ToArray());
    }

    [Fact]
    public void InvalidWriteDoesNotGrowBufferAndProvenanceKeepsItsFailureType() {
        using var stream = new OfficeBoundedMemoryStream(16);
        Assert.Throws<ArgumentException>(() => stream.Write(new byte[1], 0, 16));
        Assert.Equal(0, stream.Capacity);
        using var provenance = new OfficeProvenanceBoundedMemoryStream(1);
        provenance.WriteByte(1);
        Assert.True(OfficeProvenanceLimitException.IsOutput(Assert.Throws<InvalidDataException>(() => provenance.WriteByte(2))));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task FingerprintRejectsOversizeInputAndReleasesIt(bool seekable) {
        var stream = new ReadStream(new byte[4], seekable);
        await Assert.ThrowsAsync<InvalidDataException>(() => OfficeStreamPublication.ReadFingerprintAsync(
            _ => Task.FromResult<Stream>(stream), 3, default));
        Assert.True(stream.Disposed);
    }

    [Fact]
    public async Task ChangedSourceNeverOpensDestructiveWriteStream() {
        int writes = 0;
        IOException error = await Assert.ThrowsAsync<IOException>(() => OfficeStreamPublication.WriteVerifiedAsync(
            _ => Task.FromResult<Stream>(new MemoryStream([2])),
            _ => { writes++; return Task.FromResult<Stream>(new MemoryStream()); },
            [3], Fingerprint([1]), 20, default));
        Assert.Equal(0, writes);
        Assert.False(OfficeStreamPublication.MayHaveChangedDestination(error));
    }

    [Fact]
    public async Task ReadBackOccursAfterWriteStreamClosesAndRejectsProviderCorruption() {
        byte[] retained = [1];
        bool closed = false;
        IOException error = await Assert.ThrowsAsync<IOException>(() => OfficeStreamPublication.WriteVerifiedAsync(
            _ => Task.FromResult<Stream>(new MemoryStream(retained)),
            _ => Task.FromResult<Stream>(new CommitStream(bytes => { closed = true; retained = [.. bytes, 99]; })),
            [2, 3], Fingerprint([1]), 20, default));
        Assert.True(closed);
        Assert.True(OfficeStreamPublication.MayHaveChangedDestination(error));
        Assert.Equal(new byte[] { 2, 3, 99 }, retained);
    }

    [Fact]
    public async Task CancelledLateOpenReleasesItsStreamWithoutWriting() {
        using var cancellation = new CancellationTokenSource();
        var stream = new CommitStream(_ => { });
        OperationCanceledException error = await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfficeStreamPublication.WriteVerifiedAsync(
            _ => Task.FromResult<Stream>(new MemoryStream()),
            _ => { cancellation.Cancel(); return Task.FromResult<Stream>(stream); },
            [1, 2], null, 20, cancellation.Token));
        Assert.True(stream.Disposed);
        Assert.True(OfficeStreamPublication.MayHaveChangedDestination(error));
        Assert.Empty(stream.Published!);
    }

    [Fact]
    public void ProviderIdentityDoesNotBecomeLocalOrInheritWindowsCaseRules() {
        const string first = "content://provider/Documents/Case.pdf";
        const string second = "content://provider/Documents/case.pdf";
        Assert.Null(OfficeStorageIdentity.GetLocalPath(first));
        Assert.Equal(first, OfficeStorageIdentity.Normalize(first));
        Assert.NotEqual(OfficeStorageIdentity.GetPersistenceKey(first), OfficeStorageIdentity.GetPersistenceKey(second));
        Assert.False(OfficeStorageIdentity.AreEquivalent(first, second));
        Assert.Equal("Case.pdf", OfficeStorageIdentity.GetFileName(first));
        string path = Path.Combine(Path.GetTempPath(), "space and #.pdf");
        Assert.Equal(Path.GetFullPath(path), OfficeStorageIdentity.Normalize(new Uri(path).AbsoluteUri));
    }

    private static string Fingerprint(byte[] bytes) {
        using var hash = SHA256.Create();
        return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", string.Empty);
    }

    private sealed class ReadStream(byte[] bytes, bool seekable) : MemoryStream(bytes) {
        internal bool Disposed { get; private set; }
        public override bool CanSeek => seekable;
        protected override void Dispose(bool disposing) { Disposed = true; base.Dispose(disposing); }
    }

    private sealed class CommitStream(Action<byte[]> commit) : MemoryStream {
        internal bool Disposed { get; private set; }
        internal byte[]? Published { get; private set; }
        protected override void Dispose(bool disposing) {
            if (!Disposed) { Disposed = true; Published = ToArray(); commit(Published); }
            base.Dispose(disposing);
        }
    }
}
