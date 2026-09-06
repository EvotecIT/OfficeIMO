using Avalonia.Platform.Storage;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioStorageInputTests {
    internal static IStorageFile CreateImageFile(byte[] bytes) =>
        new StorageFile(() => Task.FromResult<Stream>(new InputStream(bytes, false))).Item;

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task ReadsProviderStreamAndReleasesStreamBeforePermission(bool seekable) {
        var stream = new InputStream([1, 2, 3], seekable);
        var file = new StorageFile(() => Task.FromResult<Stream>(stream), () => Assert.True(stream.Disposed));
        byte[]? result = await StudioStorageInput.ReadImageAsync([file.Item], default, 3);
        Assert.Equal(new byte[] { 1, 2, 3 }, result);
        Assert.Equal(1, file.DisposeCount);
        Assert.Equal(1, file.OpenCount);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task RejectsOversizeStreamsAndReleasesAccess(bool seekable) {
        var stream = new InputStream([1, 2, 3, 4], seekable);
        var file = new StorageFile(() => Task.FromResult<Stream>(stream));
        await Assert.ThrowsAsync<InvalidDataException>(() => StudioStorageInput.ReadImageAsync([file.Item], default, 3));
        Assert.True(stream.Disposed);
        Assert.Equal(1, file.DisposeCount);
    }

    [Fact]
    public async Task CancellationAfterPickerReleasesAllItemsWithoutOpening() {
        var first = new StorageFile(() => throw new InvalidOperationException("Must not open"));
        var second = new StorageFile(() => throw new InvalidOperationException("Must not open"));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            StudioStorageInput.ReadImageAsync([first.Item, second.Item, first.Item], new CancellationToken(true)));
        Assert.Equal(1, first.DisposeCount);
        Assert.Equal(1, second.DisposeCount);
        Assert.Equal(0, first.OpenCount);
    }

    [Fact]
    public async Task CancellationWhileProviderOpensDisposesItsLateStream() {
        using var cancellation = new CancellationTokenSource();
        var stream = new InputStream([1], false);
        var file = new StorageFile(() => {
            cancellation.Cancel();
            return Task.FromResult<Stream>(stream);
        });
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            StudioStorageInput.ReadImageAsync([file.Item], cancellation.Token));
        Assert.True(stream.Disposed);
        Assert.Equal(1, file.DisposeCount);
    }

    [Fact]
    public async Task PermissionFailureReleasesPickerItem() {
        var file = new StorageFile(() => throw new UnauthorizedAccessException("Access expired"));
        await Assert.ThrowsAsync<UnauthorizedAccessException>(() => StudioStorageInput.ReadImageAsync([file.Item], default));
        Assert.Equal(1, file.DisposeCount);
    }

    [Fact]
    public async Task EmptySelectionReturnsNoImage() =>
        Assert.Null(await StudioStorageInput.ReadImageAsync([], default));

    private sealed class InputStream(byte[] bytes, bool seekable) : MemoryStream(bytes) {
        internal bool Disposed { get; private set; }
        public override bool CanSeek => seekable;
        protected override void Dispose(bool disposing) { Disposed = true; base.Dispose(disposing); }
    }

    // Avalonia marks storage interfaces as non-implementable in C#; a runtime proxy
    // supplies a provider boundary without relying on a real desktop picker.
    private sealed class StorageFile {
        internal IStorageFile Item { get; }
        internal int DisposeCount { get; private set; }
        internal int OpenCount { get; private set; }
        internal StorageFile(Func<Task<Stream>> open, Action? onDispose = null) {
            Item = System.Reflection.DispatchProxy.Create<IStorageFile, StorageProxy>();
            ((StorageProxy)(object)Item).Call = name => name switch {
                "OpenReadAsync" => Open(),
                "Dispose" => Release(),
                "get_Name" => "image.png",
                "get_Path" => new Uri("content://provider/image.png"),
                _ => throw new NotSupportedException(name)
            };
            Task<Stream> Open() { OpenCount++; return open(); }
            object? Release() { DisposeCount++; onDispose?.Invoke(); return null; }
        }
    }

    public class StorageProxy : System.Reflection.DispatchProxy {
        internal Func<string, object?> Call { get; set; } = null!;
        protected override object? Invoke(System.Reflection.MethodInfo? method, object?[]? args) => Call(method!.Name);
    }
}
