using System.Runtime.CompilerServices;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingImageCacheOwnershipTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CacheIdentityDoesNotOwnUnbudgetedSourceMemory(bool encoded) {
        var cache = new OfficeRasterImageCache(64);
        WeakReference source = AddSource(cache, encoded);
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        Assert.False(source.IsAlive);
        Assert.Equal(4, cache.UsedBytes);
        GC.KeepAlive(cache);
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static WeakReference AddSource(OfficeRasterImageCache cache, bool encoded) {
        var tiny = new OfficeRasterImage(1, 1, OfficeColor.Red);
        if (encoded) {
            var bytes = new byte[16384];
            cache.AddDecoded(bytes, tiny);
            Assert.True(cache.TryGetDecoded(bytes, out var hit));
            Assert.Same(tiny, hit);
            return new WeakReference(bytes);
        }
        var image = new OfficeRasterImage(64, 64, OfficeColor.Red);
        cache.AddFiltered(image, 1, 1, tiny);
        Assert.True(cache.TryGetFiltered(image, 1, 1, out var filtered));
        Assert.Same(tiny, filtered);
        return new WeakReference(image);
    }
}
