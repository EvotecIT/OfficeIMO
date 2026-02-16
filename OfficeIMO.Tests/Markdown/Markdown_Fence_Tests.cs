using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Fence_Tests {
    [Fact]
    public void TryReadFenceRun_ParsesBacktickFenceWithLanguage() {
        var ok = MarkdownFence.TryReadFenceRun("   ```csharp", out var marker, out var runLength, out var suffix);

        Assert.True(ok);
        Assert.Equal('`', marker);
        Assert.Equal(3, runLength);
        Assert.Equal("csharp", suffix);
    }

    [Fact]
    public void TryReadFenceRun_RejectsShortRuns() {
        var ok = MarkdownFence.TryReadFenceRun("~~text", out _, out _, out _);

        Assert.False(ok);
    }

    [Fact]
    public void BuildSafeFence_PicksShortestSafeMarker() {
        var fence = MarkdownFence.BuildSafeFence("````\n~");

        Assert.Equal("~~~", fence);
    }

    [Fact]
    public void BuildSafeFence_GrowsWhenBothMarkersAppearInContent() {
        var fence = MarkdownFence.BuildSafeFence("```\n~~~~");

        Assert.Equal("````", fence);
    }
}
