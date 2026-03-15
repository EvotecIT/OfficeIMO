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
    public void TryReadContainerAwareFenceRun_ParsesQuotedFenceWithLanguage() {
        var ok = MarkdownFence.TryReadContainerAwareFenceRun("> ```csharp", out var prefix, out var marker, out var runLength, out var suffix);

        Assert.True(ok);
        Assert.Equal("> ", prefix);
        Assert.Equal('`', marker);
        Assert.Equal(3, runLength);
        Assert.Equal("csharp", suffix);
    }

    [Fact]
    public void TryReadContainerAwareFenceRun_ParsesNestedQuotedFenceWithLanguage() {
        var ok = MarkdownFence.TryReadContainerAwareFenceRun("> > ```json", out var prefix, out var marker, out var runLength, out var suffix);

        Assert.True(ok);
        Assert.Equal("> > ", prefix);
        Assert.Equal('`', marker);
        Assert.Equal(3, runLength);
        Assert.Equal("json", suffix);
    }

    [Fact]
    public void TryReadContainerAwareFenceRun_ParsesListIndentedQuotedFenceWithLanguage() {
        var ok = MarkdownFence.TryReadContainerAwareFenceRun("  > ```mermaid", out var prefix, out var marker, out var runLength, out var suffix);

        Assert.True(ok);
        Assert.Equal("  > ", prefix);
        Assert.Equal('`', marker);
        Assert.Equal(3, runLength);
        Assert.Equal("mermaid", suffix);
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

    [Fact]
    public void ApplyTransformOutsideFencedCodeBlocks_PreservesQuotedFenceBodies() {
        var markdown = """
outside

> ```json
> {"nodes":[{"id":"A"}]}
> ```

tail
""";

        var transformed = MarkdownFence.ApplyTransformOutsideFencedCodeBlocks(markdown, static segment => segment.Replace("outside", "OUTSIDE").Replace("tail", "TAIL"));

        Assert.Contains("OUTSIDE", transformed, StringComparison.Ordinal);
        Assert.Contains("TAIL", transformed, StringComparison.Ordinal);
        Assert.Contains("> ```json", transformed, StringComparison.Ordinal);
        Assert.Contains("> {\"nodes\":[{\"id\":\"A\"}]}", transformed, StringComparison.Ordinal);
    }

    [Fact]
    public void ApplyTransformOutsideFencedCodeBlocks_DoesNotTouchFenceLikeRunsInsideLongerOuterFence() {
        var markdown = """
before

````markdown
```json
{"hello":"world"}
```
````

after
""";

        var transformed = MarkdownFence.ApplyTransformOutsideFencedCodeBlocks(markdown, static segment => segment.Replace("before", "BEFORE").Replace("after", "AFTER").Replace("world", "changed"));

        Assert.Contains("BEFORE", transformed, StringComparison.Ordinal);
        Assert.Contains("AFTER", transformed, StringComparison.Ordinal);
        Assert.Contains("{\"hello\":\"world\"}", transformed, StringComparison.Ordinal);
        Assert.DoesNotContain("{\"hello\":\"changed\"}", transformed, StringComparison.Ordinal);
    }
}
