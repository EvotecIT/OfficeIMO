using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlProvenanceWave71Tests {
    [Fact]
    public void AnimationShorthandRejectsNonFiniteIterationCounts() {
        Assert.False(HtmlResourcePipeline.TryExpandAnimationShorthandNames("1s Infinity spin", out _));
    }
}
