using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPageClipPathIntersectionChainTests {
    [Fact]
    public void AppendingToADeepChainCostsOnlyTheNewPaths() {
        // Charging the whole chain on every append would need about 4 million units here, four times the default budget.
        var budget = new PdfTextClippingBudget();

        PdfPageClipPath clip = BuildChain(budget, 4000);

        Assert.Equal(4000, clip.RenderingIntersectionCount);
    }

    [Fact]
    public void CombiningTwoChainsIsChargedToThePageClippingBudget() {
        PdfPageClipPath first = BuildChain(new PdfTextClippingBudget(), 200);
        PdfPageClipPath second = BuildChain(new PdfTextClippingBudget(), 200);
        var tightBudget = new PdfTextClippingBudget(new PdfReadLimits { MaxClippingIntersectionWorkPerPage = 100 });

        Assert.Throws<PdfReadLimitException>(() => tightBudget.ResolveActiveClip(first, second));
        PdfPageClipPath combined = new PdfTextClippingBudget().ResolveActiveClip(first, second);
        Assert.Equal(400, combined.RenderingIntersectionCount);
        Assert.Equal(200, new PdfTextClippingBudget().ResolveActiveClip(first, first).RenderingIntersectionCount);
    }

    // Every second U-shaped clip meets a non-convex clip, so each pair adds both paths to the chain.
    private static PdfPageClipPath BuildChain(PdfTextClippingBudget budget, int count) {
        PdfPageClipPath? clip = null;
        for (int i = 0; i < count; i++) clip = budget.ResolveActiveClip(clip, CreateUShape());
        return clip!.Value;
    }

    private static PdfPageClipPath CreateUShape() {
        var commands = new List<OfficePathCommand> {
            OfficePathCommand.MoveTo(new OfficePoint(0, 0)),
            OfficePathCommand.LineTo(new OfficePoint(240, 0)),
            OfficePathCommand.LineTo(new OfficePoint(240, 200)),
            OfficePathCommand.LineTo(new OfficePoint(130, 200)),
            OfficePathCommand.LineTo(new OfficePoint(130, 100)),
            OfficePathCommand.LineTo(new OfficePoint(110, 100)),
            OfficePathCommand.LineTo(new OfficePoint(110, 200)),
            OfficePathCommand.LineTo(new OfficePoint(0, 200)),
            OfficePathCommand.Close()
        };
        Assert.True(PdfPageClipPath.TryCreatePath(commands, OfficeFillRule.NonZero, out PdfPageClipPath clip));
        return clip;
    }
}
