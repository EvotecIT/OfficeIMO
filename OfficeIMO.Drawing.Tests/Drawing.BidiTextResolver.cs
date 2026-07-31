using OfficeIMO.Drawing;
using System;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Drawing {
    [Fact]
    public void BidiTextResolver_ResolvesEmbeddingsOverridesAndIsolatesWithoutPaintingControls() {
        IReadOnlyList<OfficeBidiTextRun> runs = OfficeBidiTextResolver.ResolveRuns(
            "left \u202Eabc\u202C middle \u2067שלום\u2069 right",
            OfficeTextDirection.LeftToRight);

        Assert.Equal("left abc middle שלום right", string.Concat(runs.Select(static run => run.Text)));
        Assert.DoesNotContain(runs, run => OfficeTextElements.ContainsBidiControl(run.Text));
        Assert.Contains(runs, run => run.Text == "abc" && run.Direction == OfficeTextDirection.RightToLeft);
        Assert.Contains(runs, run => run.Text == "שלום" && run.Direction == OfficeTextDirection.RightToLeft);
        Assert.Equal("left cba middle םולש right", OfficeBidiTextResolver.ToVisualOrder(
            "left \u202Eabc\u202C middle \u2067שלום\u2069 right",
            OfficeTextDirection.LeftToRight));
    }

    [Fact]
    public void BidiTextResolver_BoundsNestingAndUsesFirstStrongDirectionForFsi() {
        string nested = new string('\u2068', 200) + "שלום" + new string('\u2069', 200);

        IReadOnlyList<OfficeBidiTextRun> runs = OfficeBidiTextResolver.ResolveRuns(nested);

        OfficeBidiTextRun run = Assert.Single(runs);
        Assert.Equal("שלום", run.Text);
        Assert.Equal(OfficeTextDirection.RightToLeft, run.Direction);
    }

    [Fact]
    public void BidiTextResolver_PreservesNestedIsolateEmbeddingLevels() {
        string logical = "A\u2067שלום abc\u2069B";

        string visual = OfficeBidiTextResolver.ToVisualOrder(logical, OfficeTextDirection.LeftToRight);

        Assert.Equal("Aabc םולשB", visual);
    }

    [Fact]
    public void BidiTextResolver_OverflowControlsCannotPopAnAcceptedEmbedding() {
        string logical = "A\u202B" + new string('\u202A', 200) + new string('\u202C', 200) + "שלום\u202CB";

        string visual = OfficeBidiTextResolver.ToVisualOrder(logical, OfficeTextDirection.LeftToRight);

        Assert.Equal("AםולשB", visual);
    }

    [Fact]
    public void BidiTextResolver_HonorsCancellationDuringBoundedResolution() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => OfficeBidiTextResolver.ResolveRuns(
            new string('a', 1024),
            OfficeTextDirection.LeftToRight,
            cancellation.Token));
    }
}
