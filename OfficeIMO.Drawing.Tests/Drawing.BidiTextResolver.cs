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

    [Theory]
    [InlineData("\u200E", 2)]
    [InlineData("\u200F", 1)]
    [InlineData("\u061C", 1)]
    public void BidiTextResolver_UsesDirectionalMarksAsFirstStrongFsiSignals(string mark, int expectedNeutralLevel) {
        IReadOnlyList<OfficeBidiTextRun> runs = OfficeBidiTextResolver.ResolveRuns(
            "\u2068" + mark + "-א\u2069",
            OfficeTextDirection.LeftToRight);

        Assert.Equal(expectedNeutralLevel, Assert.Single(runs, static run => run.Text.Contains("-", StringComparison.Ordinal)).EmbeddingLevel);
    }

    [Fact]
    public void BidiTextResolver_ResolvesDeepFsiInputWithLinearPreprocessing() {
        const int isolateCount = 10_000;
        string nested = new string('\u2068', isolateCount) + "שלום" + new string('\u2069', isolateCount);

        IReadOnlyList<OfficeBidiTextRun> runs = OfficeBidiTextResolver.ResolveRuns(
            nested,
            OfficeTextDirection.LeftToRight);

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
    public void BidiTextResolver_ResolvesNeutralPunctuationFromBothStrongNeighbors() {
        string visual = OfficeBidiTextResolver.ToVisualOrder(
            "abc (אבג) def",
            OfficeTextDirection.LeftToRight);

        Assert.Equal("abc (גבא) def", visual);
    }

    [Theory]
    [InlineData("אב\u202Bcd\u202C!גד")]
    [InlineData("אב\u2067cd\u2069!גד")]
    public void BidiTextResolver_RestoresOuterStrongContextAfterDirectionalScope(string logical) {
        IReadOnlyList<OfficeBidiTextRun> runs = OfficeBidiTextResolver.ResolveRuns(
            logical,
            OfficeTextDirection.LeftToRight);

        OfficeBidiTextRun punctuation = Assert.Single(runs, static run => run.Text.Contains("!", StringComparison.Ordinal));
        Assert.Equal(OfficeTextDirection.RightToLeft, punctuation.Direction);
        Assert.Equal(1, punctuation.EmbeddingLevel);
    }

    [Theory]
    [InlineData("A\u202B!b\u202C", "Ab!")]
    [InlineData("A\u2067!b\u2069", "Ab!")]
    public void BidiTextResolver_InitializesNeutralContextInsideDirectionalScopes(string logical, string expectedVisual) {
        IReadOnlyList<OfficeBidiTextRun> runs = OfficeBidiTextResolver.ResolveRuns(
            logical,
            OfficeTextDirection.LeftToRight);

        OfficeBidiTextRun punctuation = Assert.Single(runs, static run => run.Text.Contains("!", StringComparison.Ordinal));
        Assert.Equal(OfficeTextDirection.RightToLeft, punctuation.Direction);
        Assert.Equal(1, punctuation.EmbeddingLevel);
        Assert.Equal(expectedVisual, OfficeBidiTextResolver.ToVisualOrder(logical, OfficeTextDirection.LeftToRight));
    }

    [Theory]
    [InlineData("אב\nג", "בא\nג")]
    [InlineData("אב\r\nג", "בא\r\nג")]
    [InlineData("אב\u2028ג", "בא\u2028ג")]
    [InlineData("אב\u2029ג", "בא\u2029ג")]
    public void BidiTextResolver_ReordersParagraphsWithoutMovingSeparators(string logical, string expectedVisual) {
        Assert.Equal(expectedVisual, OfficeBidiTextResolver.ToVisualOrder(logical, OfficeTextDirection.RightToLeft));
    }

    [Theory]
    [InlineData("\u200F")]
    [InlineData("\u061C")]
    public void BidiTextResolver_UsesTrailingRtlMarksAsStrongNeutralContext(string mark) {
        IReadOnlyList<OfficeBidiTextRun> runs = OfficeBidiTextResolver.ResolveRuns(
            "אבג!" + mark,
            OfficeTextDirection.LeftToRight);

        OfficeBidiTextRun punctuation = Assert.Single(runs, static run => run.Text.Contains("!", StringComparison.Ordinal));
        Assert.Equal(OfficeTextDirection.RightToLeft, punctuation.Direction);
        Assert.Equal(1, punctuation.EmbeddingLevel);
    }

    [Fact]
    public void BidiTextResolver_UsesTrailingLrmAsStrongNeutralContext() {
        IReadOnlyList<OfficeBidiTextRun> runs = OfficeBidiTextResolver.ResolveRuns(
            "abc!\u200E",
            OfficeTextDirection.RightToLeft);

        OfficeBidiTextRun punctuation = Assert.Single(runs, static run => run.Text.Contains("!", StringComparison.Ordinal));
        Assert.Equal(OfficeTextDirection.LeftToRight, punctuation.Direction);
        Assert.Equal(2, punctuation.EmbeddingLevel);
    }

    [Fact]
    public void BidiTextResolver_MirrorsPairedPunctuationAtOddLevels() {
        string visual = OfficeBidiTextResolver.ToVisualOrder(
            "(אבג)",
            OfficeTextDirection.RightToLeft);

        Assert.Equal("(גבא)", visual);
    }

    [Fact]
    public void BidiTextResolver_MirrorsUnicodeOperatorsAtOddLevels() {
        Assert.Equal("≥∋", OfficeBidiTextResolver.MirrorText("≤∈"));
        Assert.Equal("≤∈", OfficeBidiTextResolver.MirrorText("≥∋"));
    }

    [Fact]
    public void BidiTextResolver_ExposesVisualRunOrderWithoutReversingRunText() {
        string logical = "A\u2067שלום abc\u2069B";

        IReadOnlyList<OfficeBidiTextRun> runs = OfficeBidiTextResolver.ResolveVisualRuns(
            logical,
            OfficeTextDirection.LeftToRight);

        int latinIndex = runs.ToList().FindIndex(static run => run.Text.Contains("abc", StringComparison.Ordinal));
        int hebrewIndex = runs.ToList().FindIndex(static run => run.Text.Contains("שלום", StringComparison.Ordinal));
        Assert.True(latinIndex >= 0 && hebrewIndex >= 0 && latinIndex < hebrewIndex);
        Assert.Contains(runs, static run => run.Text.Contains("שלום", StringComparison.Ordinal));
    }

    [Fact]
    public void BidiTextResolver_OverflowControlsCannotPopAnAcceptedEmbedding() {
        string logical = "A\u202B" + new string('\u202A', 200) + new string('\u202C', 200) + "שלום\u202CB";

        string visual = OfficeBidiTextResolver.ToVisualOrder(logical, OfficeTextDirection.LeftToRight);

        Assert.Equal("AםולשB", visual);
    }

    [Fact]
    public void BidiTextResolver_OverflowIsolateCannotClearPendingEmbeddingOverflow() {
        string logical = "A\u202E"
            + new string('\u202A', 200)
            + "\u2067\u2069"
            + new string('\u202C', 200)
            + "abc\u202CB";

        string visual = OfficeBidiTextResolver.ToVisualOrder(logical, OfficeTextDirection.LeftToRight);

        Assert.Equal("AcbaB", visual);
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
