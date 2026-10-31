using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OfficeIMO.Reader;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderRegularExpressionSearchTests {
    [Theory]
    [InlineData(1)]
    [InlineData(20_000)]
    public void DistinctRegexValuesCorrelateAcrossNormalizedPageFragments(int limit) {
        string[] ids = Enumerable.Range(0, 20_000).Select(index => "INV-" + index.ToString("D5")).ToArray();
        var document = Document(string.Join(" ", ids),
            string.Join("  ", ids.Take(10_000)).ToLowerInvariant(),
            "  " + string.Join("  ", ids.Skip(10_000)).ToLowerInvariant());
        var result = document.Search(@"INV-\d{5}", new OfficeDocumentSearchOptions {
            UseRegularExpressions = true, MaximumResults = limit
        });
        Assert.Equal(limit, result.Hits.Count);
        Assert.All(result.Hits.Take(10_000), hit => Assert.Equal(1, Assert.Single(hit.Pages).Number));
        Assert.All(result.Hits.Skip(10_000), hit => Assert.Equal(2, Assert.Single(hit.Pages).Number));
        Assert.Equal(Enumerable.Range(0, limit).Select(index => index * 10), result.Hits.Select(hit => hit.StartIndex));
        Assert.True(result.MaximumResultsReached);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RepeatedRegexMatchesKeepOffsetsWithoutRepeatedPageScans(bool normalizedPages) {
        const int count = 10_000;
        string text = string.Join(" ", Enumerable.Repeat("a", count));
        var document = normalizedPages
            ? Document(text, string.Join("  ", Enumerable.Repeat("A", count / 2)),
                "  " + string.Join("  ", Enumerable.Repeat("A", count / 2)))
            : Document(text);
        var result = document.Search("a", new OfficeDocumentSearchOptions {
            UseRegularExpressions = true, MaximumResults = int.MaxValue
        });
        Assert.Equal(count, result.Hits.Count);
        Assert.Equal(Enumerable.Range(0, count).Select(index => index * 2), result.Hits.Select(hit => hit.StartIndex));
        if (normalizedPages) {
            Assert.All(result.Hits.Take(count / 2), hit => Assert.Equal(1, Assert.Single(hit.Pages).Number));
            Assert.All(result.Hits.Skip(count / 2), hit => Assert.Equal(2, Assert.Single(hit.Pages).Number));
        } else Assert.All(result.Hits, hit => Assert.Empty(hit.Pages));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CancellationAfterBlockEnumerationStartsStopsBothSearchModes(bool regex) {
        using var cancellation = new CancellationTokenSource();
        var document = Document("A12", "A12");
        document.Blocks = new CancellingBlocks(document.Blocks[0], cancellation);
        var error = Assert.Throws<OperationCanceledException>(() => document.Search("A12",
            new OfficeDocumentSearchOptions { UseRegularExpressions = regex }, cancellation.Token));
        Assert.Equal(cancellation.Token, error.CancellationToken);
    }

    private sealed class CancellingBlocks : IReadOnlyList<OfficeDocumentBlock> {
        private readonly OfficeDocumentBlock _block;
        private readonly CancellationTokenSource _cancellation;
        internal CancellingBlocks(OfficeDocumentBlock block, CancellationTokenSource cancellation) {
            _block = block;
            _cancellation = cancellation;
        }
        public int Count => 1;
        public OfficeDocumentBlock this[int index] => index == 0 ? _block : throw new ArgumentOutOfRangeException(nameof(index));
        public IEnumerator<OfficeDocumentBlock> GetEnumerator() {
            _cancellation.Cancel();
            yield return _block;
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    [Fact]
    public void VariableLengthMatchesKeepOffsetsAndPagesAcrossAFragmentBoundary() {
        var document = Document("A1 A222 B44", "A1 A2", "22 B44");
        var result = document.Search(@"[ab]\d+", new OfficeDocumentSearchOptions { UseRegularExpressions = true });
        Assert.Equal(new[] { 0, 3, 8 }, result.Hits.Select(hit => hit.StartIndex));
        Assert.Equal(new[] { 2, 4, 3 }, result.Hits.Select(hit => hit.Length));
        Assert.Equal(new[] { 1 }, result.Hits[0].Pages.Select(page => page.Number!.Value));
        Assert.Equal(new[] { 1, 2 }, result.Hits[1].Pages.Select(page => page.Number!.Value));
        Assert.Equal(new[] { 2 }, result.Hits[2].Pages.Select(page => page.Number!.Value));
        Assert.False(result.MaximumResultsReached);
        Assert.Empty(document.Search(@"[ab]\d+").Hits);
        Assert.Empty(document.Search(@"[ab]\d+", new OfficeDocumentSearchOptions { UseRegularExpressions = true, MatchCase = true }).Hits);
    }

    [Fact]
    public void EmptyMatchesTerminateAndRetainTheirSourceEndpoints() {
        var document = Document("alpha omega", "alpha ", "omega");
        var result = document.Search("^|$", new OfficeDocumentSearchOptions { UseRegularExpressions = true });
        Assert.Equal(new[] { 0, 11 }, result.Hits.Select(hit => hit.StartIndex));
        Assert.All(result.Hits, hit => Assert.Equal(0, hit.Length));
        Assert.Equal(new[] { 1 }, result.Hits[0].Pages.Select(page => page.Number!.Value));
        Assert.Equal(new[] { 2 }, result.Hits[1].Pages.Select(page => page.Number!.Value));
    }

    [Fact]
    public void ResultLimitAndWholeWordKeepRegexMatchesBounded() {
        var document = Document("A12 A123 xA1x", "A12 ", "A123 xA1x");
        var result = document.Search(@"A\d+", new OfficeDocumentSearchOptions {
            UseRegularExpressions = true, WholeWord = true, MaximumResults = 1
        });
        Assert.True(result.MaximumResultsReached);
        Assert.Equal(3, Assert.Single(result.Hits).Length);
        var all = document.Search(@"A\d+", new OfficeDocumentSearchOptions { UseRegularExpressions = true, WholeWord = true });
        Assert.Equal(2, all.Hits.Count);
    }

    [Fact]
    public void RepeatedWholePageCopiesUseMatchedTextAndWholeWordEvidence() {
        var document = Document("A12", "A12", "A12", "xA12x");
        var result = document.Search(@"A\d+", new OfficeDocumentSearchOptions { UseRegularExpressions = true, WholeWord = true });
        Assert.Equal(new[] { 1, 2 }, Assert.Single(result.Hits).Pages.Select(page => page.Number!.Value));
    }

    [Fact]
    public void ConflictingFragmentsDoNotInventRegexCitations() {
        var document = Document("A1 A1", "A1");
        var result = document.Search(@"A\d+", new OfficeDocumentSearchOptions { UseRegularExpressions = true });
        Assert.Equal(2, result.Hits.Count);
        Assert.All(result.Hits, hit => Assert.Empty(hit.Pages));
    }

    [Fact]
    public void CaseInsensitiveRegexCitationsKeepNormalizedPageCase() {
        var document = Document("A12", "a12");
        var result = document.Search(@"a\d+", new OfficeDocumentSearchOptions { UseRegularExpressions = true });
        Assert.Equal(new[] { 1 }, Assert.Single(result.Hits).Pages.Select(page => page.Number!.Value));
        Assert.Empty(document.Search(@"a\d+", new OfficeDocumentSearchOptions {
            UseRegularExpressions = true, MatchCase = true
        }).Hits);
    }

    [Fact]
    public void BacktrackingPatternsHonorTheConfiguredExecutionLimit() {
        string text = new string('a', 50_000) + "!";
        var document = Document(text, text);
        Assert.Throws<RegexMatchTimeoutException>(() => document.Search("^(a+)+$", new OfficeDocumentSearchOptions {
            UseRegularExpressions = true, RegularExpressionTimeout = TimeSpan.FromMilliseconds(1)
        }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CancellationStopsLiteralAndRegexSearchBeforeReturningMatches(bool regex) {
        var document = Document("A12 A345", "A12 ", "A345");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var error = Assert.Throws<OperationCanceledException>(() => document.Search(regex ? @"A\d+" : "A12",
            new OfficeDocumentSearchOptions { UseRegularExpressions = regex }, cancellation.Token));
        Assert.Equal(cancellation.Token, error.CancellationToken);
    }

    [Fact]
    public void InvalidPatternsAndUnboundedTimeoutsAreRejected() {
        var document = Document("A1", "A1");
        Assert.ThrowsAny<ArgumentException>(() => document.Search("[", new OfficeDocumentSearchOptions { UseRegularExpressions = true }));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.Search("A", new OfficeDocumentSearchOptions {
            UseRegularExpressions = true, RegularExpressionTimeout = Regex.InfiniteMatchTimeout
        }));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.Search("A", new OfficeDocumentSearchOptions {
            UseRegularExpressions = true, RegularExpressionTimeout = TimeSpan.Zero
        }));
    }

    private static OfficeDocumentReadResult Document(string source, params string[] fragments) => new() {
        Kind = ReaderInputKind.Word,
        Blocks = new[] { Block(source, null) },
        Pages = fragments.Select((text, index) => new OfficeDocumentPage {
            Number = index + 1, Location = new ReaderLocation { Page = index + 1 }, Blocks = new[] { Block(text, index + 1) }
        }).ToArray()
    };

    private static OfficeDocumentBlock Block(string text, int? page) => new() {
        Id = "source", Kind = "paragraph", Text = text,
        Location = new ReaderLocation { BlockAnchor = "source", Page = page }
    };
}
