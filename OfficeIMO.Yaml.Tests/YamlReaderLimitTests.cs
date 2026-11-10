using System.Globalization;
using System.Text;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

// Limits protect callers from hostile input; the suite contains no input large or deep enough to reach them.
public class YamlReaderLimitTests {
    private static void ReadAll(YamlReader reader) {
        while (reader.Read()) { }
    }

    [Theory]
    [InlineData("[", "]")]
    [InlineData("{a: ", "}")]
    [InlineData("- ", "")]
    public void DepthBeyondTheLimitThrowsWithItsPosition(string open, string close) {
        string yaml = open == "- "
            ? string.Concat(Enumerable.Range(0, 20).Select(i => new string(' ', i * 2) + "-\n")) + new string(' ', 40) + "x\n"
            : string.Concat(Enumerable.Repeat(open, 20)) + "x" + string.Concat(Enumerable.Repeat(close, 20));
        ReadAll(new YamlReader(yaml, new YamlReaderOptions { MaxDepth = 25 }));
        YamlLimitExceededException error = Assert.Throws<YamlLimitExceededException>(() => ReadAll(new YamlReader(yaml, new YamlReaderOptions { MaxDepth = 10 })));
        Assert.Equal(nameof(YamlReaderOptions.MaxDepth), error.LimitName);
        Assert.Equal(YamlErrorCodes.LimitExceeded, error.Code);
        Assert.True(error.Line >= 1);
    }

    // A one-line flow nest far deeper than the limit must fail without scanning the whole line or recursing.
    [Fact]
    public void DeepFlowNestingFailsFast() {
        string yaml = new string('[', 1_000_000);
        Assert.Throws<YamlLimitExceededException>(() => ReadAll(new YamlReader(yaml)));
    }

    [Theory]
    [InlineData("key: {0}\n")]
    [InlineData("key: \"{0}\"\n")]
    [InlineData("key: |\n  {0}\n")]
    public void ScalarLongerThanTheLimitThrows(string template) {
        string yaml = string.Format(CultureInfo.InvariantCulture, template, new string('x', 5000));
        ReadAll(new YamlReader(yaml, new YamlReaderOptions { MaxScalarLength = 6000 }));
        YamlLimitExceededException error = Assert.Throws<YamlLimitExceededException>(() => ReadAll(new YamlReader(yaml, new YamlReaderOptions { MaxScalarLength = 1000 })));
        Assert.Equal(nameof(YamlReaderOptions.MaxScalarLength), error.LimitName);
        Assert.Equal(1, error.Line);
    }

    [Fact]
    public void BufferedScalarLongerThanTheLimitStopsGrowingTheWindow() {
        string yaml = "key: " + new string('x', 4_000_000) + "\n";
        var options = new YamlReaderOptions { MaxScalarLength = 10_000 };
        Assert.Throws<YamlLimitExceededException>(() => ReadAll(new YamlReader(new TrickleReader(yaml, 8192), options)));
    }

    [Fact]
    public void InputLargerThanTheLimitThrowsForEveryInputKind() {
        string yaml = "a: " + new string('b', 200) + "\n";
        var options = new YamlReaderOptions { MaxInputBytes = 100 };
        Assert.Throws<YamlLimitExceededException>(() => ReadAll(new YamlReader(yaml, options)));
        Assert.Throws<YamlLimitExceededException>(() => ReadAll(new YamlReader(new StringReader(yaml), options)));
        Assert.Throws<YamlLimitExceededException>(() => ReadAll(new YamlReader(new MemoryStream(Encoding.UTF8.GetBytes(yaml)), options)));
    }

    [Fact]
    public void CancellationStopsALongReadWithoutWaitingForTheEnd() {
        var text = new StringBuilder();
        for (int i = 0; i < 100_000; i++) text.Append("- item ").Append(i).Append('\n');
        using var source = new CancellationTokenSource();
        using var reader = new YamlReader(text.ToString(), cancellationToken: source.Token);
        int events = 0;
        Assert.Throws<OperationCanceledException>(() => {
            while (reader.Read()) {
                if (++events == 100) source.Cancel();
            }
        });
        Assert.True(events < 2000, "Cancellation was observed after " + events + " events.");
    }

    // A file of comments produces no tokens, so the comment loop must check cancellation too.
    [Fact]
    public void CancellationIsObservedInsideLongCommentRuns() {
        string yaml = string.Concat(Enumerable.Repeat("# comment\n", 200_000)) + "a: 1\n";
        using var source = new CancellationTokenSource();
        source.Cancel();
        using var reader = new YamlReader(yaml, cancellationToken: source.Token);
        reader.Read();
        Assert.Throws<OperationCanceledException>(() => reader.Read());
    }
}
