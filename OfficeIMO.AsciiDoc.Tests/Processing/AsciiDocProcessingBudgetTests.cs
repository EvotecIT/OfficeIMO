using System.Threading;
using OfficeIMO.AsciiDoc;
using Xunit;

namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocProcessingBudgetTests {
    [Theory]
    [InlineData("disabled")]
    [InlineData("denied")]
    [InlineData("depth")]
    [InlineData("cycle")]
    public void VisibleIncludeFallbacksCannotExceedOutputBudget(string mode) {
        var options = new AsciiDocProcessorOptions { MaximumOutputLength = 10, SourceName = "root.adoc" };
        if (mode != "disabled") options.IncludeResolver = new Resolver(mode);
        if (mode == "depth") options.MaximumIncludeDepth = 0;
        Assert.Throws<InvalidDataException>(() => AsciiDocProcessor.Process("include::x[]\n", options));
    }

    [Fact]
    public void IncludeRequestsCarryRemainingBudgetAndCancellation() {
        using var cancellation = new CancellationTokenSource();
        var resolver = new Resolver("content");
        var options = new AsciiDocProcessorOptions { IncludeResolver = resolver, MaximumIncludedCharacters = 10 };
        AsciiDocProcessor.Process("include::one[]\ninclude::two[]\n", options, cancellation.Token);
        Assert.Equal(new[] { 10, 6 }, resolver.Requests.Select(request => request.MaximumContentLength));
        Assert.All(resolver.Requests, request => Assert.Equal(cancellation.Token, request.CancellationToken));
    }

    [Fact]
    public void RootedResolverBoundsFileReadingBeforeMaterializingInclude() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-include-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        File.WriteAllText(Path.Combine(directory, "part.adoc"), new string('a', 10000));
        try {
            Assert.Throws<InvalidDataException>(() => AsciiDocProcessor.Process("include::part.adoc[]\n",
                new AsciiDocProcessorOptions { IncludeResolver = new AsciiDocRootedFileIncludeResolver(directory), MaximumIncludedCharacters = 10 }));
        } finally {
            Directory.Delete(directory, true);
        }
    }

    private sealed class Resolver : IAsciiDocIncludeResolver {
        private readonly string _mode;
        internal Resolver(string mode) => _mode = mode;
        internal List<AsciiDocIncludeRequest> Requests { get; } = new List<AsciiDocIncludeRequest>();
        public AsciiDocIncludeResult? Resolve(AsciiDocIncludeRequest request) {
            Requests.Add(request);
            return _mode switch {
                "denied" => null,
                "cycle" => new AsciiDocIncludeResult("Body", "root.adoc"),
                _ => new AsciiDocIncludeResult("Body")
            };
        }
    }
}
