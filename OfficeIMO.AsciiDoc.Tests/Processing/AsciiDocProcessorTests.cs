namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocProcessorTests {
    [Fact]
    public void Conditionals_AreExplicitWhileOriginalDocumentStaysLossless() {
        const string source =
            ":edition: pro\n" +
            "ifdef::edition[]\nVisible\nendif::[]\n" +
            "ifndef::edition[]\nHidden\nendif::[]\n" +
            "ifdef::edition[Inline {edition}]\n" +
            "ifeval::[\"2\" >= \"1\"]\nEvaluation passed\nendif::[]\n";

        AsciiDocProcessingResult result = AsciiDocProcessor.Process(source, new AsciiDocProcessorOptions { SourceName = "main.adoc" });

        Assert.Equal(source, result.SourceDocument.ToAsciiDoc());
        Assert.Equal(":edition: pro\nVisible\nInline pro\nEvaluation passed\n", result.ProcessedSource);
        Assert.Equal(result.ProcessedSource, result.Document.ToAsciiDoc());
        Assert.Equal("pro", result.Attributes.GetValueOrDefault("edition"));
        Assert.False(result.HasErrors);
    }

    [Fact]
    public void Includes_AreResolverControlledNestedSelectedAndAttributeAware() {
        const string source = ":dir: parts\ninclude::{dir}/chapter.adoc[tag=body,leveloffset=+1]\n";
        var resolver = new DictionaryResolver(new Dictionary<string, string> {
            ["parts/chapter.adoc"] =
                "// tag::body[]\n== Included\nifdef::extra[]\nExtra\nendif::[]\ninclude::nested.adoc[lines=2..3]\n// end::body[]\n",
            ["nested.adoc"] = "skip\nNested one\nNested two\nskip\n"
        });

        AsciiDocProcessingResult result = AsciiDocProcessor.Process(source, new AsciiDocProcessorOptions {
            SourceName = "main.adoc",
            IncludeResolver = resolver,
            Attributes = new Dictionary<string, string> { ["extra"] = string.Empty }
        });

        Assert.Equal(":dir: parts\n=== Included\nExtra\nNested one\nNested two\n", result.ProcessedSource);
        Assert.Equal(2, resolver.Requests.Count);
        Assert.Equal("parts/chapter.adoc", resolver.Requests[0].Target);
        Assert.Equal("parts/chapter.adoc", resolver.Requests[1].CurrentSourceName);
        Assert.False(result.HasErrors);
    }

    [Fact]
    public void DisabledAndMissingIncludes_AreVisibleAndDiagnosedUnlessOptional() {
        const string source = "include::missing.adoc[]\ninclude::optional.adoc[opts=optional]\n";

        AsciiDocProcessingResult result = AsciiDocProcessor.Process(source);

        Assert.Equal(source, result.ProcessedSource);
        Assert.Single(result.Diagnostics);
        Assert.Equal("ADOCPROC001", result.Diagnostics[0].Code);
    }

    [Fact]
    public void IncludeCycle_IsDiagnosedAndDirectiveRemainsVisible() {
        var resolver = new DictionaryResolver(new Dictionary<string, string> {
            ["a.adoc"] = "include::a.adoc[]\n"
        });

        AsciiDocProcessingResult result = AsciiDocProcessor.Process("include::a.adoc[]\n", new AsciiDocProcessorOptions {
            SourceName = "main.adoc",
            IncludeResolver = resolver
        });

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "ADOCPROC004");
        Assert.Equal("include::a.adoc[]\n", result.ProcessedSource);
        Assert.True(result.HasErrors);
    }

    [Fact]
    public void RootedFileResolver_DeniesTraversalAbsoluteAndUriTargets() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-adoc-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            File.WriteAllText(Path.Combine(root, "allowed.adoc"), "Allowed\n");
            var resolver = new AsciiDocRootedFileIncludeResolver(root);
            var attributes = new AsciiDocDocumentAttributes(new Dictionary<string, string>());

            Assert.NotNull(resolver.Resolve(new AsciiDocIncludeRequest("allowed.adoc", null, 1, attributes)));
            Assert.Null(resolver.Resolve(new AsciiDocIncludeRequest("../outside.adoc", null, 1, attributes)));
            Assert.Null(resolver.Resolve(new AsciiDocIncludeRequest(Path.Combine(root, "allowed.adoc"), null, 1, attributes)));
            Assert.Null(resolver.Resolve(new AsciiDocIncludeRequest("https://example.test/file.adoc", null, 1, attributes)));
        } finally {
            Directory.Delete(root, true);
        }
    }

    [Fact]
    public void RootedFileResolver_PreservesFilesystemRootDirectories() {
        string file = Path.Combine(Path.GetTempPath(), "officeimo-adoc-" + Guid.NewGuid().ToString("N") + ".adoc");
        File.WriteAllText(file, "From root\n");
        try {
            string root = Path.GetPathRoot(file)!;
            string relative = file.Substring(root.Length);
            var resolver = new AsciiDocRootedFileIncludeResolver(root) { AllowSymbolicLinks = true };
            var attributes = new AsciiDocDocumentAttributes(new Dictionary<string, string>());

            AsciiDocIncludeResult? result = resolver.Resolve(new AsciiDocIncludeRequest(relative, null, 1, attributes));

            Assert.NotNull(result);
            Assert.Equal("From root\n", result!.Content);
        } finally {
            File.Delete(file);
        }
    }

    [Fact]
    public void IncludeTagFilters_SupportDoubleWildcardAndNamedExclusions() {
        const string included =
            "outside\n" +
            "# tag::public[]\npublic\n# end::public[]\n" +
            "<!-- tag::secret[] -->\nsecret\n<!-- end::secret[] -->\n";
        var resolver = new DictionaryResolver(new Dictionary<string, string> { ["part.adoc"] = included });

        AsciiDocProcessingResult result = AsciiDocProcessor.Process(
            "include::part.adoc[tags=**;!secret]\n",
            new AsciiDocProcessorOptions { SourceName = "main.adoc", IncludeResolver = resolver });

        Assert.Equal("outside\npublic\n", result.ProcessedSource);
        Assert.False(result.HasErrors);
    }

    [Fact]
    public void IncludeTagMarkers_RequireAWordBoundaryAfterTheDirective() {
        const string included =
            "tag::fake[]suffix\n" +
            "// tag::real[]\nselected\n// end::real[]\n";
        var resolver = new DictionaryResolver(new Dictionary<string, string> { ["part.adoc"] = included });

        AsciiDocProcessingResult result = AsciiDocProcessor.Process(
            "include::part.adoc[tags=**]\n",
            new AsciiDocProcessorOptions { SourceName = "main.adoc", IncludeResolver = resolver });

        Assert.Equal("tag::fake[]suffix\nselected\n", result.ProcessedSource);
    }

    [Theory]
    [InlineData("foo;!*", "foo one\nfoo two\n")]
    [InlineData("!*;foo", "outside\nfoo one\nfoo two\noutside end\n")]
    [InlineData("*;!bar", "foo one\nfoo two\n")]
    public void IncludeTagFilters_RespectWildcardBaseAndNestedExclusions(string filter, string expected) {
        const string included =
            "outside\n" +
            "// tag::foo[]\nfoo one\n" +
            "// tag::bar[]\nbar\n// end::bar[]\n" +
            "foo two\n// end::foo[]\n" +
            "outside end\n";
        var resolver = new DictionaryResolver(new Dictionary<string, string> { ["part.adoc"] = included });

        AsciiDocProcessingResult result = AsciiDocProcessor.Process(
            "include::part.adoc[tags=" + filter + "]\n",
            new AsciiDocProcessorOptions { SourceName = "main.adoc", IncludeResolver = resolver });

        Assert.Equal(expected, result.ProcessedSource);
    }

    private sealed class DictionaryResolver : IAsciiDocIncludeResolver {
        private readonly IReadOnlyDictionary<string, string> _content;

        internal DictionaryResolver(IReadOnlyDictionary<string, string> content) {
            _content = content;
        }

        internal List<AsciiDocIncludeRequest> Requests { get; } = new List<AsciiDocIncludeRequest>();

        public AsciiDocIncludeResult? Resolve(AsciiDocIncludeRequest request) {
            Requests.Add(request);
            return _content.TryGetValue(request.Target, out string? content)
                ? new AsciiDocIncludeResult(content, request.Target)
                : null;
        }
    }
}
