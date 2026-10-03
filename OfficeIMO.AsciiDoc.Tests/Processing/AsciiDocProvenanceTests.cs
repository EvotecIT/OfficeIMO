using OfficeIMO.AsciiDoc;
using Xunit;

namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocProvenanceTests {
    [Fact]
    public void NestedSelectedIncludesRetainOriginalFileLineColumnAndOffset() {
        const string root = "Root\r\ninclude::chapter.adoc[tag=body]\r\nTail";
        const string chapter = "skip\r\n// tag::body[]\r\nChapter\r\ninclude::nested.adoc[lines=2]\r\n// end::body[]\r\n";
        const string nested = "skip\r\nNested text\r\nskip\r\n";
        var result = Process(root, new Dictionary<string, string> { ["chapter.adoc"] = chapter, ["nested.adoc"] = nested });
        Assert.Equal("Root\r\nChapter\r\nNested text\r\nTail", result.ProcessedSource);
        int offset = result.ProcessedSource.IndexOf("text", StringComparison.Ordinal);
        AsciiDocSourceMapping origin = Assert.IsType<AsciiDocSourceMapping>(result.SourceMap.Find(offset));
        Assert.Equal("nested.adoc", origin.SourceName);
        Assert.True(origin.IsExact);
        AsciiDocSourcePosition position = origin.GetSourcePosition(offset);
        Assert.Equal(nested.IndexOf("text", StringComparison.Ordinal), position.Offset);
        Assert.Equal(2, position.Line);
        Assert.Equal(8, position.Column);
        Assert.Equal(root, result.SourceDocument.ToAsciiDoc());
        AsciiDocSourceMapping end = Assert.IsType<AsciiDocSourceMapping>(result.SourceMap.Find(result.ProcessedSource.Length));
        Assert.Equal("root.adoc", end.SourceName);
        Assert.Equal(root.Length, end.GetSourcePosition(result.ProcessedSource.Length).Offset);
        AssertContiguous(result);
    }

    [Fact]
    public void IncludeSelectionDiagnosticsUseOriginalLineNumbers() {
        var result = Process("include::part.adoc[lines=3..4]\n", new Dictionary<string, string> {
            ["part.adoc"] = "skip\nskip\ninclude::missing.adoc[]\n====\n"
        });
        AsciiDocProcessingDiagnostic diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("ADOCPROC003", diagnostic.Code);
        Assert.Equal("part.adoc", diagnostic.SourceName);
        Assert.Equal(3, diagnostic.Line);
        AsciiDocDiagnostic parseDiagnostic = Assert.Single(result.Document.Diagnostics);
        AsciiDocSourceMapping origin = Assert.IsType<AsciiDocSourceMapping>(result.SourceMap.Find(parseDiagnostic.Span.Start.Offset));
        Assert.Equal("part.adoc", origin.SourceName);
        Assert.Equal(4, origin.GetSourcePosition(parseDiagnostic.Span.Start.Offset).Line);
        AssertContiguous(result);
    }

    [Fact]
    public void HeadingOffsetsAndAddedLineEndingsHaveExplicitApproximateOrigins() {
        var result = Process("include::part.adoc[leveloffset=2147483647]\nTail\n", new Dictionary<string, string> {
            ["part.adoc"] = "skip\n== Heading"
        });
        Assert.Equal("skip\n====== Heading\nTail\n", result.ProcessedSource);
        int offset = result.ProcessedSource.IndexOf("Heading", StringComparison.Ordinal);
        AsciiDocSourceMapping heading = Assert.IsType<AsciiDocSourceMapping>(result.SourceMap.Find(offset));
        Assert.False(heading.IsExact);
        Assert.Equal("part.adoc", heading.SourceName);
        Assert.Equal(2, heading.OriginalSpan.Start.Line);
        Assert.Equal(heading.OriginalSpan.Start, heading.GetSourcePosition(offset));
        AsciiDocSourceMapping newline = Assert.IsType<AsciiDocSourceMapping>(result.SourceMap.Find(offset + "Heading".Length));
        Assert.False(newline.IsExact);
        Assert.Equal("root.adoc", newline.SourceName);
        Assert.Equal(1, newline.OriginalSpan.Start.Line);
        AssertContiguous(result);
    }

    [Fact]
    public void GeneratedDirectiveContentMapsToItsProducingRootRange() {
        var extensions = new AsciiDocExtensionRegistry().RegisterDirective("generated", new Replacement());
        const string root = ":enabled:\nifdef::enabled[Inline text]\ngenerated::[]\nTail\n";
        var result = AsciiDocProcessor.Process(root, new AsciiDocProcessorOptions { SourceName = "root.adoc", Extensions = extensions });
        Assert.Equal(":enabled:\nInline text\nFirst\nSecond\nTail\n", result.ProcessedSource);
        int offset = result.ProcessedSource.IndexOf("Second", StringComparison.Ordinal);
        AsciiDocSourceMapping origin = Assert.IsType<AsciiDocSourceMapping>(result.SourceMap.Find(offset));
        Assert.False(origin.IsExact);
        Assert.Equal(3, origin.GetSourcePosition(offset).Line);
        Assert.Equal("root.adoc", origin.SourceName);
        AsciiDocSourceMapping inline = Assert.IsType<AsciiDocSourceMapping>(result.SourceMap.Find(result.ProcessedSource.IndexOf("Inline", StringComparison.Ordinal)));
        Assert.False(inline.IsExact);
        Assert.Equal(2, inline.OriginalSpan.Start.Line);
        AssertContiguous(result);
    }

    [Fact]
    public void EmptyProcessedSourcesHaveNoOriginRange() {
        var result = AsciiDocProcessor.Process("ifdef::missing[]\nHidden\nendif::[]\n");
        Assert.Empty(result.ProcessedSource);
        Assert.Empty(result.SourceMap.Entries);
        Assert.Null(result.SourceMap.Find(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => result.SourceMap.Find(1));
    }

    private static AsciiDocProcessingResult Process(string root, IDictionary<string, string> sources) =>
        AsciiDocProcessor.Process(root, new AsciiDocProcessorOptions { SourceName = "root.adoc", IncludeResolver = new Resolver(sources) });
    private static void AssertContiguous(AsciiDocProcessingResult result) {
        int offset = 0;
        foreach (AsciiDocSourceMapping entry in result.SourceMap.Entries) { Assert.Equal(offset, entry.ProcessedStart); offset = entry.ProcessedEnd; }
        Assert.Equal(result.ProcessedSource.Length, offset);
    }
    private sealed class Resolver : IAsciiDocIncludeResolver {
        private readonly IDictionary<string, string> _sources;
        internal Resolver(IDictionary<string, string> sources) { _sources = sources; }
        public AsciiDocIncludeResult? Resolve(AsciiDocIncludeRequest request) => _sources.TryGetValue(request.Target, out string? text) ? new AsciiDocIncludeResult(text, request.Target) : null;
    }
    private sealed class Replacement : IAsciiDocDirectiveProcessor {
        public AsciiDocDirectiveResult Process(AsciiDocDirectiveContext context) => AsciiDocDirectiveResult.Replace("First\nSecond\n");
    }
}
