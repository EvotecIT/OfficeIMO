using System.Linq;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Native_Block_Source_Field_Accessor_Tests {
    [Fact]
    public void NativeBlock_SourceFieldAccessors_Address_Repeated_Fields() {
        const string markdown = """
> first
> second
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var quote = Assert.IsType<MarkdownNativeQuoteBlock>(Assert.Single(native.Blocks));

        var markers = quote.EnumerateSourceFields("quoteMarker").ToArray();
        Assert.Equal(2, markers.Length);
        Assert.Equal(new[] { 0, 1 }, markers.Select(marker => marker.Index).ToArray());
        Assert.All(markers, marker => Assert.Same(quote, marker.Block));

        var firstMarker = quote.FindSourceField("quoteMarker", 0);
        var secondMarker = quote.FindSourceField("quoteMarker", 1);
        Assert.NotNull(firstMarker);
        Assert.NotNull(secondMarker);
        Assert.Equal(1, firstMarker!.SourceSpan.StartLine);
        Assert.Equal(2, secondMarker!.SourceSpan.StartLine);
        Assert.Null(quote.FindSourceField("quoteMarker", 2));
        Assert.Null(quote.FindSourceField("missing"));

        Assert.True(native.TryCreateSourceSlice(secondMarker, out var slice));
        Assert.Equal(">", slice.Text);
    }

    [Fact]
    public void Snapshot_SourceFieldAccessors_Address_Repeated_Fields() {
        const string markdown = """
> first
> second
""";

        var snapshot = Assert.Single(MarkdownNativeDocument.Parse(markdown).ToSnapshot().Blocks);

        var markers = snapshot.EnumerateSourceFields("QUOTEMARKER").ToArray();
        Assert.Equal(2, markers.Length);
        Assert.Equal(new[] { 0, 1 }, markers.Select(marker => marker.Index).ToArray());

        var firstMarker = snapshot.FindSourceField("quoteMarker", 0);
        var secondMarker = snapshot.FindSourceField("quoteMarker", 1);
        Assert.NotNull(firstMarker);
        Assert.NotNull(secondMarker);
        Assert.Equal(1, firstMarker!.SourceSpan.StartLine);
        Assert.Equal(2, secondMarker!.SourceSpan.StartLine);
        Assert.Null(snapshot.FindSourceField("quoteMarker", 2));
        Assert.Null(snapshot.FindSourceField("missing"));
    }
}
