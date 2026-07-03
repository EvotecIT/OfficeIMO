using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Native_Inline_Source_Slice_Tests {
    [Fact]
    public void NativeInline_SourceSlice_Uses_Content_Span_For_Formatting_Inlines() {
        var native = MarkdownNativeDocument.Parse("A **strong** word");
        var strong = Assert.Single(native.EnumerateInlines(), inline => inline.Kind == MarkdownNativeInlineKind.Strong);

        var created = native.TryCreateSourceSlice(strong, out var slice);

        Assert.True(created);
        Assert.Equal(MarkdownSourceTextKind.Normalized, slice.TextKind);
        Assert.Equal("strong", slice.Text);
    }

    [Fact]
    public void NativeInlineMetadata_SourceSlice_Can_Address_Link_Target_And_Title() {
        var native = MarkdownNativeDocument.Parse("""See [docs](https://example.com "Docs title").""");
        var target = Assert.Single(native.EnumerateInlineMetadata("target"));
        var title = Assert.Single(native.EnumerateInlineMetadata("title"));

        Assert.True(native.TryCreateSourceSlice(target, out var targetSlice));
        Assert.True(native.TryCreateSourceSlice(title, out var titleSlice));

        Assert.Equal(MarkdownSourceTextKind.Normalized, targetSlice.TextKind);
        Assert.Equal(MarkdownSourceTextKind.Normalized, titleSlice.TextKind);
        Assert.Equal("https://example.com", targetSlice.Text);
        Assert.Equal("Docs title", titleSlice.Text);
    }

    [Fact]
    public void NativeInlineMetadata_OriginalSourceSlice_Can_Address_LineEnding_Equivalent_Input() {
        var markdown = "See [docs](https://example.com \"Docs title\").\r\n";
        var options = new MarkdownReaderOptions {
            PreserveTrivia = true
        };
        var native = MarkdownNativeDocument.Parse(markdown, options);
        var target = Assert.Single(native.EnumerateInlineMetadata("target"));
        var title = Assert.Single(native.EnumerateInlineMetadata("title"));

        Assert.True(native.TryCreateOriginalSourceSlice(target, out var targetSlice, out var targetReason));
        Assert.True(native.TryCreateOriginalSourceSlice(title, out var titleSlice, out var titleReason));

        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.None, targetReason);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.None, titleReason);
        Assert.Equal(MarkdownSourceTextKind.Original, targetSlice.TextKind);
        Assert.Equal(MarkdownSourceTextKind.Original, titleSlice.TextKind);
        Assert.Equal("https://example.com", targetSlice.Text);
        Assert.Equal("Docs title", titleSlice.Text);
    }

    [Fact]
    public void NativeInlineMetadata_OriginalSourceSlice_Returns_NotPreserved_Reason_When_Trivia_Is_Disabled() {
        var native = MarkdownNativeDocument.Parse("""See [docs](https://example.com "Docs title").""");
        var target = Assert.Single(native.EnumerateInlineMetadata("target"));

        var created = native.TryCreateOriginalSourceSlice(target, out _, out var failureReason);

        Assert.False(created);
        Assert.Equal(MarkdownOriginalSourceSliceFailureReason.OriginalMarkdownNotPreserved, failureReason);
    }
}
