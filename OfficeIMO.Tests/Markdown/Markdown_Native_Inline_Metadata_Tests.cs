using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Native_Inline_Metadata_Tests {
    [Fact]
    public void Formatting_Marker_Metadata_Is_Source_Addressable_In_Nested_Native_Inlines_And_Snapshots() {
        const string markdown = "Start **bold and _em_** end\n";

        var native = MarkdownNativeDocument.Parse(markdown);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var strong = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Strong);
        var emphasis = Assert.Single(strong.Children, inline => inline.Kind == MarkdownNativeInlineKind.Emphasis);

        var strongOpening = Assert.Single(strong.Metadata, metadata => metadata.Name == "openingMarker");
        var strongClosing = Assert.Single(strong.Metadata, metadata => metadata.Name == "closingMarker");
        var emphasisOpening = Assert.Single(emphasis.Metadata, metadata => metadata.Name == "openingMarker");
        var emphasisClosing = Assert.Single(emphasis.Metadata, metadata => metadata.Name == "closingMarker");

        Assert.Equal("**", strongOpening.Value);
        Assert.Equal("**", strongClosing.Value);
        Assert.Equal("_", emphasisOpening.Value);
        Assert.Equal("_", emphasisClosing.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 8), strongOpening.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 22, 1, 23), strongClosing.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 18, 1, 18), emphasisOpening.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 21, 1, 21), emphasisClosing.SourceSpan);

        var edited = native.CreateReplaceEdit(emphasisClosing, "*").Apply(native.SourceMarkdown);
        edited = native.CreateReplaceEdit(emphasisOpening, "*").Apply(edited);
        Assert.Equal("Start **bold and *em*** end\n", edited);

        var snapshotStrong = Assert.Single(native.ToSnapshot().Blocks[0].Inlines, inline => inline.Kind == MarkdownNativeInlineKind.Strong);
        var snapshotEmphasis = Assert.Single(snapshotStrong.Children, inline => inline.Kind == MarkdownNativeInlineKind.Emphasis);

        Assert.Equal("**", snapshotStrong.Metadata["openingMarker"]);
        Assert.Equal("**", snapshotStrong.Metadata["closingMarker"]);
        Assert.Equal("_", snapshotEmphasis.Metadata["openingMarker"]);
        Assert.Equal("_", snapshotEmphasis.Metadata["closingMarker"]);
        Assert.Equal(7, snapshotStrong.MetadataSourceSpans["openingMarker"]!.StartColumn);
        Assert.Equal(23, snapshotStrong.MetadataSourceSpans["closingMarker"]!.EndColumn);
        Assert.Equal(18, snapshotEmphasis.MetadataSourceSpans["openingMarker"]!.StartColumn);
        Assert.Equal(21, snapshotEmphasis.MetadataSourceSpans["closingMarker"]!.EndColumn);
    }

    [Fact]
    public void Code_Span_Marker_Metadata_Is_Source_Addressable_In_Native_Projection_And_Snapshots() {
        const string markdown = "Use ``code ` tick`` now\n";

        var native = MarkdownNativeDocument.Parse(markdown);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var code = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Code);
        var opening = Assert.Single(code.Metadata, metadata => metadata.Name == "openingMarker");
        var closing = Assert.Single(code.Metadata, metadata => metadata.Name == "closingMarker");

        Assert.Equal("code ` tick", code.Text);
        Assert.Equal("``", opening.Value);
        Assert.Equal("``", closing.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 6), opening.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 18, 1, 19), closing.SourceSpan);

        var edited = native.CreateReplaceEdit(closing, "```").Apply(native.SourceMarkdown);
        edited = native.CreateReplaceEdit(opening, "```").Apply(edited);
        Assert.Equal("Use ```code ` tick``` now\n", edited);

        var snapshotCode = Assert.Single(native.ToSnapshot().Blocks[0].Inlines, inline => inline.Kind == MarkdownNativeInlineKind.Code);
        Assert.Equal("``", snapshotCode.Metadata["openingMarker"]);
        Assert.Equal("``", snapshotCode.Metadata["closingMarker"]);
        Assert.Equal(5, snapshotCode.MetadataSourceSpans["openingMarker"]!.StartColumn);
        Assert.Equal(6, snapshotCode.MetadataSourceSpans["openingMarker"]!.EndColumn);
        Assert.Equal(18, snapshotCode.MetadataSourceSpans["closingMarker"]!.StartColumn);
        Assert.Equal(19, snapshotCode.MetadataSourceSpans["closingMarker"]!.EndColumn);
    }

    [Fact]
    public void Inline_Html_Tag_Marker_Metadata_And_Nested_Inlines_Are_Source_Addressable() {
        const string markdown = "Use <u>under **bold**</u> now\n";

        var native = MarkdownNativeDocument.Parse(markdown);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var htmlTag = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.HtmlTag);
        var strong = Assert.Single(htmlTag.Children, inline => inline.Kind == MarkdownNativeInlineKind.Strong);
        var opening = Assert.Single(htmlTag.Metadata, metadata => metadata.Name == "openingMarker");
        var closing = Assert.Single(htmlTag.Metadata, metadata => metadata.Name == "closingMarker");
        var strongOpening = Assert.Single(strong.Metadata, metadata => metadata.Name == "openingMarker");
        var strongClosing = Assert.Single(strong.Metadata, metadata => metadata.Name == "closingMarker");

        Assert.Equal("under bold", htmlTag.Text);
        Assert.Equal("<u>under **bold**</u>", htmlTag.Markdown);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 25), htmlTag.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 16, 1, 19), strong.SourceSpan);
        Assert.Equal("<u>", opening.Value);
        Assert.Equal("</u>", closing.Value);
        Assert.Equal("**", strongOpening.Value);
        Assert.Equal("**", strongClosing.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 7), opening.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 22, 1, 25), closing.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 14, 1, 15), strongOpening.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 20, 1, 21), strongClosing.SourceSpan);

        var edited = native.CreateReplaceEdit(closing, "</ins>").Apply(native.SourceMarkdown);
        edited = native.CreateReplaceEdit(opening, "<ins>").Apply(edited);
        Assert.Equal("Use <ins>under **bold**</ins> now\n", edited);

        var snapshotHtmlTag = Assert.Single(native.ToSnapshot().Blocks[0].Inlines, inline => inline.Kind == MarkdownNativeInlineKind.HtmlTag);
        var snapshotStrong = Assert.Single(snapshotHtmlTag.Children, inline => inline.Kind == MarkdownNativeInlineKind.Strong);

        Assert.Equal("<u>", snapshotHtmlTag.Metadata["openingMarker"]);
        Assert.Equal("</u>", snapshotHtmlTag.Metadata["closingMarker"]);
        Assert.Equal("**", snapshotStrong.Metadata["openingMarker"]);
        Assert.Equal("**", snapshotStrong.Metadata["closingMarker"]);
        Assert.Equal(5, snapshotHtmlTag.MetadataSourceSpans["openingMarker"]!.StartColumn);
        Assert.Equal(25, snapshotHtmlTag.MetadataSourceSpans["closingMarker"]!.EndColumn);
        Assert.Equal(14, snapshotStrong.MetadataSourceSpans["openingMarker"]!.StartColumn);
        Assert.Equal(21, snapshotStrong.MetadataSourceSpans["closingMarker"]!.EndColumn);
    }

    [Fact]
    public void Linked_Image_Metadata_Is_Source_Addressable_In_Native_Projection_And_Snapshots() {
        const string markdown = "Paragraph [![Alt](img.png \"Img\")](https://example.com \"Link title\").";

        var native = MarkdownNativeDocument.Parse(markdown);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var imageLink = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.ImageLink);

        Assert.Equal("Alt", imageLink.Text);
        Assert.Equal(new MarkdownSourceSpan(1, 11, 1, 67), imageLink.SourceSpan);
        Assert.Equal("Alt", imageLink.GetMetadata("alt"));
        Assert.Equal("img.png", imageLink.GetMetadata("source"));
        Assert.Equal("Img", imageLink.GetMetadata("imageTitle"));
        Assert.Equal("https://example.com", imageLink.GetMetadata("target"));
        Assert.Equal("Link title", imageLink.GetMetadata("title"));

        var alt = Assert.Single(imageLink.Metadata, metadata => metadata.Name == "alt");
        var source = Assert.Single(imageLink.Metadata, metadata => metadata.Name == "source");
        var imageTitle = Assert.Single(imageLink.Metadata, metadata => metadata.Name == "imageTitle");
        var target = Assert.Single(imageLink.Metadata, metadata => metadata.Name == "target");
        var title = Assert.Single(imageLink.Metadata, metadata => metadata.Name == "title");

        Assert.Equal(new MarkdownSourceSpan(1, 14, 1, 16), alt.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 19, 1, 25), source.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 28, 1, 30), imageTitle.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 35, 1, 53), target.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 56, 1, 65), title.SourceSpan);

        Assert.Equal(
            "Paragraph [![Logo](img.png \"Img\")](https://example.com \"Link title\").",
            native.CreateReplaceEdit(alt, "Logo").Apply(native.SourceMarkdown));
        Assert.Equal(
            "Paragraph [![Alt](logo.svg \"Img\")](https://example.com \"Link title\").",
            native.CreateReplaceEdit(source, "logo.svg").Apply(native.SourceMarkdown));
        Assert.Equal(
            "Paragraph [![Alt](img.png \"Diagram\")](https://example.com \"Link title\").",
            native.CreateReplaceEdit(imageTitle, "Diagram").Apply(native.SourceMarkdown));
        Assert.Equal(
            "Paragraph [![Alt](img.png \"Img\")](https://contoso.test \"Link title\").",
            native.CreateReplaceEdit(target, "https://contoso.test").Apply(native.SourceMarkdown));
        Assert.Equal(
            "Paragraph [![Alt](img.png \"Img\")](https://example.com \"Docs\").",
            native.CreateReplaceEdit(title, "Docs").Apply(native.SourceMarkdown));

        var snapshot = native.ToSnapshot();
        var snapshotParagraph = Assert.Single(snapshot.Blocks);
        var snapshotImageLink = Assert.Single(snapshotParagraph.Inlines, inline => inline.Kind == MarkdownNativeInlineKind.ImageLink);

        Assert.Equal("Alt", snapshotImageLink.Metadata["alt"]);
        Assert.Equal("img.png", snapshotImageLink.Metadata["source"]);
        Assert.Equal("Img", snapshotImageLink.Metadata["imageTitle"]);
        Assert.Equal("https://example.com", snapshotImageLink.Metadata["target"]);
        Assert.Equal("Link title", snapshotImageLink.Metadata["title"]);
        Assert.Equal(14, snapshotImageLink.MetadataSourceSpans["alt"]!.StartColumn);
        Assert.Equal(16, snapshotImageLink.MetadataSourceSpans["alt"]!.EndColumn);
        Assert.Equal(19, snapshotImageLink.MetadataSourceSpans["source"]!.StartColumn);
        Assert.Equal(25, snapshotImageLink.MetadataSourceSpans["source"]!.EndColumn);
        Assert.Equal(28, snapshotImageLink.MetadataSourceSpans["imageTitle"]!.StartColumn);
        Assert.Equal(30, snapshotImageLink.MetadataSourceSpans["imageTitle"]!.EndColumn);
        Assert.Equal(35, snapshotImageLink.MetadataSourceSpans["target"]!.StartColumn);
        Assert.Equal(53, snapshotImageLink.MetadataSourceSpans["target"]!.EndColumn);
        Assert.Equal(56, snapshotImageLink.MetadataSourceSpans["title"]!.StartColumn);
        Assert.Equal(65, snapshotImageLink.MetadataSourceSpans["title"]!.EndColumn);
    }
}
