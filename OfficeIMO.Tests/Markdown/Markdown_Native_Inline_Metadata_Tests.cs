using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Native_Inline_Metadata_Tests {
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
