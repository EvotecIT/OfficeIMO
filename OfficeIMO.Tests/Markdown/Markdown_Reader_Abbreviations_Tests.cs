using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Reader_Abbreviations_Tests {
    [Fact]
    public void Abbreviations_Are_Opt_In() {
        const string markdown = "*[HTML]: Hyper Text Markup Language\nHTML";

        var defaultHtml = MarkdownReader.Parse(markdown).ToHtmlFragment(CreatePlainHtmlOptions());
        Assert.DoesNotContain("<abbr", defaultHtml, StringComparison.Ordinal);

        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.Abbreviations = true;

        var html = MarkdownReader.Parse(markdown, options).ToHtmlFragment(CreatePlainHtmlOptions());
        Assert.Contains("<abbr title=\"Hyper Text Markup Language\">HTML</abbr>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Abbreviation_Definitions_Are_Collected_And_Consumed() {
        const string markdown = "*[HTML]: Hyper Text Markup Language\nHTML";
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.Abbreviations = true;

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        var definition = Assert.Single(result.AbbreviationDefinitions);
        Assert.Equal("HTML", definition.Label);
        Assert.Equal("Hyper Text Markup Language", definition.Title);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 6), definition.LabelSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 10, 1, 35), definition.TitleSourceSpan);

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(result.Document.Blocks));
        var abbreviation = Assert.IsType<AbbreviationInline>(Assert.Single(paragraph.Inlines.Nodes));
        Assert.Equal("HTML", abbreviation.Text);
        Assert.Equal("Hyper Text Markup Language", abbreviation.Title);
        Assert.Equal("HTML", result.Document.ToMarkdown().TrimEnd('\r', '\n'));
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Abbreviation_Definition_And_Inline_Metadata() {
        const string markdown = "*[HTML]: Hyper Text Markup Language\nHTML";
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.Abbreviations = true;

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        Assert.Collection(result.SyntaxTree.Children,
            definition => {
                Assert.Equal(MarkdownSyntaxKind.AbbreviationDefinition, definition.Kind);
                Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 35), definition.SourceSpan);
                Assert.Collection(definition.Children,
                    opening => Assert.Equal(MarkdownSyntaxKind.AbbreviationOpeningMarker, opening.Kind),
                    label => {
                        Assert.Equal(MarkdownSyntaxKind.AbbreviationLabel, label.Kind);
                        Assert.Equal("HTML", label.Literal);
                        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 6), label.SourceSpan);
                    },
                    separator => Assert.Equal(MarkdownSyntaxKind.AbbreviationSeparatorMarker, separator.Kind),
                    title => {
                        Assert.Equal(MarkdownSyntaxKind.AbbreviationTitle, title.Kind);
                        Assert.Equal("Hyper Text Markup Language", title.Literal);
                        Assert.Equal(new MarkdownSourceSpan(1, 10, 1, 35), title.SourceSpan);
                    });
            },
            paragraph => {
                Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraph.Kind);
                var inline = Assert.Single(paragraph.Children);
                Assert.Equal(MarkdownSyntaxKind.InlineAbbreviation, inline.Kind);
                Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 4), inline.SourceSpan);
                Assert.Collection(inline.Children,
                    text => {
                        Assert.Equal(MarkdownSyntaxKind.InlineAbbreviationText, text.Kind);
                        Assert.Equal("HTML", text.Literal);
                        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 4), text.SourceSpan);
                    },
                    title => {
                        Assert.Equal(MarkdownSyntaxKind.InlineAbbreviationTitle, title.Kind);
                        Assert.Equal("Hyper Text Markup Language", title.Literal);
                        Assert.Equal(new MarkdownSourceSpan(1, 10, 1, 35), title.SourceSpan);
                    });
            });
    }

    [Fact]
    public void Native_Abbreviation_Metadata_Preserves_Text_And_Title_Source_For_Edits_And_Snapshots() {
        const string markdown = "*[HTML]: Hyper Text Markup Language\nHTML";
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.Abbreviations = true;

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var abbreviation = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Abbreviation);
        var text = Assert.Single(abbreviation.Metadata, metadata => metadata.Name == "text");
        var title = Assert.Single(abbreviation.Metadata, metadata => metadata.Name == "title");

        Assert.Equal("HTML", abbreviation.Text);
        Assert.Equal("HTML", text.Value);
        Assert.Equal("Hyper Text Markup Language", title.Value);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 4), abbreviation.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 4), text.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 10, 1, 35), title.SourceSpan);
        Assert.Equal("*[HTML]: Hypertext Markup Language\nHTML", native.CreateReplaceEdit(title, "Hypertext Markup Language").Apply(native.SourceMarkdown));
        Assert.Equal("*[HTML]: Hyper Text Markup Language\nHTMX", native.CreateReplaceEdit(text, "HTMX").Apply(native.SourceMarkdown));

        var snapshot = Assert.Single(native.ToSnapshot().Blocks[0].Inlines, inline => inline.Kind == MarkdownNativeInlineKind.Abbreviation);
        Assert.Equal("HTML", snapshot.Metadata["text"]);
        Assert.Equal("Hyper Text Markup Language", snapshot.Metadata["title"]);
        Assert.Equal(2, snapshot.MetadataSourceSpans["text"]!.StartLine);
        Assert.Equal(10, snapshot.MetadataSourceSpans["title"]!.StartColumn);
    }

    private static HtmlOptions CreatePlainHtmlOptions() =>
        new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        };
}
