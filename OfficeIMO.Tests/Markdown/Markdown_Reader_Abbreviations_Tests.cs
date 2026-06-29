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
        Assert.Equal("*[HTML]: Hyper Text Markup Language\n\nHTML", NormalizeMarkdown(result.Document.ToMarkdown()));
    }

    [Fact]
    public void Abbreviation_Definitions_Are_Written_For_Reparse_Stability() {
        const string markdown = "*[HTML]: Hyper Text Markup Language\n\nHTML and HTML.";
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.Abbreviations = true;

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        var written = result.Document.ToMarkdown(new MarkdownWriteOptions { OutputLineEnding = "\n" }).TrimEnd('\n');

        Assert.Equal("*[HTML]: Hyper Text Markup Language\n\nHTML and HTML.", written);

        var reparsedHtml = MarkdownReader.Parse(written, options).ToHtmlFragment(CreatePlainHtmlOptions());
        Assert.Contains(
            "<abbr title=\"Hyper Text Markup Language\">HTML</abbr> and <abbr title=\"Hyper Text Markup Language\">HTML</abbr>",
            reparsedHtml,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Abbreviation_Definitions_Can_Have_Empty_Title_With_Source_And_Writer_Proof() {
        const string markdown = "*[HTML]:   \n\nHTML";
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.Abbreviations = true;

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        var definition = Assert.Single(result.AbbreviationDefinitions);
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(result.Document.Blocks));
        var abbreviation = Assert.IsType<AbbreviationInline>(Assert.Single(paragraph.Inlines.Nodes));
        var html = result.Document.ToHtmlFragment(CreatePlainHtmlOptions());
        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeAbbreviation = Assert.Single(native.EnumerateInlines(), inline => inline.Kind == MarkdownNativeInlineKind.Abbreviation);
        var nativeTitle = Assert.Single(nativeAbbreviation.Metadata, metadata => metadata.Name == "title");
        var written = result.Document.ToMarkdown(new MarkdownWriteOptions { OutputLineEnding = "\n" }).TrimEnd('\n');

        Assert.Equal("HTML", definition.Label);
        Assert.Equal(string.Empty, definition.Title);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 6), definition.LabelSourceSpan);
        Assert.Null(definition.TitleSourceSpan);
        Assert.Equal("HTML", abbreviation.Text);
        Assert.Equal(string.Empty, abbreviation.Title);
        Assert.Contains("<abbr title=\"\">HTML</abbr>", html, StringComparison.Ordinal);
        Assert.Equal(string.Empty, nativeTitle.Value);
        Assert.Null(nativeTitle.SourceSpan);
        Assert.Equal(markdown.TrimEnd('\n'), written);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void Abbreviations_Use_Markdig_Boundaries_Around_Dashes_And_Opening_Punctuation() {
        const string markdown = "*[HTML]: Hyper Text Markup Language\n\nHTML- HTML-like (HTML) 'HTML' \"HTML\" /HTML .HTML";
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.Abbreviations = true;

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        var html = result.Document.ToHtmlFragment(CreatePlainHtmlOptions());
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(result.Document.Blocks));
        var abbreviations = paragraph.Inlines.Nodes.OfType<AbbreviationInline>().ToArray();

        var abbreviation = Assert.Single(abbreviations);
        Assert.Equal("HTML", abbreviation.Text);
        Assert.Contains("<abbr title=\"Hyper Text Markup Language\">HTML</abbr>- HTML-like", html, StringComparison.Ordinal);
        Assert.Contains("(HTML) &#39;HTML&#39; &quot;HTML&quot; /HTML .HTML", html, StringComparison.Ordinal);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void Abbreviations_Expand_Inside_Unresolved_Bracket_Text_Like_Markdig() {
        const string markdown = "*[HTML]: Hyper Text\n\n[HTML]";
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.Abbreviations = true;

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        var html = result.Document.ToHtmlFragment(CreatePlainHtmlOptions());
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(result.Document.Blocks));
        var abbreviation = Assert.IsType<AbbreviationInline>(paragraph.Inlines.Nodes[1]);

        Assert.IsType<TextRun>(paragraph.Inlines.Nodes[0]);
        Assert.IsType<TextRun>(paragraph.Inlines.Nodes[2]);
        Assert.Equal("HTML", abbreviation.Text);
        Assert.Equal("Hyper Text", abbreviation.Title);
        Assert.Equal(new MarkdownSourceSpan(3, 2, 3, 5), MarkdownInlineSourceSpans.Get(abbreviation));
        Assert.Contains("[<abbr title=\"Hyper Text\">HTML</abbr>]", html, StringComparison.Ordinal);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void List_Item_Abbreviation_Definitions_Are_Consumed_And_Apply_Document_Wide() {
        const string markdown = "- *[HTML]: Hyper Text\n- HTML";
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.Abbreviations = true;

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        var html = result.Document.ToHtmlFragment(CreatePlainHtmlOptions());
        var definition = Assert.Single(result.AbbreviationDefinitions);
        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(result.Document.Blocks));
        var first = list.Items[0];
        var second = list.Items[1];
        var abbreviation = Assert.IsType<AbbreviationInline>(Assert.Single(second.Content.Nodes));
        var nativeList = Assert.IsType<MarkdownNativeListBlock>(Assert.Single(MarkdownNativeDocument.Parse(markdown, options).Blocks));
        var nativeAbbreviation = Assert.Single(nativeList.Items[1].InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Abbreviation);
        var nativeTitle = Assert.Single(nativeAbbreviation.Metadata, metadata => metadata.Name == "title");

        Assert.Empty(first.Content.Nodes);
        Assert.Equal("HTML", definition.Label);
        Assert.Equal("Hyper Text", definition.Title);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 8), definition.LabelSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 12, 1, 21), definition.TitleSourceSpan);
        Assert.Equal("HTML", abbreviation.Text);
        Assert.Equal("Hyper Text", abbreviation.Title);
        Assert.Contains("<li></li><li><abbr title=\"Hyper Text\">HTML</abbr></li>", html, StringComparison.Ordinal);
        Assert.Equal(new MarkdownSourceSpan(1, 12, 1, 21), nativeTitle.SourceSpan);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.SemanticTreeIsWellFormed(result.Document);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
    }

    [Fact]
    public void Abbreviations_Can_Render_NonAscii_Text_Literally_For_Markdig_Style_Output() {
        const string markdown = "*[åbc]: Unicode\n\nåbc ÅBC";
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.Abbreviations = true;

        var html = MarkdownReader.Parse(markdown, options).ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        });

        Assert.Contains("<abbr title=\"Unicode\">åbc</abbr> ÅBC", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Abbreviation_Definitions_Are_Written_After_FrontMatter() {
        const string markdown = "---\ntitle: Doc\n---\n\n  *[HTML]:   Hyper Text Markup Language\n\nHTML";
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.Abbreviations = true;

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        var written = result.Document.ToMarkdown(new MarkdownWriteOptions { OutputLineEnding = "\n" }).TrimEnd('\n');

        Assert.Equal("---\ntitle: Doc\n---\n\n  *[HTML]:   Hyper Text Markup Language\n\nHTML", written);
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

    [Fact]
    public void Abbreviations_Propagate_Into_Nested_Containers_And_Table_Cell_Ast() {
        const string markdown = "*[HTML]: Hyper Text Markup Language\n\n> HTML quoted\n\n| Term |\n| --- |\n| HTML |";
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.Abbreviations = true;

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        var quote = Assert.IsType<QuoteBlock>(result.Document.Blocks[0]);
        var quotedParagraph = Assert.IsType<ParagraphBlock>(Assert.Single(quote.ChildBlocks));
        var quotedAbbreviation = Assert.IsType<AbbreviationInline>(quotedParagraph.Inlines.Nodes[0]);
        Assert.Equal("HTML", quotedAbbreviation.Text);
        Assert.Equal("Hyper Text Markup Language", quotedAbbreviation.Title);

        var table = Assert.IsType<TableBlock>(result.Document.Blocks[1]);
        var tableAbbreviation = Assert.IsType<AbbreviationInline>(Assert.Single(table.RowInlines[0][0].Nodes));
        Assert.Equal("HTML", tableAbbreviation.Text);
        Assert.Equal("Hyper Text Markup Language", tableAbbreviation.Title);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeTable = Assert.IsType<MarkdownNativeTableBlock>(native.Blocks[1]);
        var nativeCellAbbreviation = Assert.Single(nativeTable.Rows[0][0].InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Abbreviation);
        Assert.Equal("HTML", nativeCellAbbreviation.Text);
        Assert.Equal("Hyper Text Markup Language", Assert.Single(nativeCellAbbreviation.Metadata, metadata => metadata.Name == "title").Value);
    }

    private static HtmlOptions CreatePlainHtmlOptions() =>
        new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        };

    private static string NormalizeMarkdown(string markdown) =>
        markdown.Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd('\n');
}
