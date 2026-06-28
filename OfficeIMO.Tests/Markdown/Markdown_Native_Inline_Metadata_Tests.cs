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
        var content = Assert.Single(code.Metadata, metadata => metadata.Name == "content");
        var closing = Assert.Single(code.Metadata, metadata => metadata.Name == "closingMarker");

        Assert.Equal("code ` tick", code.Text);
        Assert.Equal("``", opening.Value);
        Assert.Equal("code ` tick", content.Value);
        Assert.Equal("``", closing.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 6), opening.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 17), content.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 18, 1, 19), closing.SourceSpan);
        Assert.Collection(code.SyntaxNode.Children,
            openingMarker => Assert.Equal(MarkdownSyntaxKind.InlineOpeningMarker, openingMarker.Kind),
            contentToken => Assert.Equal(MarkdownSyntaxKind.InlineCodeSpanContent, contentToken.Kind),
            closingMarker => Assert.Equal(MarkdownSyntaxKind.InlineClosingMarker, closingMarker.Kind));
        Assert.Empty(code.Children);

        var edited = native.CreateReplaceEdit(closing, "```").Apply(native.SourceMarkdown);
        edited = native.CreateReplaceEdit(opening, "```").Apply(edited);
        Assert.Equal("Use ```code ` tick``` now\n", edited);
        Assert.Equal("Use ``updated`` now\n", native.CreateReplaceEdit(content, "updated").Apply(native.SourceMarkdown));

        var snapshotCode = Assert.Single(native.ToSnapshot().Blocks[0].Inlines, inline => inline.Kind == MarkdownNativeInlineKind.Code);
        Assert.Equal("``", snapshotCode.Metadata["openingMarker"]);
        Assert.Equal("code ` tick", snapshotCode.Metadata["content"]);
        Assert.Equal("``", snapshotCode.Metadata["closingMarker"]);
        Assert.Equal(5, snapshotCode.MetadataSourceSpans["openingMarker"]!.StartColumn);
        Assert.Equal(6, snapshotCode.MetadataSourceSpans["openingMarker"]!.EndColumn);
        Assert.Equal(7, snapshotCode.MetadataSourceSpans["content"]!.StartColumn);
        Assert.Equal(17, snapshotCode.MetadataSourceSpans["content"]!.EndColumn);
        Assert.Equal(18, snapshotCode.MetadataSourceSpans["closingMarker"]!.StartColumn);
        Assert.Equal(19, snapshotCode.MetadataSourceSpans["closingMarker"]!.EndColumn);
    }

    [Fact]
    public void Backslash_Escape_Metadata_Is_Source_Addressable_In_Native_Projection_And_Snapshots() {
        const string markdown = "Use \\*literal\\* and C:\\Temp\n";

        var native = MarkdownNativeDocument.Parse(markdown);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var escaped = paragraph.InlineRuns
            .Where(inline => inline.Kind == MarkdownNativeInlineKind.Text && inline.GetMetadata("escapeMarker") != null)
            .ToArray();

        Assert.Equal(2, escaped.Length);
        Assert.All(escaped, inline => Assert.Equal("*", inline.Text));
        Assert.DoesNotContain(
            paragraph.InlineRuns,
            inline => inline.Kind == MarkdownNativeInlineKind.Text && inline.Text == "\\" && inline.GetMetadata("escapeMarker") != null);

        var firstMarker = Assert.Single(escaped[0].Metadata, metadata => metadata.Name == "escapeMarker");
        var firstCharacter = Assert.Single(escaped[0].Metadata, metadata => metadata.Name == "escapedCharacter");
        var secondMarker = Assert.Single(escaped[1].Metadata, metadata => metadata.Name == "escapeMarker");
        var secondCharacter = Assert.Single(escaped[1].Metadata, metadata => metadata.Name == "escapedCharacter");

        Assert.Equal("\\", firstMarker.Value);
        Assert.Equal("*", firstCharacter.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 5), firstMarker.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 6), firstCharacter.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 14, 1, 14), secondMarker.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 15, 1, 15), secondCharacter.SourceSpan);

        Assert.Equal("Use *literal\\* and C:\\Temp\n", native.CreateReplaceEdit(firstMarker, string.Empty).Apply(native.SourceMarkdown));
        Assert.Equal("Use \\_literal\\* and C:\\Temp\n", native.CreateReplaceEdit(firstCharacter, "_").Apply(native.SourceMarkdown));
        Assert.Equal("Use \\*literal* and C:\\Temp\n", native.CreateReplaceEdit(secondMarker, string.Empty).Apply(native.SourceMarkdown));
        Assert.Equal("Use \\*literal\\_ and C:\\Temp\n", native.CreateReplaceEdit(secondCharacter, "_").Apply(native.SourceMarkdown));

        var snapshotEscaped = native.ToSnapshot().Blocks[0].Inlines
            .Where(inline => inline.Kind == MarkdownNativeInlineKind.Text && inline.Metadata.ContainsKey("escapeMarker"))
            .ToArray();

        Assert.Equal(2, snapshotEscaped.Length);
        Assert.Equal("\\", snapshotEscaped[0].Metadata["escapeMarker"]);
        Assert.Equal("*", snapshotEscaped[0].Metadata["escapedCharacter"]);
        Assert.Equal(5, snapshotEscaped[0].MetadataSourceSpans["escapeMarker"]!.StartColumn);
        Assert.Equal(6, snapshotEscaped[0].MetadataSourceSpans["escapedCharacter"]!.StartColumn);
        Assert.Equal(14, snapshotEscaped[1].MetadataSourceSpans["escapeMarker"]!.StartColumn);
        Assert.Equal(15, snapshotEscaped[1].MetadataSourceSpans["escapedCharacter"]!.StartColumn);
    }

    [Fact]
    public void Decoded_Html_Entity_Metadata_Preserves_Source_Text_For_Edits_And_Snapshots() {
        const string markdown = "Use &amp; and &#35; symbols\n";

        var native = MarkdownNativeDocument.Parse(markdown);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var entities = paragraph.InlineRuns
            .Where(inline => inline.Kind == MarkdownNativeInlineKind.Text && inline.GetMetadata("sourceText") != null)
            .ToArray();

        Assert.Equal(2, entities.Length);
        Assert.Equal("&", entities[0].Text);
        Assert.Equal("#", entities[1].Text);

        var ampSource = Assert.Single(entities[0].Metadata, metadata => metadata.Name == "sourceText");
        var hashSource = Assert.Single(entities[1].Metadata, metadata => metadata.Name == "sourceText");

        Assert.Equal("&amp;", ampSource.Value);
        Assert.Equal("&#35;", hashSource.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 9), ampSource.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 15, 1, 19), hashSource.SourceSpan);
        Assert.Equal("Use &lt; and &#35; symbols\n", native.CreateReplaceEdit(ampSource, "&lt;").Apply(native.SourceMarkdown));
        Assert.Equal("Use &amp; and &#x23; symbols\n", native.CreateReplaceEdit(hashSource, "&#x23;").Apply(native.SourceMarkdown));

        var snapshotEntities = native.ToSnapshot().Blocks[0].Inlines
            .Where(inline => inline.Kind == MarkdownNativeInlineKind.Text && inline.Metadata.ContainsKey("sourceText"))
            .ToArray();

        Assert.Equal(2, snapshotEntities.Length);
        Assert.Equal("&amp;", snapshotEntities[0].Metadata["sourceText"]);
        Assert.Equal("&#35;", snapshotEntities[1].Metadata["sourceText"]);
        Assert.Equal(5, snapshotEntities[0].MetadataSourceSpans["sourceText"]!.StartColumn);
        Assert.Equal(9, snapshotEntities[0].MetadataSourceSpans["sourceText"]!.EndColumn);
        Assert.Equal(15, snapshotEntities[1].MetadataSourceSpans["sourceText"]!.StartColumn);
        Assert.Equal(19, snapshotEntities[1].MetadataSourceSpans["sourceText"]!.EndColumn);
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
    public void Hard_Break_Marker_Metadata_Preserves_Source_Spelling_For_Edits_And_Snapshots() {
        var spaces = MarkdownNativeDocument.Parse("Alpha  \nbravo\n");
        var spacesBreak = Assert.Single(
            Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(spaces.Blocks)).InlineRuns,
            inline => inline.Kind == MarkdownNativeInlineKind.HardBreak);
        var spacesMarker = Assert.Single(spacesBreak.Metadata, metadata => metadata.Name == "marker");

        Assert.Equal("  ", spacesMarker.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 7), spacesBreak.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 7), spacesMarker.SourceSpan);
        Assert.Equal("Alpha\\\nbravo\n", spaces.CreateReplaceEdit(spacesMarker, "\\").Apply(spaces.SourceMarkdown));
        Assert.Equal("  ", spaces.ToSnapshot().Blocks[0].Inlines.Single(inline => inline.Kind == MarkdownNativeInlineKind.HardBreak).Metadata["marker"]);
        Assert.Equal(7, spaces.ToSnapshot().Blocks[0].Inlines.Single(inline => inline.Kind == MarkdownNativeInlineKind.HardBreak).MetadataSourceSpans["marker"]!.EndColumn);

        var backslash = MarkdownNativeDocument.Parse("Alpha\\\nbravo\n");
        var backslashBreak = Assert.Single(
            Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(backslash.Blocks)).InlineRuns,
            inline => inline.Kind == MarkdownNativeInlineKind.HardBreak);
        var backslashMarker = Assert.Single(backslashBreak.Metadata, metadata => metadata.Name == "marker");

        Assert.Equal("\\", backslashMarker.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 6), backslashMarker.SourceSpan);
        Assert.Equal("Alpha  \nbravo\n", backslash.CreateReplaceEdit(backslashMarker, "  ").Apply(backslash.SourceMarkdown));
        Assert.Equal("\\", backslash.ToSnapshot().Blocks[0].Inlines.Single(inline => inline.Kind == MarkdownNativeInlineKind.HardBreak).Metadata["marker"]);

        var html = MarkdownNativeDocument.Parse("Alpha<br />bravo\n");
        var htmlBreak = Assert.Single(
            Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(html.Blocks)).InlineRuns,
            inline => inline.Kind == MarkdownNativeInlineKind.HardBreak);
        var htmlMarker = Assert.Single(htmlBreak.Metadata, metadata => metadata.Name == "marker");

        Assert.Equal("<br />", htmlMarker.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 11), htmlBreak.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 11), htmlMarker.SourceSpan);
        Assert.Equal("Alpha<br>bravo\n", html.CreateReplaceEdit(htmlMarker, "<br>").Apply(html.SourceMarkdown));
        Assert.Equal("<br />", html.ToSnapshot().Blocks[0].Inlines.Single(inline => inline.Kind == MarkdownNativeInlineKind.HardBreak).Metadata["marker"]);
    }

    [Fact]
    public void Autolink_Metadata_Preserves_Target_And_Angle_Markers_For_Edits_And_Snapshots() {
        const string markdown = "Go <https://example.com/docs> and mailto:user@example.com\n";

        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var links = paragraph.InlineRuns.Where(inline => inline.Kind == MarkdownNativeInlineKind.Link).ToArray();
        Assert.Equal(2, links.Length);

        var angle = links[0];
        var bare = links[1];
        var angleTarget = Assert.Single(angle.Metadata, metadata => metadata.Name == "target");
        var opening = Assert.Single(angle.Metadata, metadata => metadata.Name == "openingMarker");
        var closing = Assert.Single(angle.Metadata, metadata => metadata.Name == "closingMarker");
        var bareTarget = Assert.Single(bare.Metadata, metadata => metadata.Name == "target");

        Assert.Equal("https://example.com/docs", angle.Text);
        Assert.Equal("https://example.com/docs", angleTarget.Value);
        Assert.Equal("<", opening.Value);
        Assert.Equal(">", closing.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 28), angleTarget.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 4, 1, 4), opening.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 29, 1, 29), closing.SourceSpan);
        Assert.Equal("mailto:user@example.com", bareTarget.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 35, 1, 57), bareTarget.SourceSpan);

        var edited = native.CreateReplaceEdit(angleTarget, "https://contoso.test/docs").Apply(native.SourceMarkdown);
        Assert.Equal("Go <https://contoso.test/docs> and mailto:user@example.com\n", edited);
        edited = native.CreateReplaceEdit(opening, "[").Apply(native.SourceMarkdown);
        edited = native.CreateReplaceEdit(closing, "]").Apply(edited);
        Assert.Equal("Go [https://example.com/docs] and mailto:user@example.com\n", edited);
        Assert.Equal("Go <https://example.com/docs> and mailto:team@example.com\n", native.CreateReplaceEdit(bareTarget, "mailto:team@example.com").Apply(native.SourceMarkdown));

        var snapshotLinks = native.ToSnapshot().Blocks[0].Inlines.Where(inline => inline.Kind == MarkdownNativeInlineKind.Link).ToArray();
        Assert.Equal("https://example.com/docs", snapshotLinks[0].Metadata["target"]);
        Assert.Equal("<", snapshotLinks[0].Metadata["openingMarker"]);
        Assert.Equal(">", snapshotLinks[0].Metadata["closingMarker"]);
        Assert.Equal(5, snapshotLinks[0].MetadataSourceSpans["target"]!.StartColumn);
        Assert.Equal(29, snapshotLinks[0].MetadataSourceSpans["closingMarker"]!.EndColumn);
        Assert.Equal("mailto:user@example.com", snapshotLinks[1].Metadata["target"]);
        Assert.Equal(35, snapshotLinks[1].MetadataSourceSpans["target"]!.StartColumn);
    }

    [Fact]
    public void Footnote_Reference_Metadata_Preserves_Label_And_Delimiter_Markers_For_Edits_And_Snapshots() {
        const string markdown = "See [^note]\n\n[^note]: Body\n";

        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(native.Blocks[0]);
        var footnote = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.FootnoteRef);
        var label = Assert.Single(footnote.Metadata, metadata => metadata.Name == "label");
        var opening = Assert.Single(footnote.Metadata, metadata => metadata.Name == "openingMarker");
        var closing = Assert.Single(footnote.Metadata, metadata => metadata.Name == "closingMarker");

        Assert.Equal("note", label.Value);
        Assert.Equal("[^", opening.Value);
        Assert.Equal("]", closing.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 10), label.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 6), opening.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 11, 1, 11), closing.SourceSpan);
        Assert.Equal("See [^todo]\n\n[^note]: Body\n", native.CreateReplaceEdit(label, "todo").Apply(native.SourceMarkdown));

        var edited = native.CreateReplaceEdit(opening, "{^").Apply(native.SourceMarkdown);
        edited = native.CreateReplaceEdit(closing, "}").Apply(edited);
        Assert.Equal("See {^note}\n\n[^note]: Body\n", edited);

        var snapshotFootnote = Assert.Single(native.ToSnapshot().Blocks[0].Inlines, inline => inline.Kind == MarkdownNativeInlineKind.FootnoteRef);
        Assert.Equal("note", snapshotFootnote.Metadata["label"]);
        Assert.Equal("[^", snapshotFootnote.Metadata["openingMarker"]);
        Assert.Equal("]", snapshotFootnote.Metadata["closingMarker"]);
        Assert.Equal(5, snapshotFootnote.MetadataSourceSpans["openingMarker"]!.StartColumn);
        Assert.Equal(11, snapshotFootnote.MetadataSourceSpans["closingMarker"]!.EndColumn);
    }

    [Fact]
    public void Inline_Link_Metadata_Preserves_Target_Title_And_Delimiter_Markers_For_Edits_And_Snapshots() {
        const string markdown = "See [docs](https://example.com \"Title\") now\n";

        var native = MarkdownNativeDocument.Parse(markdown);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var link = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Link);
        var target = Assert.Single(link.Metadata, metadata => metadata.Name == "target");
        var title = Assert.Single(link.Metadata, metadata => metadata.Name == "title");
        var opening = Assert.Single(link.Metadata, metadata => metadata.Name == "openingMarker");
        var separator = Assert.Single(link.Metadata, metadata => metadata.Name == "separatorMarker");
        var closing = Assert.Single(link.Metadata, metadata => metadata.Name == "closingMarker");

        Assert.Equal("https://example.com", target.Value);
        Assert.Equal("Title", title.Value);
        Assert.Equal("[", opening.Value);
        Assert.Equal("](", separator.Value);
        Assert.Equal(")", closing.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 12, 1, 30), target.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 33, 1, 37), title.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 5), opening.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 10, 1, 11), separator.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 39, 1, 39), closing.SourceSpan);

        Assert.Equal(
            "See [docs](https://contoso.test \"Title\") now\n",
            native.CreateReplaceEdit(target, "https://contoso.test").Apply(native.SourceMarkdown));
        Assert.Equal(
            "See [docs](https://example.com \"Docs\") now\n",
            native.CreateReplaceEdit(title, "Docs").Apply(native.SourceMarkdown));

        var edited = native.CreateReplaceEdit(opening, "{").Apply(native.SourceMarkdown);
        edited = native.CreateReplaceEdit(closing, "}").Apply(edited);
        Assert.Equal("See {docs](https://example.com \"Title\"} now\n", edited);
        Assert.Equal(
            "See [docs]{https://example.com \"Title\") now\n",
            native.CreateReplaceEdit(separator, "]{").Apply(native.SourceMarkdown));

        var snapshotLink = Assert.Single(native.ToSnapshot().Blocks[0].Inlines, inline => inline.Kind == MarkdownNativeInlineKind.Link);
        Assert.Equal("https://example.com", snapshotLink.Metadata["target"]);
        Assert.Equal("Title", snapshotLink.Metadata["title"]);
        Assert.Equal("[", snapshotLink.Metadata["openingMarker"]);
        Assert.Equal("](", snapshotLink.Metadata["separatorMarker"]);
        Assert.Equal(")", snapshotLink.Metadata["closingMarker"]);
        Assert.Equal(5, snapshotLink.MetadataSourceSpans["openingMarker"]!.StartColumn);
        Assert.Equal(10, snapshotLink.MetadataSourceSpans["separatorMarker"]!.StartColumn);
        Assert.Equal(39, snapshotLink.MetadataSourceSpans["closingMarker"]!.EndColumn);
    }

    [Fact]
    public void Reference_Link_Metadata_Preserves_Inline_Delimiter_Markers_For_Edits_And_Snapshots() {
        const string markdown = "See [docs][hero], [api][], and [guide]\n\n[hero]: https://example.com/docs \"Docs\"\n[api]: /api\n[guide]: /guide\n";

        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(native.Blocks[0]);
        var links = paragraph.InlineRuns.Where(inline => inline.Kind == MarkdownNativeInlineKind.Link).ToArray();
        Assert.Equal(3, links.Length);

        var fullOpening = Assert.Single(links[0].Metadata, metadata => metadata.Name == "openingMarker");
        var fullSeparator = Assert.Single(links[0].Metadata, metadata => metadata.Name == "separatorMarker");
        var fullClosing = Assert.Single(links[0].Metadata, metadata => metadata.Name == "closingMarker");
        var collapsedOpening = Assert.Single(links[1].Metadata, metadata => metadata.Name == "openingMarker");
        var collapsedSeparator = Assert.Single(links[1].Metadata, metadata => metadata.Name == "separatorMarker");
        var collapsedClosing = Assert.Single(links[1].Metadata, metadata => metadata.Name == "closingMarker");
        var shortcutOpening = Assert.Single(links[2].Metadata, metadata => metadata.Name == "openingMarker");
        var shortcutClosing = Assert.Single(links[2].Metadata, metadata => metadata.Name == "closingMarker");

        Assert.Equal("https://example.com/docs", links[0].GetMetadata("target"));
        Assert.Equal("/api", links[1].GetMetadata("target"));
        Assert.Equal("/guide", links[2].GetMetadata("target"));
        Assert.Equal("[", fullOpening.Value);
        Assert.Equal("][", fullSeparator.Value);
        Assert.Equal("]", fullClosing.Value);
        Assert.Equal("[", collapsedOpening.Value);
        Assert.Equal("][", collapsedSeparator.Value);
        Assert.Equal("]", collapsedClosing.Value);
        Assert.Equal("[", shortcutOpening.Value);
        Assert.Equal("]", shortcutClosing.Value);
        Assert.DoesNotContain(links[2].Metadata, metadata => metadata.Name == "separatorMarker");
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 5), fullOpening.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 10, 1, 11), fullSeparator.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 16, 1, 16), fullClosing.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 19, 1, 19), collapsedOpening.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 23, 1, 24), collapsedSeparator.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 25, 1, 25), collapsedClosing.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 32, 1, 32), shortcutOpening.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 38, 1, 38), shortcutClosing.SourceSpan);

        var referenceEdited = native.CreateReplaceEdit(fullSeparator, "](").Apply(native.SourceMarkdown);
        referenceEdited = native.CreateReplaceEdit(fullClosing, ")").Apply(referenceEdited);
        Assert.Equal("See [docs](hero), [api][], and [guide]\n\n[hero]: https://example.com/docs \"Docs\"\n[api]: /api\n[guide]: /guide\n", referenceEdited);

        var edited = native.CreateReplaceEdit(shortcutOpening, "{").Apply(native.SourceMarkdown);
        edited = native.CreateReplaceEdit(shortcutClosing, "}").Apply(edited);
        Assert.Equal("See [docs][hero], [api][], and {guide}\n\n[hero]: https://example.com/docs \"Docs\"\n[api]: /api\n[guide]: /guide\n", edited);

        var snapshotLinks = native.ToSnapshot().Blocks[0].Inlines.Where(inline => inline.Kind == MarkdownNativeInlineKind.Link).ToArray();
        Assert.Equal("[", snapshotLinks[0].Metadata["openingMarker"]);
        Assert.Equal("][", snapshotLinks[0].Metadata["separatorMarker"]);
        Assert.Equal("]", snapshotLinks[0].Metadata["closingMarker"]);
        Assert.Equal("[", snapshotLinks[1].Metadata["openingMarker"]);
        Assert.Equal("][", snapshotLinks[1].Metadata["separatorMarker"]);
        Assert.Equal("]", snapshotLinks[1].Metadata["closingMarker"]);
        Assert.Equal("[", snapshotLinks[2].Metadata["openingMarker"]);
        Assert.Equal("]", snapshotLinks[2].Metadata["closingMarker"]);
        Assert.False(snapshotLinks[2].Metadata.ContainsKey("separatorMarker"));
        Assert.Equal(5, snapshotLinks[0].MetadataSourceSpans["openingMarker"]!.StartColumn);
        Assert.Equal(10, snapshotLinks[0].MetadataSourceSpans["separatorMarker"]!.StartColumn);
        Assert.Equal(24, snapshotLinks[1].MetadataSourceSpans["separatorMarker"]!.EndColumn);
        Assert.Equal(25, snapshotLinks[1].MetadataSourceSpans["closingMarker"]!.EndColumn);
        Assert.Equal(38, snapshotLinks[2].MetadataSourceSpans["closingMarker"]!.EndColumn);
    }

    [Fact]
    public void Inline_Image_Metadata_Preserves_Source_Title_And_Delimiter_Markers_For_Edits_And_Snapshots() {
        const string markdown = "Look ![Alt](img.png \"Title\") now\n";

        var native = MarkdownNativeDocument.Parse(markdown);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var image = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Image);
        var alt = Assert.Single(image.Metadata, metadata => metadata.Name == "alt");
        var source = Assert.Single(image.Metadata, metadata => metadata.Name == "source");
        var title = Assert.Single(image.Metadata, metadata => metadata.Name == "imageTitle");
        var opening = Assert.Single(image.Metadata, metadata => metadata.Name == "openingMarker");
        var separator = Assert.Single(image.Metadata, metadata => metadata.Name == "separatorMarker");
        var closing = Assert.Single(image.Metadata, metadata => metadata.Name == "closingMarker");

        Assert.Equal("Alt", alt.Value);
        Assert.Equal("img.png", source.Value);
        Assert.Equal("Title", title.Value);
        Assert.Equal("![", opening.Value);
        Assert.Equal("](", separator.Value);
        Assert.Equal(")", closing.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 8, 1, 10), alt.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 13, 1, 19), source.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 22, 1, 26), title.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 7), opening.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 11, 1, 12), separator.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 28, 1, 28), closing.SourceSpan);

        Assert.Equal(
            "Look ![Logo](img.png \"Title\") now\n",
            native.CreateReplaceEdit(alt, "Logo").Apply(native.SourceMarkdown));
        Assert.Equal(
            "Look ![Alt](logo.svg \"Title\") now\n",
            native.CreateReplaceEdit(source, "logo.svg").Apply(native.SourceMarkdown));
        Assert.Equal(
            "Look ![Alt](img.png \"Diagram\") now\n",
            native.CreateReplaceEdit(title, "Diagram").Apply(native.SourceMarkdown));

        var edited = native.CreateReplaceEdit(opening, "?[").Apply(native.SourceMarkdown);
        edited = native.CreateReplaceEdit(closing, "]").Apply(edited);
        Assert.Equal("Look ?[Alt](img.png \"Title\"] now\n", edited);
        Assert.Equal(
            "Look ![Alt]{img.png \"Title\") now\n",
            native.CreateReplaceEdit(separator, "]{").Apply(native.SourceMarkdown));

        var snapshotImage = Assert.Single(native.ToSnapshot().Blocks[0].Inlines, inline => inline.Kind == MarkdownNativeInlineKind.Image);
        Assert.Equal("Alt", snapshotImage.Metadata["alt"]);
        Assert.Equal("img.png", snapshotImage.Metadata["source"]);
        Assert.Equal("Title", snapshotImage.Metadata["imageTitle"]);
        Assert.Equal("![", snapshotImage.Metadata["openingMarker"]);
        Assert.Equal("](", snapshotImage.Metadata["separatorMarker"]);
        Assert.Equal(")", snapshotImage.Metadata["closingMarker"]);
        Assert.Equal(6, snapshotImage.MetadataSourceSpans["openingMarker"]!.StartColumn);
        Assert.Equal(12, snapshotImage.MetadataSourceSpans["separatorMarker"]!.EndColumn);
        Assert.Equal(28, snapshotImage.MetadataSourceSpans["closingMarker"]!.EndColumn);
    }

    [Fact]
    public void Reference_Image_Metadata_Preserves_Inline_Delimiter_Markers_For_Edits_And_Snapshots() {
        const string markdown = "Look ![Alt][img] and ![Icon]\n\n[img]: img.png \"Img\"\n[Icon]: icon.png\n";

        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(native.Blocks[0]);
        var images = paragraph.InlineRuns.Where(inline => inline.Kind == MarkdownNativeInlineKind.Image).ToArray();
        Assert.Equal(2, images.Length);

        var fullOpening = Assert.Single(images[0].Metadata, metadata => metadata.Name == "openingMarker");
        var fullSeparator = Assert.Single(images[0].Metadata, metadata => metadata.Name == "separatorMarker");
        var fullClosing = Assert.Single(images[0].Metadata, metadata => metadata.Name == "closingMarker");
        var shortcutOpening = Assert.Single(images[1].Metadata, metadata => metadata.Name == "openingMarker");
        var shortcutClosing = Assert.Single(images[1].Metadata, metadata => metadata.Name == "closingMarker");

        Assert.Equal("Alt", images[0].GetMetadata("alt"));
        Assert.Equal("img.png", images[0].GetMetadata("source"));
        Assert.Equal("Icon", images[1].GetMetadata("alt"));
        Assert.Equal("icon.png", images[1].GetMetadata("source"));
        Assert.Equal("![", fullOpening.Value);
        Assert.Equal("][", fullSeparator.Value);
        Assert.Equal("]", fullClosing.Value);
        Assert.Equal("![", shortcutOpening.Value);
        Assert.Equal("]", shortcutClosing.Value);
        Assert.DoesNotContain(images[1].Metadata, metadata => metadata.Name == "separatorMarker");
        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 7), fullOpening.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 11, 1, 12), fullSeparator.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 16, 1, 16), fullClosing.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 22, 1, 23), shortcutOpening.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 28, 1, 28), shortcutClosing.SourceSpan);

        var referenceEdited = native.CreateReplaceEdit(fullSeparator, "](").Apply(native.SourceMarkdown);
        referenceEdited = native.CreateReplaceEdit(fullClosing, ")").Apply(referenceEdited);
        Assert.Equal("Look ![Alt](img) and ![Icon]\n\n[img]: img.png \"Img\"\n[Icon]: icon.png\n", referenceEdited);

        var edited = native.CreateReplaceEdit(shortcutOpening, "?[").Apply(native.SourceMarkdown);
        edited = native.CreateReplaceEdit(shortcutClosing, "}").Apply(edited);
        Assert.Equal("Look ![Alt][img] and ?[Icon}\n\n[img]: img.png \"Img\"\n[Icon]: icon.png\n", edited);

        var snapshotImages = native.ToSnapshot().Blocks[0].Inlines.Where(inline => inline.Kind == MarkdownNativeInlineKind.Image).ToArray();
        Assert.Equal("![", snapshotImages[0].Metadata["openingMarker"]);
        Assert.Equal("][", snapshotImages[0].Metadata["separatorMarker"]);
        Assert.Equal("]", snapshotImages[0].Metadata["closingMarker"]);
        Assert.Equal("![", snapshotImages[1].Metadata["openingMarker"]);
        Assert.Equal("]", snapshotImages[1].Metadata["closingMarker"]);
        Assert.False(snapshotImages[1].Metadata.ContainsKey("separatorMarker"));
        Assert.Equal(6, snapshotImages[0].MetadataSourceSpans["openingMarker"]!.StartColumn);
        Assert.Equal(11, snapshotImages[0].MetadataSourceSpans["separatorMarker"]!.StartColumn);
        Assert.Equal(16, snapshotImages[0].MetadataSourceSpans["closingMarker"]!.EndColumn);
        Assert.Equal(22, snapshotImages[1].MetadataSourceSpans["openingMarker"]!.StartColumn);
        Assert.Equal(28, snapshotImages[1].MetadataSourceSpans["closingMarker"]!.EndColumn);
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
        var opening = Assert.Single(imageLink.Metadata, metadata => metadata.Name == "openingMarker");
        var separator = Assert.Single(imageLink.Metadata, metadata => metadata.Name == "separatorMarker");
        var closing = Assert.Single(imageLink.Metadata, metadata => metadata.Name == "closingMarker");

        Assert.Equal(new MarkdownSourceSpan(1, 14, 1, 16), alt.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 19, 1, 25), source.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 28, 1, 30), imageTitle.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 35, 1, 53), target.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 56, 1, 65), title.SourceSpan);
        Assert.Equal("[", opening.Value);
        Assert.Equal("](", separator.Value);
        Assert.Equal(")", closing.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 11, 1, 11), opening.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 33, 1, 34), separator.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 67, 1, 67), closing.SourceSpan);

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
        var edited = native.CreateReplaceEdit(opening, "{").Apply(native.SourceMarkdown);
        edited = native.CreateReplaceEdit(closing, "}").Apply(edited);
        Assert.Equal(
            "Paragraph {![Alt](img.png \"Img\")](https://example.com \"Link title\"}.",
            edited);
        edited = native.CreateReplaceEdit(separator, "]{").Apply(native.SourceMarkdown);
        Assert.Equal(
            "Paragraph [![Alt](img.png \"Img\")]{https://example.com \"Link title\").",
            edited);

        var snapshot = native.ToSnapshot();
        var snapshotParagraph = Assert.Single(snapshot.Blocks);
        var snapshotImageLink = Assert.Single(snapshotParagraph.Inlines, inline => inline.Kind == MarkdownNativeInlineKind.ImageLink);

        Assert.Equal("Alt", snapshotImageLink.Metadata["alt"]);
        Assert.Equal("img.png", snapshotImageLink.Metadata["source"]);
        Assert.Equal("Img", snapshotImageLink.Metadata["imageTitle"]);
        Assert.Equal("https://example.com", snapshotImageLink.Metadata["target"]);
        Assert.Equal("Link title", snapshotImageLink.Metadata["title"]);
        Assert.Equal("[", snapshotImageLink.Metadata["openingMarker"]);
        Assert.Equal("](", snapshotImageLink.Metadata["separatorMarker"]);
        Assert.Equal(")", snapshotImageLink.Metadata["closingMarker"]);
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
        Assert.Equal(11, snapshotImageLink.MetadataSourceSpans["openingMarker"]!.StartColumn);
        Assert.Equal(33, snapshotImageLink.MetadataSourceSpans["separatorMarker"]!.StartColumn);
        Assert.Equal(34, snapshotImageLink.MetadataSourceSpans["separatorMarker"]!.EndColumn);
        Assert.Equal(67, snapshotImageLink.MetadataSourceSpans["closingMarker"]!.EndColumn);
    }
}
