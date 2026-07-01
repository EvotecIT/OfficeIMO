using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Native_Block_Source_Field_Tests {
    [Fact]
    public void CodeFence_SourceFields_Expose_Fence_Token_Values() {
        const string markdown = "~~~~csharp {#code .wide}\nConsole.WriteLine(1);\n~~~~\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var code = Assert.IsType<MarkdownNativeCodeBlock>(Assert.Single(native.Blocks));

        Assert.Equal("~~~~", code.OpeningFence);
        Assert.Equal("~~~~", code.ClosingFence);

        var openingFence = Assert.Single(native.EnumerateBlockSourceFields("openingFence"));
        Assert.Same(code, openingFence.Block);
        Assert.Equal("~~~~", openingFence.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 4), openingFence.SourceSpan);

        var infoString = Assert.Single(native.EnumerateBlockSourceFields("infoString"));
        Assert.Equal("csharp {#code .wide}", infoString.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 24), infoString.SourceSpan);

        var attributes = Assert.Single(native.EnumerateBlockSourceFields("attributes"));
        Assert.Equal("{#code .wide}", attributes.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 12, 1, 24), attributes.SourceSpan);

        var content = Assert.Single(native.EnumerateBlockSourceFields("content"));
        Assert.Equal("Console.WriteLine(1);", content.Value);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 21), content.SourceSpan);

        var closingFence = Assert.Single(native.EnumerateBlockSourceFields("closingFence"));
        Assert.Equal("~~~~", closingFence.Value);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 4), closingFence.SourceSpan);

        Assert.Equal("openingFence", native.FindBlockSourceFieldAtPosition(1, 2)?.Name);
        Assert.Equal("closingFence", native.FindBlockSourceFieldAtPosition(3, 3)?.Name);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal("~~~~", snapshot.Fields["openingFence"]);
        Assert.Equal("~~~~", snapshot.Fields["closingFence"]);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "openingFence"
            && field.Value == "~~~~"
            && field.SourceSpan.StartColumn == 1
            && field.SourceSpan.EndColumn == 4);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "closingFence"
            && field.Value == "~~~~"
            && field.SourceSpan.StartLine == 3
            && field.SourceSpan.StartColumn == 1
            && field.SourceSpan.EndColumn == 4);

        var openingEdited = native.CreateReplaceEdit(openingFence, "```").Apply(native.SourceMarkdown);
        Assert.StartsWith("```csharp {#code .wide}", openingEdited.Replace("\r\n", "\n"), System.StringComparison.Ordinal);

        var closingEdited = native.CreateReplaceEdit(closingFence, "```").Apply(native.SourceMarkdown);
        Assert.Contains("\n```\n", closingEdited.Replace("\r\n", "\n"), System.StringComparison.Ordinal);
    }

    [Fact]
    public void GenericAttributes_Block_SourceField_Is_Source_Addressable() {
        const string markdown = "Alpha paragraph {#intro .lead}\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));

        Assert.Equal("intro", paragraph.Paragraph.Attributes.ElementId);
        Assert.Equal(new[] { "lead" }, paragraph.Paragraph.Attributes.Classes.ToArray());

        var attributes = Assert.Single(native.EnumerateBlockSourceFields("attributes"));
        Assert.Same(paragraph, attributes.Block);
        Assert.Equal("{#intro .lead}", attributes.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 17, 1, 30), attributes.SourceSpan);

        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 20));
        Assert.Equal("attributes", found.Name);
        Assert.Equal(attributes.SourceSpan, found.SourceSpan);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "attributes"
            && field.Value == "{#intro .lead}"
            && field.SourceSpan.StartColumn == 17
            && field.SourceSpan.EndColumn == 30);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(attributes, "{#summary .wide}"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("Alpha paragraph {#summary .wide}\n", roundtrip.Markdown);
    }

    [Fact]
    public void GenericAttributes_Heading_SourceField_Is_Source_Addressable() {
        const string markdown = "# Heading {#title .hero}\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var heading = Assert.IsType<MarkdownNativeHeadingBlock>(Assert.Single(native.Blocks));

        Assert.Equal("title", heading.Heading.Attributes.ElementId);
        Assert.Equal(new[] { "hero" }, heading.Heading.Attributes.Classes.ToArray());
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 9), heading.TextSourceSpan);

        var attributes = Assert.Single(native.EnumerateBlockSourceFields("attributes"));
        Assert.Same(heading, attributes.Block);
        Assert.Equal("{#title .hero}", attributes.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 11, 1, 24), attributes.SourceSpan);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(attributes, "{#docs .anchor}"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("# Heading {#docs .anchor}\n", roundtrip.Markdown);
    }

    [Fact]
    public void GenericAttributes_Table_SourceField_Is_Source_Addressable() {
        const string markdown = "| A {#tbl .wide title=\"Quarterly\"} |\n|---|\n| B |\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true,
            Tables = true
        };

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var table = Assert.IsType<MarkdownNativeTableBlock>(Assert.Single(native.Blocks));
        var sourceTable = Assert.IsType<TableBlock>(table.SourceBlock);

        Assert.Equal("tbl", sourceTable.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, sourceTable.Attributes.Classes.ToArray());

        var attributes = Assert.Single(native.EnumerateBlockSourceFields("attributes"));
        Assert.Same(table, attributes.Block);
        Assert.Equal("{#tbl .wide title=\"Quarterly\"}", attributes.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 34), attributes.SourceSpan);

        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 8));
        Assert.Equal("attributes", found.Name);
        Assert.Equal(attributes.SourceSpan, found.SourceSpan);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "attributes"
            && field.Value == "{#tbl .wide title=\"Quarterly\"}"
            && field.SourceSpan.StartColumn == 5
            && field.SourceSpan.EndColumn == 34);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(attributes, "{#grid .compact}"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("| A {#grid .compact} |\n|---|\n| B |\n", roundtrip.Markdown);
    }

    [Fact]
    public void GenericAttributes_ListItem_SourceField_Is_Source_Addressable() {
        const string markdown = "- item {#li .selected}\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var list = Assert.IsType<MarkdownNativeListBlock>(Assert.Single(native.Blocks));
        var item = Assert.Single(list.Items);

        Assert.Equal("li", item.Item.Attributes.ElementId);
        Assert.Equal(new[] { "selected" }, item.Item.Attributes.Classes.ToArray());

        var attributes = Assert.Single(native.EnumerateBlockSourceFields("attributes"));
        Assert.Same(list, attributes.Block);
        Assert.Equal(0, attributes.Index);
        Assert.Equal("{#li .selected}", attributes.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 8, 1, 22), attributes.SourceSpan);

        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 12));
        Assert.Equal("attributes", found.Name);
        Assert.Equal(attributes.SourceSpan, found.SourceSpan);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "attributes"
            && field.Value == "{#li .selected}"
            && field.Index == 0
            && field.SourceSpan.StartColumn == 8
            && field.SourceSpan.EndColumn == 22);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(attributes, "{#task .done}"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("- item {#task .done}\n", roundtrip.Markdown);
    }

    [Fact]
    public void Paragraph_Text_SourceField_Uses_Paragraph_Source_Span() {
        const string markdown = "Old **body**\n";

        var native = MarkdownNativeDocument.Parse(markdown);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));

        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 12), paragraph.TextSourceSpan);

        var text = Assert.Single(native.EnumerateBlockSourceFields("paragraphText"));
        Assert.Same(paragraph, text.Block);
        Assert.Equal("Old body", text.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 12), text.SourceSpan);

        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 6));
        Assert.Equal("paragraphText", found.Name);
        Assert.Equal(text.Value, found.Value);
        Assert.Equal(text.SourceSpan, found.SourceSpan);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal(1, snapshot.FieldSourceSpans["paragraphText"]!.StartColumn);
        Assert.Equal(12, snapshot.FieldSourceSpans["paragraphText"]!.EndColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "paragraphText"
            && field.Value == "Old body"
            && field.SourceSpan.StartColumn == 1
            && field.SourceSpan.EndColumn == 12);

        var edited = native.CreateReplaceEdit(text, "New _body_").Apply(native.SourceMarkdown);
        Assert.Equal("New _body_", edited.TrimEnd('\r', '\n'));
    }

    [Fact]
    public void Heading_SourceFields_Use_Atx_Closing_Marker_Token_Span() {
        const string markdown = "  ###   Trimmed ###   \n";

        var native = MarkdownNativeDocument.Parse(markdown);
        var heading = Assert.IsType<MarkdownNativeHeadingBlock>(Assert.Single(native.Blocks));

        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 5), heading.LevelSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 9, 1, 15), heading.TextSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 5), heading.Heading.LevelSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 9, 1, 15), heading.Heading.TextSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 5), heading.OpeningMarkerSourceSpan);
        Assert.Equal("###", heading.OpeningMarkerText);
        Assert.Equal(new MarkdownSourceSpan(1, 17, 1, 19), heading.ClosingMarkerSourceSpan);
        Assert.Equal("###", heading.ClosingMarkerText);

        var openingMarker = Assert.Single(native.EnumerateBlockSourceFields("openingMarker"));
        Assert.Same(heading, openingMarker.Block);
        Assert.Equal("###", openingMarker.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 5), openingMarker.SourceSpan);

        var closingMarker = Assert.Single(native.EnumerateBlockSourceFields("closingMarker"));
        Assert.Same(heading, closingMarker.Block);
        Assert.Equal("###", closingMarker.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 17, 1, 19), closingMarker.SourceSpan);

        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 18));
        Assert.Equal("closingMarker", found.Name);
        Assert.Equal(closingMarker.Value, found.Value);
        Assert.Equal(closingMarker.SourceSpan, found.SourceSpan);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "openingMarker"
            && field.Value == "###"
            && field.SourceSpan.StartColumn == 3
            && field.SourceSpan.EndColumn == 5);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "closingMarker"
            && field.Value == "###"
            && field.SourceSpan.StartColumn == 17
            && field.SourceSpan.EndColumn == 19);

        var openingEdited = native.CreateReplaceEdit(openingMarker, "##").Apply(native.SourceMarkdown);
        Assert.Equal("  ##   Trimmed ###", openingEdited.TrimEnd('\r', '\n', ' '));

        var edited = native.CreateReplaceEdit(closingMarker, "##").Apply(native.SourceMarkdown);
        Assert.Equal("  ###   Trimmed ##", edited.TrimEnd('\r', '\n', ' '));
    }

    [Fact]
    public void Heading_SourceFields_Use_Setext_Underline_Marker_Token_Span() {
        const string markdown = """
Heading Title
-------------
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var heading = Assert.IsType<MarkdownNativeHeadingBlock>(Assert.Single(native.Blocks));

        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 13), heading.LevelSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 13), heading.TextSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 13), heading.Heading.LevelSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 13), heading.Heading.TextSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 13), heading.SetextUnderlineMarkerSourceSpan);
        Assert.Equal("-------------", heading.SetextUnderlineMarkerText);

        var underlineMarker = Assert.Single(native.EnumerateBlockSourceFields("setextUnderlineMarker"));
        Assert.Same(heading, underlineMarker.Block);
        Assert.Equal("-------------", underlineMarker.Value);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 13), underlineMarker.SourceSpan);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "setextUnderlineMarker"
            && field.Value == "-------------"
            && field.SourceSpan.StartLine == 2
            && field.SourceSpan.EndColumn == 13);

        var edited = native.CreateReplaceEdit(underlineMarker, "===").Apply(native.SourceMarkdown);
        Assert.Equal("Heading Title\n===", edited.Replace("\r\n", "\n").TrimEnd('\n'));
    }

    [Fact]
    public void Footnote_Label_SourceField_Uses_Indented_Label_Token_Span() {
        const string markdown = """
  [^note]: Footnote body
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var footnote = Assert.IsType<MarkdownNativeFootnoteDefinitionBlock>(Assert.Single(native.Blocks));

        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 4), footnote.OpeningMarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 8), footnote.LabelSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 9, 1, 10), footnote.SeparatorMarkerSourceSpan);

        var openingMarker = Assert.Single(native.EnumerateBlockSourceFields("footnoteOpeningMarker"));
        Assert.Same(footnote, openingMarker.Block);
        Assert.Equal("[^", openingMarker.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 4), openingMarker.SourceSpan);

        var label = Assert.Single(native.EnumerateBlockSourceFields("label"));
        Assert.Same(footnote, label.Block);
        Assert.Equal("note", label.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 8), label.SourceSpan);
        var separatorMarker = Assert.Single(native.EnumerateBlockSourceFields("footnoteSeparatorMarker"));
        Assert.Same(footnote, separatorMarker.Block);
        Assert.Equal("]:", separatorMarker.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 9, 1, 10), separatorMarker.SourceSpan);
        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 6));
        Assert.Equal(label.Name, found.Name);
        Assert.Equal(label.Value, found.Value);
        Assert.Equal(label.SourceSpan, found.SourceSpan);
        var foundOpeningMarker = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 3));
        Assert.Equal(openingMarker.Name, foundOpeningMarker.Name);
        Assert.Equal(openingMarker.SourceSpan, foundOpeningMarker.SourceSpan);
        var foundSeparatorMarker = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 10));
        Assert.Equal(separatorMarker.Name, foundSeparatorMarker.Name);
        Assert.Equal(separatorMarker.SourceSpan, foundSeparatorMarker.SourceSpan);

        var snapshot = native.ToSnapshot().Blocks[0];
        Assert.Equal(3, snapshot.FieldSourceSpans["footnoteOpeningMarker"]!.StartColumn);
        Assert.Equal(4, snapshot.FieldSourceSpans["footnoteOpeningMarker"]!.EndColumn);
        Assert.Equal(9, snapshot.FieldSourceSpans["footnoteSeparatorMarker"]!.StartColumn);
        Assert.Equal(10, snapshot.FieldSourceSpans["footnoteSeparatorMarker"]!.EndColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "footnoteOpeningMarker"
            && field.Value == "[^"
            && field.SourceSpan.StartColumn == 3
            && field.SourceSpan.EndColumn == 4);
        var snapshotLabel = Assert.Single(snapshot.SourceFields, field => field.Name == "label");
        Assert.Equal("label", snapshotLabel.Name);
        Assert.Equal("note", snapshotLabel.Value);
        Assert.Equal(1, snapshotLabel.SourceSpan.StartLine);
        Assert.Equal(5, snapshotLabel.SourceSpan.StartColumn);
        Assert.Equal(1, snapshotLabel.SourceSpan.EndLine);
        Assert.Equal(8, snapshotLabel.SourceSpan.EndColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "footnoteSeparatorMarker"
            && field.Value == "]:"
            && field.SourceSpan.StartColumn == 9
            && field.SourceSpan.EndColumn == 10);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "footnoteBody"
            && field.Value == "Footnote body"
            && field.SourceSpan.StartColumn == 12
            && field.SourceSpan.EndColumn == 24);

        var edited = native.CreateReplaceEdit(label, "memo").Apply(native.SourceMarkdown);
        Assert.Equal("  [^memo]: Footnote body", edited.TrimEnd('\r', '\n'));

        var editedOpening = native.CreateReplaceEdit(openingMarker, "[^x-").Apply(native.SourceMarkdown);
        Assert.Equal("  [^x-note]: Footnote body", editedOpening.TrimEnd('\r', '\n'));

        var editedSeparator = native.CreateReplaceEdit(separatorMarker, "]: ").Apply(native.SourceMarkdown);
        Assert.Equal("  [^note]:  Footnote body", editedSeparator.TrimEnd('\r', '\n'));
    }

    [Fact]
    public void Footnote_Body_SourceField_Uses_Definition_Body_Span() {
        const string markdown = """
[^note]: Old body
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var footnote = Assert.IsType<MarkdownNativeFootnoteDefinitionBlock>(Assert.Single(native.Blocks));

        Assert.Equal(new MarkdownSourceSpan(1, 10, 1, 17), footnote.BodySourceSpan);

        var body = Assert.Single(native.EnumerateBlockSourceFields("footnoteBody"));
        Assert.Same(footnote, body.Block);
        Assert.Equal("Old body", body.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 10, 1, 17), body.SourceSpan);

        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 12));
        Assert.Equal("footnoteBody", found.Name);
        Assert.Equal("Old body", found.Value);
        Assert.Equal(body.SourceSpan, found.SourceSpan);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal(10, snapshot.FieldSourceSpans["footnoteBody"]!.StartColumn);
        Assert.Equal(17, snapshot.FieldSourceSpans["footnoteBody"]!.EndColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "footnoteBody"
            && field.Value == "Old body"
            && field.SourceSpan.StartColumn == 10
            && field.SourceSpan.EndColumn == 17);

        var edited = native.CreateReplaceEdit(body, "New body").Apply(native.SourceMarkdown);
        Assert.Equal("[^note]: New body", edited.TrimEnd('\r', '\n'));
    }

    [Fact]
    public void Html_Comment_SourceFields_Use_Delimiter_And_Body_Token_Spans() {
        const string markdown = """
<!-- keep
this comment
-->
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var comment = Assert.IsType<MarkdownNativeHtmlBlock>(Assert.Single(native.Blocks));

        Assert.True(comment.IsComment);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 4), comment.OpeningMarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 2, 12), comment.BodySourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 3), comment.ClosingMarkerSourceSpan);
        Assert.Equal(" keep\nthis comment", comment.CommentBody);

        var html = Assert.Single(native.EnumerateBlockSourceFields("html"));
        Assert.Same(comment, html.Block);
        Assert.Equal("<!-- keep\nthis comment\n-->", html.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 3), html.SourceSpan);

        var openingMarker = Assert.Single(native.EnumerateBlockSourceFields("htmlCommentOpeningMarker"));
        Assert.Same(comment, openingMarker.Block);
        Assert.Equal("<!--", openingMarker.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 4), openingMarker.SourceSpan);

        var body = Assert.Single(native.EnumerateBlockSourceFields("htmlCommentBody"));
        Assert.Same(comment, body.Block);
        Assert.Equal(" keep\nthis comment", body.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 2, 12), body.SourceSpan);

        var closingMarker = Assert.Single(native.EnumerateBlockSourceFields("htmlCommentClosingMarker"));
        Assert.Same(comment, closingMarker.Block);
        Assert.Equal("-->", closingMarker.Value);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 3), closingMarker.SourceSpan);

        Assert.Equal("htmlCommentOpeningMarker", native.FindBlockSourceFieldAtPosition(1, 2)!.Name);
        Assert.Equal("htmlCommentBody", native.FindBlockSourceFieldAtPosition(2, 4)!.Name);
        Assert.Equal("htmlCommentClosingMarker", native.FindBlockSourceFieldAtPosition(3, 2)!.Name);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal(1, snapshot.FieldSourceSpans["html"]!.StartLine);
        Assert.Equal(4, snapshot.FieldSourceSpans["htmlCommentOpeningMarker"]!.EndColumn);
        Assert.Equal(12, snapshot.FieldSourceSpans["htmlCommentBody"]!.EndColumn);
        Assert.Equal(3, snapshot.FieldSourceSpans["htmlCommentClosingMarker"]!.EndColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "htmlCommentBody"
            && field.Value == " keep\nthis comment"
            && field.SourceSpan.StartLine == 1
            && field.SourceSpan.EndLine == 2);

        var bodyEdited = native.CreateReplaceEdit(body, " updated\nbody").Apply(native.SourceMarkdown);
        Assert.Equal("<!-- updated\nbody\n-->", bodyEdited.TrimEnd('\r', '\n'));

        var openingEdited = native.CreateReplaceEdit(openingMarker, "<!--!").Apply(native.SourceMarkdown);
        Assert.Equal("<!--! keep\nthis comment\n-->", openingEdited.TrimEnd('\r', '\n'));

        var closingEdited = native.CreateReplaceEdit(closingMarker, "--!>").Apply(native.SourceMarkdown);
        Assert.Equal("<!-- keep\nthis comment\n--!>", closingEdited.TrimEnd('\r', '\n'));
    }

    [Fact]
    public void Html_Raw_SourceFields_Use_Tag_Frame_And_Body_Token_Spans() {
        const string markdown = """
<div>
Raw body
</div>
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var html = Assert.IsType<MarkdownNativeHtmlBlock>(Assert.Single(native.Blocks));

        Assert.False(html.IsComment);
        Assert.Equal("<div>", html.OpeningTag);
        Assert.Equal("Raw body", html.Body);
        Assert.Equal("</div>", html.ClosingTag);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 5), html.OpeningTagSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 8), html.RawBodySourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 6), html.ClosingTagSourceSpan);

        var whole = Assert.Single(native.EnumerateBlockSourceFields("html"));
        Assert.Same(html, whole.Block);
        Assert.Equal("<div>\nRaw body\n</div>", whole.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 6), whole.SourceSpan);

        var openingTag = Assert.Single(native.EnumerateBlockSourceFields("htmlOpeningTag"));
        Assert.Same(html, openingTag.Block);
        Assert.Equal("<div>", openingTag.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 5), openingTag.SourceSpan);

        var body = Assert.Single(native.EnumerateBlockSourceFields("htmlBody"));
        Assert.Same(html, body.Block);
        Assert.Equal("Raw body", body.Value);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 8), body.SourceSpan);

        var closingTag = Assert.Single(native.EnumerateBlockSourceFields("htmlClosingTag"));
        Assert.Same(html, closingTag.Block);
        Assert.Equal("</div>", closingTag.Value);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 6), closingTag.SourceSpan);

        Assert.Equal("htmlOpeningTag", native.FindBlockSourceFieldAtPosition(1, 3)!.Name);
        Assert.Equal("htmlBody", native.FindBlockSourceFieldAtPosition(2, 4)!.Name);
        Assert.Equal("htmlClosingTag", native.FindBlockSourceFieldAtPosition(3, 3)!.Name);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal(5, snapshot.FieldSourceSpans["htmlOpeningTag"]!.EndColumn);
        Assert.Equal(8, snapshot.FieldSourceSpans["htmlBody"]!.EndColumn);
        Assert.Equal(6, snapshot.FieldSourceSpans["htmlClosingTag"]!.EndColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "htmlBody"
            && field.Value == "Raw body"
            && field.SourceSpan.StartLine == 2
            && field.SourceSpan.EndColumn == 8);

        var openingEdited = native.CreateReplaceEdit(openingTag, "<section>").Apply(native.SourceMarkdown);
        Assert.Equal("<section>\nRaw body\n</div>", openingEdited.TrimEnd('\r', '\n'));

        var bodyEdited = native.CreateReplaceEdit(body, "Updated").Apply(native.SourceMarkdown);
        Assert.Equal("<div>\nUpdated\n</div>", bodyEdited.TrimEnd('\r', '\n'));

        var closingEdited = native.CreateReplaceEdit(closingTag, "</section>").Apply(native.SourceMarkdown);
        Assert.Equal("<div>\nRaw body\n</section>", closingEdited.TrimEnd('\r', '\n'));
    }

    [Fact]
    public void Html_Raw_SourceFields_Use_Delimited_Marker_And_Body_Token_Spans() {
        const string markdown = """
<![CDATA[
x < y
]]>
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var html = Assert.IsType<MarkdownNativeHtmlBlock>(Assert.Single(native.Blocks));

        Assert.False(html.IsComment);
        Assert.Equal("<![CDATA[", html.OpeningMarker);
        Assert.Equal("x < y", html.Body);
        Assert.Equal("]]>", html.ClosingMarker);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 9), html.RawOpeningMarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 5), html.RawBodySourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 3), html.RawClosingMarkerSourceSpan);

        var openingMarker = Assert.Single(native.EnumerateBlockSourceFields("htmlOpeningMarker"));
        Assert.Same(html, openingMarker.Block);
        Assert.Equal("<![CDATA[", openingMarker.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 9), openingMarker.SourceSpan);

        var body = Assert.Single(native.EnumerateBlockSourceFields("htmlBody"));
        Assert.Same(html, body.Block);
        Assert.Equal("x < y", body.Value);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 5), body.SourceSpan);

        var closingMarker = Assert.Single(native.EnumerateBlockSourceFields("htmlClosingMarker"));
        Assert.Same(html, closingMarker.Block);
        Assert.Equal("]]>", closingMarker.Value);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 3), closingMarker.SourceSpan);

        Assert.Equal("htmlOpeningMarker", native.FindBlockSourceFieldAtPosition(1, 3)!.Name);
        Assert.Equal("htmlBody", native.FindBlockSourceFieldAtPosition(2, 3)!.Name);
        Assert.Equal("htmlClosingMarker", native.FindBlockSourceFieldAtPosition(3, 2)!.Name);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal(9, snapshot.FieldSourceSpans["htmlOpeningMarker"]!.EndColumn);
        Assert.Equal(5, snapshot.FieldSourceSpans["htmlBody"]!.EndColumn);
        Assert.Equal(3, snapshot.FieldSourceSpans["htmlClosingMarker"]!.EndColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "htmlClosingMarker"
            && field.Value == "]]>"
            && field.SourceSpan.StartLine == 3);

        var bodyEdited = native.CreateReplaceEdit(body, "a > b").Apply(native.SourceMarkdown);
        Assert.Equal("<![CDATA[\na > b\n]]>", bodyEdited.TrimEnd('\r', '\n'));

        var openingEdited = native.CreateReplaceEdit(openingMarker, "<![CDATA[!").Apply(native.SourceMarkdown);
        Assert.Equal("<![CDATA[!\nx < y\n]]>", openingEdited.TrimEnd('\r', '\n'));

        var closingEdited = native.CreateReplaceEdit(closingMarker, "]]!>").Apply(native.SourceMarkdown);
        Assert.Equal("<![CDATA[\nx < y\n]]!>", closingEdited.TrimEnd('\r', '\n'));
    }

    [Fact]
    public void Html_Raw_SourceFields_Use_ProcessingInstruction_Marker_And_Body_Token_Spans() {
        const string markdown = """
<?php

  echo '>';

?>
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var html = Assert.IsType<MarkdownNativeHtmlBlock>(Assert.Single(native.Blocks));

        Assert.False(html.IsComment);
        Assert.Equal("<?", html.OpeningMarker);
        Assert.Equal("php\n\n  echo '>';", html.Body);
        Assert.Equal("?>", html.ClosingMarker);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 2), html.RawOpeningMarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 3, 11), html.RawBodySourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 2), html.RawClosingMarkerSourceSpan);

        var openingMarker = Assert.Single(native.EnumerateBlockSourceFields("htmlOpeningMarker"));
        Assert.Same(html, openingMarker.Block);
        Assert.Equal("<?", openingMarker.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 2), openingMarker.SourceSpan);

        var body = Assert.Single(native.EnumerateBlockSourceFields("htmlBody"));
        Assert.Same(html, body.Block);
        Assert.Equal("php\n\n  echo '>';", body.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 3, 11), body.SourceSpan);

        var closingMarker = Assert.Single(native.EnumerateBlockSourceFields("htmlClosingMarker"));
        Assert.Same(html, closingMarker.Block);
        Assert.Equal("?>", closingMarker.Value);
        Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 2), closingMarker.SourceSpan);

        Assert.Equal("htmlOpeningMarker", native.FindBlockSourceFieldAtPosition(1, 2)!.Name);
        Assert.Equal("htmlBody", native.FindBlockSourceFieldAtPosition(3, 6)!.Name);
        Assert.Equal("htmlClosingMarker", native.FindBlockSourceFieldAtPosition(5, 2)!.Name);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal(2, snapshot.FieldSourceSpans["htmlOpeningMarker"]!.EndColumn);
        Assert.Equal(11, snapshot.FieldSourceSpans["htmlBody"]!.EndColumn);
        Assert.Equal(2, snapshot.FieldSourceSpans["htmlClosingMarker"]!.EndColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "htmlOpeningMarker"
            && field.Value == "<?"
            && field.SourceSpan.StartLine == 1);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "htmlClosingMarker"
            && field.Value == "?>"
            && field.SourceSpan.StartLine == 5);

        var bodyEdited = native.CreateReplaceEdit(body, "xml version=\"1.0\"").Apply(native.SourceMarkdown);
        Assert.Equal("<?xml version=\"1.0\"\n\n?>", bodyEdited.TrimEnd('\r', '\n'));

        var openingEdited = native.CreateReplaceEdit(openingMarker, "<?!").Apply(native.SourceMarkdown);
        Assert.Equal("<?!php\n\n  echo '>';\n\n?>", openingEdited.TrimEnd('\r', '\n'));

        var closingEdited = native.CreateReplaceEdit(closingMarker, "?>>").Apply(native.SourceMarkdown);
        Assert.Equal("<?php\n\n  echo '>';\n\n?>>", closingEdited.TrimEnd('\r', '\n'));
    }

    [Fact]
    public void Html_Raw_SourceFields_Use_Declaration_Marker_And_Body_Token_Spans() {
        const string markdown = "<!DOCTYPE html>\n";

        var native = MarkdownNativeDocument.Parse(markdown);
        var html = Assert.IsType<MarkdownNativeHtmlBlock>(Assert.Single(native.Blocks));

        Assert.False(html.IsComment);
        Assert.Equal("<!", html.OpeningMarker);
        Assert.Equal("DOCTYPE html", html.Body);
        Assert.Equal(">", html.ClosingMarker);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 2), html.RawOpeningMarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 14), html.RawBodySourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 15, 1, 15), html.RawClosingMarkerSourceSpan);

        var openingMarker = Assert.Single(native.EnumerateBlockSourceFields("htmlOpeningMarker"));
        Assert.Same(html, openingMarker.Block);
        Assert.Equal("<!", openingMarker.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 2), openingMarker.SourceSpan);

        var body = Assert.Single(native.EnumerateBlockSourceFields("htmlBody"));
        Assert.Same(html, body.Block);
        Assert.Equal("DOCTYPE html", body.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 14), body.SourceSpan);

        var closingMarker = Assert.Single(native.EnumerateBlockSourceFields("htmlClosingMarker"));
        Assert.Same(html, closingMarker.Block);
        Assert.Equal(">", closingMarker.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 15, 1, 15), closingMarker.SourceSpan);

        Assert.Equal("htmlOpeningMarker", native.FindBlockSourceFieldAtPosition(1, 2)!.Name);
        Assert.Equal("htmlBody", native.FindBlockSourceFieldAtPosition(1, 10)!.Name);
        Assert.Equal("htmlClosingMarker", native.FindBlockSourceFieldAtPosition(1, 15)!.Name);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal(2, snapshot.FieldSourceSpans["htmlOpeningMarker"]!.EndColumn);
        Assert.Equal(14, snapshot.FieldSourceSpans["htmlBody"]!.EndColumn);
        Assert.Equal(15, snapshot.FieldSourceSpans["htmlClosingMarker"]!.EndColumn);

        var bodyEdited = native.CreateReplaceEdit(body, "doctype html").Apply(native.SourceMarkdown);
        Assert.Equal("<!doctype html>", bodyEdited.TrimEnd('\r', '\n'));

        var openingEdited = native.CreateReplaceEdit(openingMarker, "<!-").Apply(native.SourceMarkdown);
        Assert.Equal("<!-DOCTYPE html>", openingEdited.TrimEnd('\r', '\n'));

        var closingEdited = native.CreateReplaceEdit(closingMarker, "/>").Apply(native.SourceMarkdown);
        Assert.Equal("<!DOCTYPE html/>", closingEdited.TrimEnd('\r', '\n'));
    }

    [Fact]
    public void Container_Body_SourceFields_Use_Structured_Child_Block_Spans() {
        var markdown = """
> [!TIP] Title
> Body line

<details>
<summary>More</summary>

Inside
</details>
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var callout = Assert.IsType<MarkdownNativeCalloutBlock>(native.Blocks[0]);
        var details = Assert.IsType<MarkdownNativeDetailsBlock>(native.Blocks[1]);

        Assert.Equal(new MarkdownSourceSpan(2, 3, 2, 11), callout.BodySourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 9), details.OpeningTagSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 23), details.SummarySourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 9), details.SummaryOpeningTagSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 10, 5, 13), details.SummaryTextSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 14, 5, 23), details.SummaryClosingTagSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(7, 1, 7, 6), details.BodySourceSpan);
        Assert.Equal(new MarkdownSourceSpan(8, 1, 8, 10), details.ClosingTagSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 9), Assert.Single(details.SyntaxNode.Children, child => child.Kind == MarkdownSyntaxKind.DetailsOpeningTag).SourceSpan);
        var summaryNode = Assert.Single(details.SyntaxNode.Children, child => child.Kind == MarkdownSyntaxKind.Summary);
        Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 9), Assert.Single(summaryNode.Children, child => child.Kind == MarkdownSyntaxKind.SummaryOpeningTag).SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 10, 5, 13), Assert.Single(summaryNode.Children, child => child.Kind == MarkdownSyntaxKind.SummaryText).SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 14, 5, 23), Assert.Single(summaryNode.Children, child => child.Kind == MarkdownSyntaxKind.SummaryClosingTag).SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(8, 1, 8, 10), Assert.Single(details.SyntaxNode.Children, child => child.Kind == MarkdownSyntaxKind.DetailsClosingTag).SourceSpan);

        var calloutBody = Assert.Single(native.EnumerateBlockSourceFields("calloutBody"));
        Assert.Same(callout, calloutBody.Block);
        Assert.Equal("Body line", calloutBody.Value);
        Assert.Equal(new MarkdownSourceSpan(2, 3, 2, 11), calloutBody.SourceSpan);

        var detailsOpeningTag = Assert.Single(native.EnumerateBlockSourceFields("detailsOpeningTag"));
        Assert.Same(details, detailsOpeningTag.Block);
        Assert.Equal("<details>", detailsOpeningTag.Value);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 9), detailsOpeningTag.SourceSpan);

        var summaryOpeningTag = Assert.Single(native.EnumerateBlockSourceFields("summaryOpeningTag"));
        Assert.Same(details, summaryOpeningTag.Block);
        Assert.Equal("<summary>", summaryOpeningTag.Value);
        Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 9), summaryOpeningTag.SourceSpan);

        var summaryText = Assert.Single(native.EnumerateBlockSourceFields("summaryText"));
        Assert.Same(details, summaryText.Block);
        Assert.Equal("More", summaryText.Value);
        Assert.Equal(new MarkdownSourceSpan(5, 10, 5, 13), summaryText.SourceSpan);

        var summaryClosingTag = Assert.Single(native.EnumerateBlockSourceFields("summaryClosingTag"));
        Assert.Same(details, summaryClosingTag.Block);
        Assert.Equal("</summary>", summaryClosingTag.Value);
        Assert.Equal(new MarkdownSourceSpan(5, 14, 5, 23), summaryClosingTag.SourceSpan);

        var detailsBody = Assert.Single(native.EnumerateBlockSourceFields("detailsBody"));
        Assert.Same(details, detailsBody.Block);
        Assert.Null(detailsBody.Value);
        Assert.Equal(new MarkdownSourceSpan(7, 1, 7, 6), detailsBody.SourceSpan);

        var detailsClosingTag = Assert.Single(native.EnumerateBlockSourceFields("detailsClosingTag"));
        Assert.Same(details, detailsClosingTag.Block);
        Assert.Equal("</details>", detailsClosingTag.Value);
        Assert.Equal(new MarkdownSourceSpan(8, 1, 8, 10), detailsClosingTag.SourceSpan);

        Assert.Equal("calloutBody", native.FindBlockSourceFieldAtPosition(2, 4)!.Name);
        Assert.Equal("detailsOpeningTag", native.FindBlockSourceFieldAtPosition(4, 3)!.Name);
        Assert.Equal("summaryOpeningTag", native.FindBlockSourceFieldAtPosition(5, 3)!.Name);
        Assert.Equal("summaryText", native.FindBlockSourceFieldAtPosition(5, 11)!.Name);
        Assert.Equal("summaryClosingTag", native.FindBlockSourceFieldAtPosition(5, 15)!.Name);
        Assert.Equal("detailsBody", native.FindBlockSourceFieldAtPosition(7, 3)!.Name);
        Assert.Equal("detailsClosingTag", native.FindBlockSourceFieldAtPosition(8, 3)!.Name);

        var snapshot = native.ToSnapshot();
        Assert.Equal(3, snapshot.Blocks[0].FieldSourceSpans["calloutBody"]!.StartColumn);
        Assert.Equal(11, snapshot.Blocks[0].FieldSourceSpans["calloutBody"]!.EndColumn);
        Assert.Equal(1, snapshot.Blocks[1].FieldSourceSpans["detailsOpeningTag"]!.StartColumn);
        Assert.Equal(9, snapshot.Blocks[1].FieldSourceSpans["detailsOpeningTag"]!.EndColumn);
        Assert.Equal(1, snapshot.Blocks[1].FieldSourceSpans["summaryOpeningTag"]!.StartColumn);
        Assert.Equal(9, snapshot.Blocks[1].FieldSourceSpans["summaryOpeningTag"]!.EndColumn);
        Assert.Equal(10, snapshot.Blocks[1].FieldSourceSpans["summaryText"]!.StartColumn);
        Assert.Equal(13, snapshot.Blocks[1].FieldSourceSpans["summaryText"]!.EndColumn);
        Assert.Equal(14, snapshot.Blocks[1].FieldSourceSpans["summaryClosingTag"]!.StartColumn);
        Assert.Equal(23, snapshot.Blocks[1].FieldSourceSpans["summaryClosingTag"]!.EndColumn);
        Assert.Equal(1, snapshot.Blocks[1].FieldSourceSpans["detailsBody"]!.StartColumn);
        Assert.Equal(6, snapshot.Blocks[1].FieldSourceSpans["detailsBody"]!.EndColumn);
        Assert.Equal(1, snapshot.Blocks[1].FieldSourceSpans["detailsClosingTag"]!.StartColumn);
        Assert.Equal(10, snapshot.Blocks[1].FieldSourceSpans["detailsClosingTag"]!.EndColumn);
        Assert.Contains(snapshot.Blocks[0].SourceFields, field => field.Name == "calloutBody" && field.Value == "Body line");
        Assert.Contains(snapshot.Blocks[1].SourceFields, field => field.Name == "detailsOpeningTag" && field.Value == "<details>");
        Assert.Contains(snapshot.Blocks[1].SourceFields, field => field.Name == "summaryOpeningTag" && field.Value == "<summary>");
        Assert.Contains(snapshot.Blocks[1].SourceFields, field => field.Name == "summaryText" && field.Value == "More");
        Assert.Contains(snapshot.Blocks[1].SourceFields, field => field.Name == "summaryClosingTag" && field.Value == "</summary>");
        Assert.Contains(snapshot.Blocks[1].SourceFields, field => field.Name == "detailsBody" && field.Value == null);
        Assert.Contains(snapshot.Blocks[1].SourceFields, field => field.Name == "detailsClosingTag" && field.Value == "</details>");

        Assert.Contains("> Updated", native.CreateReplaceEdit(calloutBody, "Updated").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("<details open>", native.CreateReplaceEdit(detailsOpeningTag, "<details open>").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("<summary class=\"lead\">", native.CreateReplaceEdit(summaryOpeningTag, "<summary class=\"lead\">").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("<summary>Less</summary>", native.CreateReplaceEdit(summaryText, "Less").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("</span>", native.CreateReplaceEdit(summaryClosingTag, "</span>").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("Outside", native.CreateReplaceEdit(detailsBody, "Outside").Apply(native.SourceMarkdown), StringComparison.Ordinal);
        Assert.Contains("</section>", native.CreateReplaceEdit(detailsClosingTag, "</section>").Apply(native.SourceMarkdown), StringComparison.Ordinal);
    }

    [Fact]
    public void Details_SummaryText_SourceField_Preserves_Source_Whitespace() {
        var markdown = """
<details>
<summary>  More context  </summary>
</details>
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var details = Assert.IsType<MarkdownNativeDetailsBlock>(Assert.Single(native.Blocks));
        var summaryText = Assert.Single(native.EnumerateBlockSourceFields("summaryText"));

        Assert.Equal("More context", details.Summary);
        Assert.Equal("  More context  ", summaryText.Value);
        Assert.Equal(new MarkdownSourceSpan(2, 10, 2, 25), summaryText.SourceSpan);
        Assert.Contains("<summary>Less</summary>", native.CreateReplaceEdit(summaryText, "Less").Apply(native.SourceMarkdown), StringComparison.Ordinal);
    }

    [Fact]
    public void Callout_LazyContinuation_Body_SourceField_Uses_Unquoted_And_Quoted_Line_Spans() {
        var markdown = """
> [!NOTE]
Lazy body
> quoted tail
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var callout = Assert.IsType<MarkdownNativeCalloutBlock>(Assert.Single(native.Blocks));

        Assert.Equal(new MarkdownSourceSpan(2, 1, 3, 13), callout.BodySourceSpan);

        var calloutBody = Assert.Single(native.EnumerateBlockSourceFields("calloutBody"));
        Assert.Same(callout, calloutBody.Block);
        Assert.Equal("Lazy body quoted tail", calloutBody.Value);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 3, 13), calloutBody.SourceSpan);

        Assert.Equal("calloutBody", native.FindBlockSourceFieldAtPosition(2, 5)!.Name);
        Assert.Equal("calloutBody", native.FindBlockSourceFieldAtPosition(3, 5)!.Name);

        Assert.True(native.TryCreateSourceSlice(calloutBody, out var normalizedSlice));
        Assert.Equal("Lazy body\n> quoted tail", normalizedSlice.Text.Replace("\r\n", "\n"));

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal(1, snapshot.FieldSourceSpans["calloutBody"]!.StartColumn);
        Assert.Equal(13, snapshot.FieldSourceSpans["calloutBody"]!.EndColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "calloutBody"
            && field.Value == "Lazy body quoted tail");

        var edited = native.CreateReplaceEdit(calloutBody, "Updated lazy body").Apply(native.SourceMarkdown);
        Assert.Contains("> [!NOTE]\nUpdated lazy body", edited.Replace("\r\n", "\n"), StringComparison.Ordinal);
    }

    [Fact]
    public void Quote_Body_SourceField_Uses_Structured_Child_Block_Span() {
        const string markdown = "> Old quote\n";

        var native = MarkdownNativeDocument.Parse(markdown);
        var quote = Assert.IsType<MarkdownNativeQuoteBlock>(Assert.Single(native.Blocks));

        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 11), quote.BodySourceSpan);

        var body = Assert.Single(native.EnumerateBlockSourceFields("quoteBody"));
        Assert.Same(quote, body.Block);
        Assert.Null(body.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 11), body.SourceSpan);

        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 5));
        Assert.Equal("quoteBody", found.Name);
        Assert.Equal(body.SourceSpan, found.SourceSpan);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal(3, snapshot.FieldSourceSpans["quoteBody"]!.StartColumn);
        Assert.Equal(11, snapshot.FieldSourceSpans["quoteBody"]!.EndColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "quoteBody"
            && field.Value == null
            && field.SourceSpan.StartColumn == 3
            && field.SourceSpan.EndColumn == 11);

        var edited = native.CreateReplaceEdit(body, "New quote").Apply(native.SourceMarkdown);
        Assert.Equal("> New quote", edited.TrimEnd('\r', '\n'));
    }

    [Fact]
    public void List_SourceFields_Use_Item_And_Task_Marker_Token_Spans() {
        var markdown = """
* [X]	Upper
* Plain

10) Ordered
11) Next
""";

        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var unordered = Assert.IsType<MarkdownNativeListBlock>(native.Blocks[0]);
        var ordered = Assert.IsType<MarkdownNativeListBlock>(native.Blocks[1]);

        Assert.Equal("*", unordered.Items[0].MarkerText);
        Assert.Equal("[X]", unordered.Items[0].TaskMarkerText);
        Assert.Equal("10)", ordered.Items[0].MarkerText);
        Assert.Equal("11)", ordered.Items[1].MarkerText);

        var listMarkers = native.EnumerateBlockSourceFields("listMarker").ToArray();
        Assert.Collection(
            listMarkers,
            field => {
                Assert.Same(unordered, field.Block);
                Assert.Equal("*", field.Value);
                Assert.Equal(0, field.Index);
                Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 1), field.SourceSpan);
            },
            field => {
                Assert.Same(unordered, field.Block);
                Assert.Equal("*", field.Value);
                Assert.Equal(1, field.Index);
                Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 1), field.SourceSpan);
            },
            field => {
                Assert.Same(ordered, field.Block);
                Assert.Equal("10)", field.Value);
                Assert.Equal(0, field.Index);
                Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 3), field.SourceSpan);
            },
            field => {
                Assert.Same(ordered, field.Block);
                Assert.Equal("11)", field.Value);
                Assert.Equal(1, field.Index);
                Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 3), field.SourceSpan);
            });

        var taskMarker = Assert.Single(native.EnumerateBlockSourceFields("taskMarker"));
        Assert.Same(unordered, taskMarker.Block);
        Assert.Equal("[X]", taskMarker.Value);
        Assert.Equal(0, taskMarker.Index);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 5), taskMarker.SourceSpan);

        var foundListMarker = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(4, 2));
        Assert.Equal("listMarker", foundListMarker.Name);
        Assert.Equal("10)", foundListMarker.Value);

        var foundTaskMarker = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 4));
        Assert.Equal("taskMarker", foundTaskMarker.Name);
        Assert.Equal("[X]", foundTaskMarker.Value);

        var snapshot = native.ToSnapshot();
        Assert.Contains(snapshot.Blocks[0].SourceFields, field =>
            field.Name == "listMarker"
            && field.Value == "*"
            && field.Index == 0
            && field.SourceSpan.StartColumn == 1);
        Assert.Contains(snapshot.Blocks[0].SourceFields, field =>
            field.Name == "taskMarker"
            && field.Value == "[X]"
            && field.Index == 0
            && field.SourceSpan.StartColumn == 3
            && field.SourceSpan.EndColumn == 5);
        Assert.Equal("*", snapshot.Blocks[0].Items[0].MarkerText);
        Assert.Equal("[X]", snapshot.Blocks[0].Items[0].TaskMarkerText);
        Assert.Equal("10)", snapshot.Blocks[1].Items[0].MarkerText);

        var taskEdited = native.CreateReplaceEdit(taskMarker, "[ ]").Apply(native.SourceMarkdown);
        Assert.StartsWith("* [ ]\tUpper", taskEdited, StringComparison.Ordinal);

        var orderedEdited = native.CreateReplaceEdit(listMarkers[2], "7.").Apply(native.SourceMarkdown);
        Assert.Contains("7. Ordered", orderedEdited, StringComparison.Ordinal);
        Assert.Contains("11) Next", orderedEdited, StringComparison.Ordinal);
    }

    [Fact]
    public void ThematicBreak_SourceFields_Use_Exact_Marker_Token_Spans() {
        const string markdown = "  * * *  \n\nAfter";

        var native = MarkdownNativeDocument.Parse(markdown, MarkdownReaderOptions.CreateCommonMarkProfile());
        var thematicBreak = Assert.IsType<MarkdownNativeThematicBreakBlock>(native.Blocks[0]);

        Assert.Equal("---", thematicBreak.Marker);
        Assert.Equal("* * *", thematicBreak.MarkerText);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 7), thematicBreak.MarkerSourceSpan);

        var marker = Assert.Single(native.EnumerateBlockSourceFields("marker"));
        Assert.Same(thematicBreak, marker.Block);
        Assert.Equal("* * *", marker.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 7), marker.SourceSpan);
        var foundMarker = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 5));
        Assert.Equal(marker.Name, foundMarker.Name);
        Assert.Equal(marker.Value, foundMarker.Value);
        Assert.Equal(marker.SourceSpan, foundMarker.SourceSpan);

        var snapshot = native.ToSnapshot();
        var snapshotField = Assert.Single(snapshot.Blocks[0].SourceFields);
        Assert.Equal("marker", snapshotField.Name);
        Assert.Equal("* * *", snapshotField.Value);
        Assert.Equal(3, snapshotField.SourceSpan.StartColumn);
        Assert.Equal(7, snapshotField.SourceSpan.EndColumn);

        var edited = native.CreateReplaceEdit(marker, "___").Apply(native.SourceMarkdown);
        Assert.StartsWith("  ___  ", edited, StringComparison.Ordinal);
    }

    [Fact]
    public void Image_SourceFields_Use_Image_Token_Spans() {
        const string markdown = "![Alt text](https://example.com/image.png \"Image title\")\n";

        var native = MarkdownNativeDocument.Parse(markdown);
        var image = Assert.IsType<MarkdownNativeImageBlock>(Assert.Single(native.Blocks));

        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 10), image.AltSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 13, 1, 41), image.SourceSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 44, 1, 54), image.TitleSourceSpan);

        var alt = Assert.Single(native.EnumerateBlockSourceFields("alt"));
        Assert.Same(image, alt.Block);
        Assert.Equal("Alt text", alt.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 10), alt.SourceSpan);

        var source = Assert.Single(native.EnumerateBlockSourceFields("source"));
        Assert.Equal("https://example.com/image.png", source.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 13, 1, 41), source.SourceSpan);

        var title = Assert.Single(native.EnumerateBlockSourceFields("title"));
        Assert.Equal("Image title", title.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 44, 1, 54), title.SourceSpan);

        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 45));
        Assert.Equal("title", found.Name);
        Assert.Equal("Image title", found.Value);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal(3, snapshot.FieldSourceSpans["alt"]!.StartColumn);
        Assert.Equal(41, snapshot.FieldSourceSpans["source"]!.EndColumn);
        Assert.Equal(44, snapshot.FieldSourceSpans["title"]!.StartColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "source"
            && field.Value == "https://example.com/image.png"
            && field.SourceSpan.StartColumn == 13
            && field.SourceSpan.EndColumn == 41);

        var edited = native.CreateReplaceEdit(source, "/media/new.png").Apply(native.SourceMarkdown);
        Assert.Equal("![Alt text](/media/new.png \"Image title\")", edited.TrimEnd('\r', '\n'));
    }

    [Fact]
    public void Linked_Image_SourceFields_Use_Image_And_Link_Token_Spans() {
        var markdown = """
[![Alt text](https://example.com/image.png "Image title")](https://example.com/docs "Link title")
_Caption_
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var image = Assert.IsType<MarkdownNativeImageBlock>(Assert.Single(native.Blocks));

        Assert.Equal(new MarkdownSourceSpan(1, 4, 1, 11), image.AltSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 14, 1, 42), image.SourceSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 45, 1, 55), image.TitleSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 60, 1, 83), image.LinkUrlSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 86, 1, 95), image.LinkTitleSourceSpan);

        var linkUrl = Assert.Single(native.EnumerateBlockSourceFields("linkUrl"));
        Assert.Same(image, linkUrl.Block);
        Assert.Equal("https://example.com/docs", linkUrl.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 60, 1, 83), linkUrl.SourceSpan);

        var linkTitle = Assert.Single(native.EnumerateBlockSourceFields("linkTitle"));
        Assert.Equal("Link title", linkTitle.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 86, 1, 95), linkTitle.SourceSpan);

        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 62));
        Assert.Equal("linkUrl", found.Name);
        Assert.Equal("https://example.com/docs", found.Value);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        Assert.Equal(60, snapshot.FieldSourceSpans["linkUrl"]!.StartColumn);
        Assert.Equal(95, snapshot.FieldSourceSpans["linkTitle"]!.EndColumn);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "linkTitle"
            && field.Value == "Link title"
            && field.SourceSpan.StartColumn == 86
            && field.SourceSpan.EndColumn == 95);

        var edited = native.CreateReplaceEdit(linkTitle, "More docs").Apply(native.SourceMarkdown);
        Assert.Contains("](https://example.com/docs \"More docs\")", edited, StringComparison.Ordinal);
        Assert.Contains("_Caption_", edited, StringComparison.Ordinal);
    }

    [Fact]
    public void FrontMatter_SourceFields_Use_Entry_Key_And_Value_Token_Spans() {
        var markdown = "--- \n" + """
title: Doc
published: true
tags: [a, b]
summary: |
  First line
  Second line
---

# Heading
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var frontMatter = Assert.IsType<MarkdownNativeFrontMatterBlock>(native.Blocks[0]);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 3), frontMatter.OpeningFenceSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 7, 13), frontMatter.BodySourceSpan);
        Assert.Equal(new MarkdownSourceSpan(8, 1, 8, 3), frontMatter.ClosingFenceSourceSpan);
        Assert.Contains("summary: |", frontMatter.RawYaml!, StringComparison.Ordinal);

        Assert.Collection(frontMatter.Entries,
            entry => {
                Assert.Equal("title", entry.Key);
                Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 5), entry.KeySourceSpan);
                Assert.Equal(new MarkdownSourceSpan(2, 8, 2, 10), entry.ValueSourceSpan);
            },
            entry => {
                Assert.Equal("published", entry.Key);
                Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 9), entry.KeySourceSpan);
                Assert.Equal(new MarkdownSourceSpan(3, 12, 3, 15), entry.ValueSourceSpan);
            },
            entry => {
                Assert.Equal("tags", entry.Key);
                Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 4), entry.KeySourceSpan);
                Assert.Equal(new MarkdownSourceSpan(4, 7, 4, 12), entry.ValueSourceSpan);
            },
            entry => {
                Assert.Equal("summary", entry.Key);
                Assert.Equal(new MarkdownSourceSpan(5, 1, 5, 7), entry.KeySourceSpan);
                Assert.Equal(new MarkdownSourceSpan(6, 3, 7, 13), entry.ValueSourceSpan);
            });

        var keys = native.EnumerateBlockSourceFields("frontMatterKey").ToArray();
        Assert.Equal(4, keys.Length);
        Assert.Equal("published", keys[1].Value);
        Assert.Equal(1, keys[1].Index);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 9), keys[1].SourceSpan);

        var values = native.EnumerateBlockSourceFields("frontMatterValue").ToArray();
        Assert.Equal(4, values.Length);
        Assert.Equal("true", values[1].Value);
        Assert.Equal("a, b", values[2].Value);
        Assert.Equal("First line\nSecond line", values[3].Value!.Replace("\r\n", "\n"));
        Assert.Equal(new MarkdownSourceSpan(6, 3, 7, 13), values[3].SourceSpan);

        var body = Assert.Single(native.EnumerateBlockSourceFields("frontMatterBody"));
        Assert.Contains("tags: [a, b]", body.Value!, StringComparison.Ordinal);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 7, 13), body.SourceSpan);

        var openingFence = Assert.Single(native.EnumerateBlockSourceFields("openingFence"));
        Assert.Equal("---", openingFence.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 3), openingFence.SourceSpan);

        var closingFence = Assert.Single(native.EnumerateBlockSourceFields("closingFence"));
        Assert.Equal("---", closingFence.Value);
        Assert.Equal(new MarkdownSourceSpan(8, 1, 8, 3), closingFence.SourceSpan);

        var found = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(4, 8));
        Assert.Equal("frontMatterValue", found.Name);
        Assert.Equal(2, found.Index);
        Assert.Equal("a, b", found.Value);

        var bodyFound = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(4, 5));
        Assert.Equal("frontMatterBody", bodyFound.Name);

        var openingFound = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(1, 2));
        Assert.Equal("openingFence", openingFound.Name);

        var closingFound = Assert.IsType<MarkdownNativeBlockSourceField>(native.FindBlockSourceFieldAtPosition(8, 2));
        Assert.Equal("closingFence", closingFound.Name);

        var snapshot = native.ToSnapshot().Blocks[0];
        Assert.Equal(frontMatter.RawYaml, snapshot.Fields["rawYaml"]);
        Assert.Equal(1, snapshot.FieldSourceSpans["openingFence"]!.StartLine);
        Assert.Equal(2, snapshot.FieldSourceSpans["frontMatterBody"]!.StartLine);
        Assert.Equal(8, snapshot.FieldSourceSpans["closingFence"]!.StartLine);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "frontMatterKey"
            && field.Value == "summary"
            && field.Index == 3
            && field.SourceSpan.StartLine == 5
            && field.SourceSpan.EndColumn == 7);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "frontMatterValue"
            && field.Value!.Replace("\r\n", "\n") == "First line\nSecond line"
            && field.Index == 3
            && field.SourceSpan.StartLine == 6
            && field.SourceSpan.StartColumn == 3);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "openingFence"
            && field.Value == "---"
            && field.SourceSpan.StartLine == 1
            && field.SourceSpan.EndColumn == 3);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "frontMatterBody"
            && field.Value!.Contains("summary: |", StringComparison.Ordinal)
            && field.SourceSpan.StartLine == 2
            && field.SourceSpan.EndLine == 7);
        Assert.Contains(snapshot.SourceFields, field =>
            field.Name == "closingFence"
            && field.Value == "---"
            && field.SourceSpan.StartLine == 8
            && field.SourceSpan.EndColumn == 3);

        var edited = native.CreateReplaceEdit(values[0], "Better").Apply(native.SourceMarkdown);
        Assert.Contains("title: Better", edited, StringComparison.Ordinal);
        Assert.Contains("published: true", edited, StringComparison.Ordinal);

        var fenceEdited = native.CreateReplaceEdit(openingFence, "...").Apply(native.SourceMarkdown);
        Assert.StartsWith("... \n", fenceEdited.Replace("\r\n", "\n"), StringComparison.Ordinal);

        var closingFenceEdited = native.CreateReplaceEdit(closingFence, "...").Apply(native.SourceMarkdown);
        Assert.Contains("\n...\n\n# Heading", closingFenceEdited.Replace("\r\n", "\n"), StringComparison.Ordinal);
    }

    [Fact]
    public void ToSnapshot_Projects_SourceFields_From_The_Same_Native_Block_Field_Enumeration() {
        var markdown = """
# Title

> [!NOTE] Heads up
> Body
> > Nested

<details>
<summary>More</summary>

Inside
</details>

[^note]: Footnote

```cs
Console.WriteLine();
```

| A | B |
| --- | --- |
| 1 | 2 |

> Quote
> Again

![Alt text](https://example.com/image.png "Image title")

---
""";

        var native = MarkdownNativeDocument.Parse(markdown);
        var blocks = native.DescendantBlocksAndSelf().ToArray();
        var snapshots = Flatten(native.ToSnapshot().Blocks).ToArray();

        Assert.Equal(blocks.Length, snapshots.Length);

        for (var blockIndex = 0; blockIndex < blocks.Length; blockIndex++) {
            Assert.Equal(blocks[blockIndex].Id, snapshots[blockIndex].Id);
            var fields = MarkdownNativeDocument.EnumerateBlockSourceFields(blocks[blockIndex]).ToArray();

            Assert.Equal(fields.Length, snapshots[blockIndex].SourceFields.Count);
            for (var fieldIndex = 0; fieldIndex < fields.Length; fieldIndex++) {
                var field = fields[fieldIndex];
                var snapshot = snapshots[blockIndex].SourceFields[fieldIndex];

                Assert.Equal(field.Name, snapshot.Name);
                Assert.Equal(field.Value, snapshot.Value);
                Assert.Equal(field.Index, snapshot.Index);
                Assert.Equal(field.SourceSpan.StartLine, snapshot.SourceSpan.StartLine);
                Assert.Equal(field.SourceSpan.StartColumn, snapshot.SourceSpan.StartColumn);
                Assert.Equal(field.SourceSpan.EndLine, snapshot.SourceSpan.EndLine);
                Assert.Equal(field.SourceSpan.EndColumn, snapshot.SourceSpan.EndColumn);
            }
        }

        var allFields = snapshots.SelectMany(block => block.SourceFields).ToArray();
        Assert.Contains(allFields, field => field.Name == "paragraphText" && field.Value == "Body");
        Assert.Contains(allFields, field => field.Name == "quoteMarker" && field.Index == 0 && field.SourceSpan.StartLine == 23);
        Assert.Contains(allFields, field => field.Name == "quoteMarker" && field.Index == 1 && field.SourceSpan.StartLine == 24);
        Assert.Contains(allFields, field => field.Name == "quoteBody" && field.SourceSpan.StartLine == 23);
        Assert.Contains(allFields, field => field.Name == "calloutBody" && field.Value!.Replace("\r\n", "\n") == "Body\n\n> Nested");
        Assert.Contains(allFields, field => field.Name == "detailsBody");
        Assert.Contains(allFields, field => field.Name == "label" && field.Value == "note");
        Assert.Contains(allFields, field => field.Name == "footnoteBody" && field.Value == "Footnote");
        Assert.Contains(allFields, field => field.Name == "marker" && field.Value == "---");
        Assert.Contains(allFields, field => field.Name == "source" && field.Value == "https://example.com/image.png");
    }

    private static IEnumerable<MarkdownNativeBlockSnapshot> Flatten(IReadOnlyList<MarkdownNativeBlockSnapshot> blocks) {
        for (var i = 0; i < blocks.Count; i++) {
            yield return blocks[i];

            foreach (var child in Flatten(blocks[i].Children)) {
                yield return child;
            }

            foreach (var item in blocks[i].Items) {
                foreach (var child in Flatten(item.Children)) {
                    yield return child;
                }
            }

            foreach (var cell in blocks[i].HeaderCells) {
                foreach (var child in Flatten(cell.Children)) {
                    yield return child;
                }
            }

            foreach (var row in blocks[i].Rows) {
                foreach (var cell in row) {
                    foreach (var child in Flatten(cell.Children)) {
                        yield return child;
                    }
                }
            }
        }
    }
}
