using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_GenericAttributes_Syntax_Tests {
    [Fact]
    public void ParseWithSyntaxTree_Captures_Block_GenericAttribute_Tokens() {
        const string markdown = "# Heading {#title .hero}\n\nAlpha paragraph {#intro .lead}\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var heading = result.FinalSyntaxTree.Children[0];
        var paragraph = result.FinalSyntaxTree.Children[1];
        var headingAttributes = Assert.Single(heading.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);
        var paragraphAttributes = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        Assert.Equal("{#title .hero}", headingAttributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 11, 1, 24), headingAttributes.SourceSpan);
        Assert.True(heading.SourceSpan!.Value.Contains(headingAttributes.SourceSpan!.Value));

        Assert.Equal("{#intro .lead}", paragraphAttributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(3, 17, 3, 30), paragraphAttributes.SourceSpan);
        Assert.True(paragraph.SourceSpan!.Value.Contains(paragraphAttributes.SourceSpan!.Value));

        Assert.Equal(MarkdownSyntaxKind.GenericAttributeBlock, result.FindDeepestFinalNodeAtPosition(1, 20)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.GenericAttributeBlock, result.FindDeepestFinalNodeAtPosition(3, 20)!.Kind);

        Assert.True(result.TryCreateOriginalSourceSlice(headingAttributes, out var headingSlice));
        Assert.Equal("{#title .hero}", headingSlice.Text);
        Assert.True(result.TryCreateOriginalSourceSlice(paragraphAttributes, out var paragraphSlice));
        Assert.Equal("{#intro .lead}", paragraphSlice.Text);
    }

    [Fact]
    public void Paragraph_GenericAttributes_Preserve_Consumed_Separator_Whitespace() {
        const string markdown = "Alpha paragraph  {#intro .lead}\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true
        };

        var document = MarkdownReader.Parse(markdown, options);
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));

        Assert.Equal("  ", paragraph.GenericAttributeConsumedWhitespace);
        Assert.Equal("Alpha paragraph  {#intro .lead}", ((IMarkdownBlock)paragraph).RenderMarkdown());
        Assert.Equal(
            "<p id=\"intro\" class=\"lead\">Alpha paragraph  </p>",
            document.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                EscapeNonAsciiText = false
            }));
    }

    [Fact]
    public void Standalone_GenericAttributes_Attach_To_Following_Heading_With_Source_Backup() {
        const string markdown = "{#intro .wide}\n# Heading\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var document = MarkdownReader.Parse(markdown, options);
        var headingBlock = Assert.IsType<HeadingBlock>(Assert.Single(document.Blocks));

        Assert.Equal("intro", headingBlock.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, headingBlock.Attributes.Classes);
        Assert.Equal("# Heading {#intro .wide}", ((IMarkdownBlock)headingBlock).RenderMarkdown());

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var heading = Assert.Single(result.FinalSyntaxTree.Children);
        var attributes = Assert.Single(heading.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);
        var headingText = Assert.Single(heading.Descendants(), node => node.Kind == MarkdownSyntaxKind.HeadingText);

        Assert.Equal(MarkdownSyntaxKind.Heading, heading.Kind);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 2, 9), heading.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 3, 2, 9), headingText.SourceSpan);
        Assert.Equal("{#intro .wide}", attributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 14), attributes.SourceSpan);
        Assert.True(result.TryCreateOriginalSourceSlice(attributes, out var slice));
        Assert.Equal("{#intro .wide}", slice.Text);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeHeading = Assert.IsType<MarkdownNativeHeadingBlock>(Assert.Single(native.Blocks));
        var field = Assert.Single(native.EnumerateBlockSourceFields("attributes"));

        Assert.Equal("Heading", nativeHeading.Text);
        Assert.Same(nativeHeading, field.Block);
        Assert.Equal("{#intro .wide}", field.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 14), field.SourceSpan);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(field, "{#docs .anchor}"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("{#docs .anchor}\n# Heading\n", roundtrip.Markdown);
    }

    [Fact]
    public void AtxHeading_GenericAttributes_After_ClosingMarker_Are_SourceBacked() {
        const string markdown = "# Heading # {#intro .wide}\n";
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.GenericAttributes = true;
        options.PreserveTrivia = true;

        var document = MarkdownReader.Parse(markdown, options);
        var headingBlock = Assert.IsType<HeadingBlock>(Assert.Single(document.Blocks));

        Assert.Equal("Heading", headingBlock.Text);
        Assert.Equal("intro", headingBlock.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, headingBlock.Attributes.Classes);
        Assert.Equal(new MarkdownSourceSpan(1, 11, 1, 11), headingBlock.ClosingMarkerSourceSpan);

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var heading = Assert.Single(result.FinalSyntaxTree.Children);
        var headingText = Assert.Single(heading.Descendants(), node => node.Kind == MarkdownSyntaxKind.HeadingText);
        var closingMarker = Assert.Single(heading.Children, node => node.Kind == MarkdownSyntaxKind.HeadingClosingMarker);
        var attributes = Assert.Single(heading.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        Assert.Equal("Heading", headingText.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 9), headingText.SourceSpan);
        Assert.Equal("#", closingMarker.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 11, 1, 11), closingMarker.SourceSpan);
        Assert.Equal("{#intro .wide}", attributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 13, 1, 26), attributes.SourceSpan);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeHeading = Assert.IsType<MarkdownNativeHeadingBlock>(Assert.Single(native.Blocks));
        var attributeField = Assert.Single(native.EnumerateBlockSourceFields("attributes"));
        var closingField = Assert.Single(native.EnumerateBlockSourceFields("closingMarker"));

        Assert.Same(nativeHeading, attributeField.Block);
        Assert.Equal("{#intro .wide}", attributeField.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 13, 1, 26), attributeField.SourceSpan);
        Assert.Same(nativeHeading, closingField.Block);
        Assert.Equal("#", closingField.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 11, 1, 11), closingField.SourceSpan);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(attributeField, "{#docs .anchor}"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("# Heading # {#docs .anchor}\n", roundtrip.Markdown);
    }

    [Fact]
    public void Standalone_GenericAttributes_Before_ReferenceDefinition_Create_Attributed_Paragraph() {
        const string markdown = "{#ref .wide}\n[id]: https://example.com\n\n[site][id]\n";
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.GenericAttributes = true;
        options.PreserveTrivia = true;

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
        Assert.Empty(result.ReferenceLinkDefinitions);

        var first = Assert.IsType<ParagraphBlock>(result.Document.Blocks[0]);
        var second = Assert.IsType<ParagraphBlock>(result.Document.Blocks[1]);

        Assert.Equal("ref", first.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, first.Attributes.Classes);
        Assert.Equal("[id]: https://example.com", InlinePlainText.Extract(first.Inlines));
        Assert.Equal("[site][id]", InlinePlainText.Extract(second.Inlines));

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(native.Blocks[0]);
        Assert.Empty(native.ReferenceLinkDefinitions);
        Assert.Equal("[id]: https://example.com", paragraph.Text);

        var attributes = Assert.Single(native.EnumerateBlockSourceFields("attributes"));
        Assert.Same(native.Blocks[0], attributes.Block);
        Assert.Equal("{#ref .wide}", attributes.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 12), attributes.SourceSpan);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(attributes, "{#literal .ref}"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("{#literal .ref}\n[id]: https://example.com\n\n[site][id]\n", roundtrip.Markdown);
    }

    [Fact]
    public void Standalone_GenericAttributes_Attach_To_Following_List_With_Source_Backup() {
        const string markdown = "{#list .wide}\n- item\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var document = MarkdownReader.Parse(markdown, options);
        var listBlock = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));

        Assert.Equal("list", listBlock.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, listBlock.Attributes.Classes);
        Assert.Equal(
            "{#list .wide}\n- item",
            ((IMarkdownBlock)listBlock).RenderMarkdown().Replace("\r\n", "\n"));

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var list = Assert.Single(result.FinalSyntaxTree.Children);
        var attributes = Assert.Single(list.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);
        var item = Assert.Single(list.Descendants(), node => node.Kind == MarkdownSyntaxKind.ListItem);

        Assert.Equal(MarkdownSyntaxKind.UnorderedList, list.Kind);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 2, 6), list.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 6), item.SourceSpan);
        Assert.Equal("{#list .wide}", attributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 13), attributes.SourceSpan);
        Assert.True(result.TryCreateOriginalSourceSlice(attributes, out var slice));
        Assert.Equal("{#list .wide}", slice.Text);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeList = Assert.IsType<MarkdownNativeListBlock>(Assert.Single(native.Blocks));
        var field = Assert.Single(native.EnumerateBlockSourceFields("attributes"));

        Assert.Same(nativeList, field.Block);
        Assert.Equal("{#list .wide}", field.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 13), field.SourceSpan);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(field, "{#docs .items}"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("{#docs .items}\n- item\n", roundtrip.Markdown);
    }

    [Fact]
    public void Standalone_GenericAttributes_Attach_To_Following_PipeTable_With_Source_Backup() {
        const string markdown = "{#tbl .wide}\n| A |\n|---|\n| B |\n";
        var options = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        options.GenericAttributes = true;
        options.PreserveTrivia = true;

        var document = MarkdownReader.Parse(markdown, options);
        var tableBlock = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));

        Assert.Equal("tbl", tableBlock.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, tableBlock.Attributes.Classes);
        Assert.Equal(
            "| A {#tbl .wide} |\n| --- |\n| B |",
            ((IMarkdownBlock)tableBlock).RenderMarkdown().Replace("\r\n", "\n"));
        Assert.Equal(
            "<table id=\"tbl\" class=\"wide\"><thead><tr><th>A</th></tr></thead><tbody><tr><td>B</td></tr></tbody></table>",
            document.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                EscapeNonAsciiText = false
            }));

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var table = Assert.Single(result.FinalSyntaxTree.Children);
        var attributes = Assert.Single(table.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);
        var header = Assert.Single(table.Children, node => node.Kind == MarkdownSyntaxKind.TableHeader);
        var alignment = Assert.Single(table.Children, node => node.Kind == MarkdownSyntaxKind.TableAlignmentRow);
        var row = Assert.Single(table.Children, node => node.Kind == MarkdownSyntaxKind.TableRow);

        Assert.Equal(MarkdownSyntaxKind.Table, table.Kind);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 4, 5), table.SourceSpan);
        Assert.Equal("{#tbl .wide}", attributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 12), attributes.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 3, 2, 3), Assert.Single(header.Children).SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 5), alignment.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 3, 4, 3), Assert.Single(row.Children).SourceSpan);
        Assert.True(result.TryCreateOriginalSourceSlice(attributes, out var slice));
        Assert.Equal("{#tbl .wide}", slice.Text);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeTable = Assert.IsType<MarkdownNativeTableBlock>(Assert.Single(native.Blocks));
        var field = Assert.Single(native.EnumerateBlockSourceFields("attributes"));

        Assert.Same(nativeTable, field.Block);
        Assert.Equal("{#tbl .wide}", field.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 12), field.SourceSpan);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(field, "{#docs .grid}"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("{#docs .grid}\n| A |\n|---|\n| B |\n", roundtrip.Markdown);
    }

    [Fact]
    public void Standalone_GenericAttributes_Attach_To_Following_FencedCode_With_Source_Backup() {
        const string markdown = "{#code .wide}\n```cs\nvar x = 1;\n```\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var document = MarkdownReader.Parse(markdown, options);
        var codeBlock = Assert.IsType<CodeBlock>(Assert.Single(document.Blocks));

        Assert.Equal("code", codeBlock.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, codeBlock.Attributes.Classes);
        Assert.Equal(
            "{#code .wide}\n```cs\nvar x = 1;\n```",
            ((IMarkdownBlock)codeBlock).RenderMarkdown().Replace("\r\n", "\n"));
        Assert.Equal(
            "<pre><code id=\"code\" class=\"language-cs wide\">var x = 1;\n</code></pre>",
            document.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                EscapeNonAsciiText = false
            }));

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var code = Assert.Single(result.FinalSyntaxTree.Children);
        var attributes = Assert.Single(code.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);
        var openingFence = Assert.Single(code.Children, node => node.Kind == MarkdownSyntaxKind.CodeFenceOpening);
        var info = Assert.Single(code.Children, node => node.Kind == MarkdownSyntaxKind.CodeFenceInfo);
        var content = Assert.Single(code.Children, node => node.Kind == MarkdownSyntaxKind.CodeContent);
        var closingFence = Assert.Single(code.Children, node => node.Kind == MarkdownSyntaxKind.CodeFenceClosing);

        Assert.Equal(MarkdownSyntaxKind.CodeBlock, code.Kind);
        Assert.Equal("{#code .wide}", attributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 13), attributes.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 3), openingFence.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 4, 2, 5), info.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 10), content.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 3), closingFence.SourceSpan);
        Assert.True(result.TryCreateOriginalSourceSlice(attributes, out var slice));
        Assert.Equal("{#code .wide}", slice.Text);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeCode = Assert.IsType<MarkdownNativeCodeBlock>(Assert.Single(native.Blocks));
        var field = Assert.Single(native.EnumerateBlockSourceFields("attributes"));

        Assert.Equal("code", nativeCode.ElementId);
        Assert.Equal(new[] { "wide" }, nativeCode.Classes);
        Assert.Same(nativeCode, field.Block);
        Assert.Equal("{#code .wide}", field.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 13), field.SourceSpan);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(field, "{#sample .snippet}"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("{#sample .snippet}\n```cs\nvar x = 1;\n```\n", roundtrip.Markdown);
    }

    [Fact]
    public void ListContained_Standalone_GenericAttributes_Attach_To_Following_FencedCode_With_Source_Backup() {
        const string markdown = "- item\n\n  {#code .wide}\n  ```cs\n  x\n  ```\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var document = MarkdownReader.Parse(markdown, options);
        var listBlock = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var item = Assert.Single(listBlock.Items);
        var codeBlock = Assert.IsType<CodeBlock>(Assert.Single(item.Children));

        Assert.Equal("code", codeBlock.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, codeBlock.Attributes.Classes);
        Assert.Equal(
            "<ul><li><p>item</p><pre><code id=\"code\" class=\"language-cs wide\">x\n</code></pre></li></ul>",
            document.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                EscapeNonAsciiText = false
            }));

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var code = Assert.Single(result.FinalSyntaxTree.Descendants(), node => node.Kind == MarkdownSyntaxKind.CodeBlock);
        var attributes = Assert.Single(code.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);
        var openingFence = Assert.Single(code.Children, node => node.Kind == MarkdownSyntaxKind.CodeFenceOpening);
        var info = Assert.Single(code.Children, node => node.Kind == MarkdownSyntaxKind.CodeFenceInfo);
        var content = Assert.Single(code.Children, node => node.Kind == MarkdownSyntaxKind.CodeContent);
        var closingFence = Assert.Single(code.Children, node => node.Kind == MarkdownSyntaxKind.CodeFenceClosing);

        Assert.Equal("{#code .wide}", attributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(3, 3, 3, 15), attributes.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 3, 4, 5), openingFence.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 6, 4, 7), info.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 3, 5, 3), content.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(6, 3, 6, 5), closingFence.SourceSpan);
        Assert.True(result.TryCreateOriginalSourceSlice(attributes, out var slice));
        Assert.Equal("{#code .wide}", slice.Text);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeList = Assert.IsType<MarkdownNativeListBlock>(Assert.Single(native.Blocks));
        var nativeCode = Assert.Single(Assert.Single(nativeList.Items).Children.OfType<MarkdownNativeCodeBlock>());
        var field = Assert.Single(native.EnumerateBlockSourceFields("attributes"));

        Assert.Equal("code", nativeCode.ElementId);
        Assert.Equal(new[] { "wide" }, nativeCode.Classes);
        Assert.Equal(new MarkdownSourceSpan(4, 3, 4, 5), nativeCode.OpeningFenceSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 6, 4, 7), nativeCode.InfoStringSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(5, 3, 5, 3), nativeCode.ContentSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(6, 3, 6, 5), nativeCode.ClosingFenceSourceSpan);
        Assert.Same(nativeCode, field.Block);
        Assert.Equal("{#code .wide}", field.Value);
        Assert.Equal(new MarkdownSourceSpan(3, 3, 3, 15), field.SourceSpan);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(field, "{#sample .snippet}"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("- item\n\n  {#sample .snippet}\n  ```cs\n  x\n  ```\n", roundtrip.Markdown);
    }

    [Fact]
    public void FencedCode_InfoString_GenericAttributes_Attach_To_Code_With_Source_Info() {
        const string markdown = "```{#code .wide}\nvar x = 1;\n```\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var document = MarkdownReader.Parse(markdown, options);
        var codeBlock = Assert.IsType<CodeBlock>(Assert.Single(document.Blocks));

        Assert.Equal(string.Empty, codeBlock.Language);
        Assert.Equal("{#code .wide}", codeBlock.InfoString);
        Assert.Equal("code", codeBlock.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, codeBlock.Attributes.Classes);
        Assert.Equal(
            "```{#code .wide}\nvar x = 1;\n```",
            ((IMarkdownBlock)codeBlock).RenderMarkdown().Replace("\r\n", "\n"));
        Assert.Equal(
            "<pre><code id=\"code\" class=\"wide\">var x = 1;\n</code></pre>",
            document.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                EscapeNonAsciiText = false
            }));

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var code = Assert.Single(result.FinalSyntaxTree.Children);
        var openingFence = Assert.Single(code.Children, node => node.Kind == MarkdownSyntaxKind.CodeFenceOpening);
        var info = Assert.Single(code.Children, node => node.Kind == MarkdownSyntaxKind.CodeFenceInfo);
        var content = Assert.Single(code.Children, node => node.Kind == MarkdownSyntaxKind.CodeContent);
        var closingFence = Assert.Single(code.Children, node => node.Kind == MarkdownSyntaxKind.CodeFenceClosing);

        Assert.Equal(MarkdownSyntaxKind.CodeBlock, code.Kind);
        Assert.DoesNotContain(code.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);
        Assert.Equal("{#code .wide}", info.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 3), openingFence.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 4, 1, 16), info.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 10), content.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 3), closingFence.SourceSpan);
        Assert.True(result.TryCreateOriginalSourceSlice(info, out var slice));
        Assert.Equal("{#code .wide}", slice.Text);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeCode = Assert.IsType<MarkdownNativeCodeBlock>(Assert.Single(native.Blocks));
        var infoField = Assert.Single(native.EnumerateBlockSourceFields("infoString"));
        var attributeField = Assert.Single(native.EnumerateBlockSourceFields("attributes"));
        var selectedField = native.FindBlockSourceFieldAtPosition(1, 4);

        Assert.Equal("code", nativeCode.ElementId);
        Assert.Equal(new[] { "wide" }, nativeCode.Classes);
        Assert.Same(nativeCode, infoField.Block);
        Assert.Equal("{#code .wide}", infoField.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 4, 1, 16), infoField.SourceSpan);
        Assert.Same(nativeCode, attributeField.Block);
        Assert.Equal("{#code .wide}", attributeField.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 4, 1, 16), attributeField.SourceSpan);
        Assert.Equal("attributes", selectedField?.Name);
        Assert.Equal(attributeField.SourceSpan, selectedField?.SourceSpan);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        var snapshotAttributes = Assert.Single(snapshot.EnumerateSourceFields("attributes"));
        Assert.Equal("{#code .wide}", snapshot.Fields["attributes"]);
        Assert.Equal(4, snapshot.FieldSourceSpans["attributes"]!.StartColumn);
        Assert.Equal(16, snapshot.FieldSourceSpans["attributes"]!.EndColumn);
        Assert.Equal("{#code .wide}", snapshotAttributes.Value);
        Assert.Equal(4, snapshotAttributes.SourceSpan.StartColumn);
        Assert.Equal(16, snapshotAttributes.SourceSpan.EndColumn);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(attributeField, "{#sample .snippet}"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("```{#sample .snippet}\nvar x = 1;\n```\n", roundtrip.Markdown);
    }

    [Fact]
    public void FencedCode_InfoString_GenericAttributes_With_Language_Render_On_Code() {
        const string markdown = "```cs {#code .wide}\nvar x = 1;\n```\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var document = MarkdownReader.Parse(markdown, options);
        var codeBlock = Assert.IsType<CodeBlock>(Assert.Single(document.Blocks));

        Assert.Equal("cs", codeBlock.Language);
        Assert.Equal("cs {#code .wide}", codeBlock.InfoString);
        Assert.Equal("code", codeBlock.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, codeBlock.Attributes.Classes);
        Assert.Equal(
            "<pre><code id=\"code\" class=\"wide language-cs\">var x = 1;\n</code></pre>",
            document.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                EscapeNonAsciiText = false
            }));

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeCode = Assert.IsType<MarkdownNativeCodeBlock>(Assert.Single(native.Blocks));
        var attributeField = Assert.Single(native.EnumerateBlockSourceFields("attributes"));
        var selectedField = native.FindBlockSourceFieldAtPosition(1, 7);

        Assert.Equal("cs", nativeCode.Language);
        Assert.Same(nativeCode, attributeField.Block);
        Assert.Equal("{#code .wide}", attributeField.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 7, 1, 19), attributeField.SourceSpan);
        Assert.Equal("attributes", selectedField?.Name);
        Assert.Equal(attributeField.SourceSpan, selectedField?.SourceSpan);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        var snapshotAttributes = Assert.Single(snapshot.EnumerateSourceFields("attributes"));
        Assert.Equal("{#code .wide}", snapshot.Fields["attributes"]);
        Assert.Equal(7, snapshot.FieldSourceSpans["attributes"]!.StartColumn);
        Assert.Equal(19, snapshot.FieldSourceSpans["attributes"]!.EndColumn);
        Assert.Equal("{#code .wide}", snapshotAttributes.Value);
        Assert.Equal(7, snapshotAttributes.SourceSpan.StartColumn);
        Assert.Equal(19, snapshotAttributes.SourceSpan.EndColumn);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(attributeField, "{#sample .snippet}"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("```cs {#sample .snippet}\nvar x = 1;\n```\n", roundtrip.Markdown);
    }

    [Fact]
    public void FencedCode_InfoString_GenericAttributes_Ignore_Opaque_Metadata_For_Code_Html() {
        const string markdown = "```cs linenums {#code .wide}\nvar x = 1;\n```\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var document = MarkdownReader.Parse(markdown, options);
        var codeBlock = Assert.IsType<CodeBlock>(Assert.Single(document.Blocks));

        Assert.Equal("cs", codeBlock.Language);
        Assert.Equal("cs linenums {#code .wide}", codeBlock.InfoString);
        Assert.True(codeBlock.FenceInfo.TryGetAttribute("linenums", out var linenums));
        Assert.Equal("true", linenums);
        Assert.Equal("code", codeBlock.FenceInfo.GenericAttributes.ElementId);
        Assert.Equal(new[] { "wide" }, codeBlock.FenceInfo.GenericAttributes.Classes);
        Assert.False(codeBlock.FenceInfo.GenericAttributes.TryGetAttribute("linenums", out _));
        Assert.Equal(
            "<pre><code id=\"code\" class=\"wide language-cs\">var x = 1;\n</code></pre>",
            document.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                EscapeNonAsciiText = false
            }));

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeCode = Assert.IsType<MarkdownNativeCodeBlock>(Assert.Single(native.Blocks));
        var infoField = Assert.Single(native.EnumerateBlockSourceFields("infoString"));
        var attributeField = Assert.Single(native.EnumerateBlockSourceFields("attributes"));
        var selectedField = native.FindBlockSourceFieldAtPosition(1, 16);

        Assert.Equal("cs", nativeCode.Language);
        Assert.Same(nativeCode, infoField.Block);
        Assert.Equal("cs linenums {#code .wide}", infoField.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 4, 1, 28), infoField.SourceSpan);
        Assert.Same(nativeCode, attributeField.Block);
        Assert.Equal("{#code .wide}", attributeField.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 16, 1, 28), attributeField.SourceSpan);
        Assert.Equal("attributes", selectedField?.Name);
        Assert.Equal(attributeField.SourceSpan, selectedField?.SourceSpan);

        var snapshot = Assert.Single(native.ToSnapshot().Blocks);
        var snapshotAttributes = Assert.Single(snapshot.EnumerateSourceFields("attributes"));
        Assert.Equal("{#code .wide}", snapshot.Fields["attributes"]);
        Assert.Equal(16, snapshot.FieldSourceSpans["attributes"]!.StartColumn);
        Assert.Equal(28, snapshot.FieldSourceSpans["attributes"]!.EndColumn);
        Assert.Equal("{#code .wide}", snapshotAttributes.Value);
        Assert.Equal(16, snapshotAttributes.SourceSpan.StartColumn);
        Assert.Equal(28, snapshotAttributes.SourceSpan.EndColumn);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(attributeField, "{#sample .snippet}"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("```cs linenums {#sample .snippet}\nvar x = 1;\n```\n", roundtrip.Markdown);
    }

    [Fact]
    public void Standalone_GenericAttributes_Attach_To_Following_ImageBlock_With_Source_Backup() {
        const string markdown = "{#img .wide}\n![Alt](image.png \"Title\")\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var document = MarkdownReader.Parse(markdown, options);
        var imageBlock = Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));

        Assert.Equal("img", imageBlock.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, imageBlock.Attributes.Classes);
        Assert.Equal("![Alt](image.png \"Title\"){#img .wide}", ((IMarkdownBlock)imageBlock).RenderMarkdown());
        Assert.Equal(
            "<img src=\"image.png\" alt=\"Alt\" title=\"Title\" id=\"img\" class=\"wide\" />",
            document.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                EscapeNonAsciiText = false
            }));

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var image = Assert.Single(result.FinalSyntaxTree.Children);
        var attributes = Assert.Single(image.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);
        var alt = Assert.Single(image.Children, node => node.Kind == MarkdownSyntaxKind.ImageAlt);
        var source = Assert.Single(image.Children, node => node.Kind == MarkdownSyntaxKind.ImageSource);
        var title = Assert.Single(image.Children, node => node.Kind == MarkdownSyntaxKind.ImageTitle);

        Assert.Equal(MarkdownSyntaxKind.Image, image.Kind);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 2, 25), image.SourceSpan);
        Assert.Equal("{#img .wide}", attributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 12), attributes.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 3, 2, 5), alt.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 8, 2, 16), source.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 19, 2, 23), title.SourceSpan);
        Assert.True(result.TryCreateOriginalSourceSlice(attributes, out var slice));
        Assert.Equal("{#img .wide}", slice.Text);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeImage = Assert.IsType<MarkdownNativeImageBlock>(Assert.Single(native.Blocks));
        var field = Assert.Single(native.EnumerateBlockSourceFields("attributes"));

        Assert.Same(nativeImage, field.Block);
        Assert.Equal("{#img .wide}", field.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 12), field.SourceSpan);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(field, "{#photo .hero}"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("{#photo .hero}\n![Alt](image.png \"Title\")\n", roundtrip.Markdown);
    }

    [Fact]
    public void BareUrl_Paragraph_GenericAttributes_Preserve_NoSpace_Source_And_Literal_Text() {
        const string markdown = "https://example.com{#auto .wide}\n";
        var options = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        options.GenericAttributes = true;
        options.PreserveTrivia = true;

        var document = MarkdownReader.Parse(markdown, options);
        var paragraphBlock = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));

        Assert.Equal("auto", paragraphBlock.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, paragraphBlock.Attributes.Classes);
        Assert.Equal("https://example.com{#auto .wide}", ((IMarkdownBlock)paragraphBlock).RenderMarkdown());

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var paragraph = Assert.Single(result.FinalSyntaxTree.Children);
        var attributes = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        Assert.Equal("{#auto .wide}", attributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 20, 1, 32), attributes.SourceSpan);
        Assert.True(result.TryCreateOriginalSourceSlice(attributes, out var slice));
        Assert.Equal("{#auto .wide}", slice.Text);
        Assert.DoesNotContain(
            result.FinalSyntaxTree.Descendants(),
            node => node.Kind == MarkdownSyntaxKind.InlineLink);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeParagraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var inline = Assert.Single(nativeParagraph.InlineRuns);
        var field = Assert.Single(native.EnumerateBlockSourceFields("attributes"));

        Assert.Equal("https://example.com", nativeParagraph.Text);
        Assert.Equal(MarkdownNativeInlineKind.Text, inline.Kind);
        Assert.Equal("https://example.com", inline.Text);
        Assert.Same(nativeParagraph, field.Block);
        Assert.Equal("{#auto .wide}", field.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 20, 1, 32), field.SourceSpan);
        Assert.Empty(native.EnumerateInlineMetadata("attributes"));
    }

    [Fact]
    public void Paragraph_StandaloneGenericAttributeContinuation_Is_Consumed_Without_Metadata() {
        const string markdown = "Paragraph\n{#literal .wide}\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var block = Assert.IsType<ParagraphBlock>(Assert.Single(result.Document.Blocks));
        Assert.True(block.Attributes.IsEmpty);
        Assert.Equal("Paragraph", InlinePlainText.Extract(block.Inlines));

        var paragraph = Assert.Single(result.FinalSyntaxTree.Children, node => node.Kind == MarkdownSyntaxKind.Paragraph);
        Assert.DoesNotContain(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeParagraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));

        Assert.Equal("Paragraph", nativeParagraph.Text);
        Assert.Empty(native.EnumerateBlockSourceFields("attributes"));
        Assert.Empty(native.EnumerateInlineMetadata("attributes"));
    }

    [Fact]
    public void PlainText_Paragraph_GenericAttributes_Preserve_NoSpace_Source_And_Target_Paragraph() {
        const string markdown = "word{#plain .wide}\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var document = MarkdownReader.Parse(markdown, options);
        var paragraphBlock = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));

        Assert.Equal("plain", paragraphBlock.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, paragraphBlock.Attributes.Classes);
        Assert.Equal("word{#plain .wide}", document.ToMarkdown(new MarkdownWriteOptions { OutputLineEnding = "\n" }).TrimEnd('\n'));

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var paragraph = Assert.Single(result.FinalSyntaxTree.Children, node => node.Kind == MarkdownSyntaxKind.Paragraph);
        var attributes = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        Assert.Equal("{#plain .wide}", attributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 18), attributes.SourceSpan);
        Assert.True(paragraph.SourceSpan!.Value.Contains(attributes.SourceSpan!.Value));
        Assert.True(result.TryCreateOriginalSourceSlice(attributes, out var slice));
        Assert.Equal("{#plain .wide}", slice.Text);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeParagraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var inline = Assert.Single(nativeParagraph.InlineRuns);
        var field = Assert.Single(native.EnumerateBlockSourceFields("attributes"));

        Assert.Equal("word", nativeParagraph.Text);
        Assert.Equal(MarkdownNativeInlineKind.Text, inline.Kind);
        Assert.Equal("word", inline.Text);
        Assert.Same(nativeParagraph, field.Block);
        Assert.Equal("{#plain .wide}", field.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 18), field.SourceSpan);
        Assert.Empty(native.EnumerateInlineMetadata("attributes"));
    }

    [Theory]
    [InlineData("text`{#plain .wide}\n", "text`", 6, 19)]
    [InlineData("text``{#plain .wide}\n", "text``", 7, 20)]
    [InlineData("`{#plain .wide}\n", "`", 2, 15)]
    [InlineData("``{#plain .wide}\n", "``", 3, 16)]
    public void UnmatchedBacktickRun_Paragraph_GenericAttributes_Preserve_NoSpace_Source_And_Target_Paragraph(
        string markdown,
        string expectedText,
        int attributeStartColumn,
        int attributeEndColumn) {
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var paragraph = Assert.Single(result.FinalSyntaxTree.Children, node => node.Kind == MarkdownSyntaxKind.Paragraph);
        var attributes = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        Assert.Equal("{#plain .wide}", attributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, attributeStartColumn, 1, attributeEndColumn), attributes.SourceSpan);
        Assert.True(paragraph.SourceSpan!.Value.Contains(attributes.SourceSpan!.Value));
        Assert.True(result.TryCreateOriginalSourceSlice(attributes, out var slice));
        Assert.Equal("{#plain .wide}", slice.Text);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeParagraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var field = Assert.Single(native.EnumerateBlockSourceFields("attributes"));

        Assert.Equal(expectedText, nativeParagraph.Text);
        Assert.Same(nativeParagraph, field.Block);
        Assert.Equal("{#plain .wide}", field.Value);
        Assert.Equal(new MarkdownSourceSpan(1, attributeStartColumn, 1, attributeEndColumn), field.SourceSpan);
        Assert.Empty(native.EnumerateInlineMetadata("attributes"));
    }

    [Theory]
    [InlineData("\\*{#esc .wide}\n", "*")]
    [InlineData("\\_{#esc .wide}\n", "_")]
    [InlineData("\\`{#esc .wide}\n", "`")]
    [InlineData("\\){#esc .wide}\n", ")")]
    [InlineData("\\]{#esc .wide}\n", "]")]
    public void EscapedPunctuation_Paragraph_GenericAttributes_Preserve_NoSpace_Source_And_Target_Paragraph(
        string markdown,
        string expectedText) {
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var paragraph = Assert.Single(result.FinalSyntaxTree.Children, node => node.Kind == MarkdownSyntaxKind.Paragraph);
        var attributes = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        Assert.Equal("{#esc .wide}", attributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 14), attributes.SourceSpan);
        Assert.True(paragraph.SourceSpan!.Value.Contains(attributes.SourceSpan!.Value));
        Assert.True(result.TryCreateOriginalSourceSlice(attributes, out var slice));
        Assert.Equal("{#esc .wide}", slice.Text);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeParagraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var field = Assert.Single(native.EnumerateBlockSourceFields("attributes"));

        Assert.Equal(expectedText, nativeParagraph.Text);
        Assert.Same(nativeParagraph, field.Block);
        Assert.Equal("{#esc .wide}", field.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 14), field.SourceSpan);
        Assert.Empty(native.EnumerateInlineMetadata("attributes"));
    }

    [Theory]
    [InlineData("&copy;{#e .wide}\n", "\u00A9{#e .wide}")]
    [InlineData("&#42;{#e .wide}\n", "*{#e .wide}")]
    [InlineData("&#x2A;{#e .wide}\n", "*{#e .wide}")]
    public void CharacterReference_Paragraph_GenericAttributes_Stay_Literal_Without_Metadata(
        string markdown,
        string expectedText) {
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var document = MarkdownReader.Parse(markdown, options);
        var block = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));

        Assert.True(block.Attributes.IsEmpty);
        Assert.Equal(expectedText, InlinePlainText.Extract(block.Inlines));

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var paragraph = Assert.Single(result.FinalSyntaxTree.Children, node => node.Kind == MarkdownSyntaxKind.Paragraph);

        Assert.DoesNotContain(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeParagraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));

        Assert.Equal(expectedText, nativeParagraph.Text);
        Assert.Empty(native.EnumerateBlockSourceFields("attributes"));
        Assert.Empty(native.EnumerateInlineMetadata("attributes"));
    }

    [Fact]
    public void AbbreviationEnding_Paragraph_GenericAttributes_Preserve_NoSpace_Source_And_Target_Paragraph() {
        const string markdown = "*[HTML]: Hyper Text Markup Language\n\nHTML{#abbr .wide}\n";
        var options = new MarkdownReaderOptions {
            Abbreviations = true,
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var paragraph = Assert.Single(result.FinalSyntaxTree.Children, node => node.Kind == MarkdownSyntaxKind.Paragraph);
        var paragraphAttributes = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);
        var abbreviation = Assert.Single(paragraph.Descendants(), node => node.Kind == MarkdownSyntaxKind.InlineAbbreviation);

        Assert.Equal("{#abbr .wide}", paragraphAttributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(3, 5, 3, 17), paragraphAttributes.SourceSpan);
        Assert.True(paragraph.SourceSpan!.Value.Contains(paragraphAttributes.SourceSpan!.Value));
        Assert.True(abbreviation.Attributes.IsEmpty);
        Assert.DoesNotContain(abbreviation.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        Assert.True(result.TryCreateOriginalSourceSlice(paragraphAttributes, out var slice));
        Assert.Equal("{#abbr .wide}", slice.Text);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeParagraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var field = Assert.Single(native.EnumerateBlockSourceFields("attributes"));

        Assert.Equal("HTML", nativeParagraph.Text);
        Assert.Same(nativeParagraph, field.Block);
        Assert.Equal("{#abbr .wide}", field.Value);
        Assert.Equal(new MarkdownSourceSpan(3, 5, 3, 17), field.SourceSpan);
        Assert.Empty(native.EnumerateInlineMetadata("attributes"));

        Assert.Equal("*[HTML]: Hyper Text Markup Language\n\nHTML{#abbr .wide}", result.Document.ToMarkdown(new MarkdownWriteOptions { OutputLineEnding = "\n" }).TrimEnd('\n'));
    }

    [Theory]
    [InlineData("[site{#txt .wide}](https://example.com)\n", "txt", "site", "{#txt .wide}", 1, 6, 1, 17)]
    [InlineData("![alt{#alt .wide}](img.png)\n", "alt", "alt", "{#alt .wide}", 1, 6, 1, 17)]
    public void NestedInlineContent_GenericAttributes_Promote_To_Paragraph(string markdown, string expectedId, string expectedText, string expectedSourceText, int startLine, int startColumn, int endLine, int endColumn) {
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.GenericAttributes = true;
        options.PreserveTrivia = true;

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var semanticParagraph = Assert.IsType<ParagraphBlock>(Assert.Single(result.Document.Blocks));
        Assert.Equal(expectedId, semanticParagraph.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, semanticParagraph.Attributes.Classes);
        Assert.Equal(expectedText, InlinePlainText.Extract(semanticParagraph.Inlines));

        var paragraph = Assert.Single(result.FinalSyntaxTree.Children, node => node.Kind == MarkdownSyntaxKind.Paragraph);
        var attributes = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);
        var expectedSpan = new MarkdownSourceSpan(startLine, startColumn, endLine, endColumn);

        Assert.Equal(expectedSourceText, attributes.Literal);
        Assert.Equal(expectedSpan, attributes.SourceSpan);
        Assert.True(paragraph.SourceSpan!.Value.Contains(attributes.SourceSpan!.Value));
        Assert.True(result.TryCreateOriginalSourceSlice(attributes, out var slice));
        Assert.Equal(expectedSourceText, slice.Text);

        Assert.DoesNotContain(result.FinalSyntaxTree.Descendants()
            .Where(node => node.Kind is MarkdownSyntaxKind.InlineLink or MarkdownSyntaxKind.InlineImage),
            node => node.Children.Any(child => child.Kind == MarkdownSyntaxKind.GenericAttributeBlock));

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeParagraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var field = Assert.Single(native.EnumerateBlockSourceFields("attributes"));

        Assert.Equal(expectedText, nativeParagraph.Text);
        Assert.Same(nativeParagraph, field.Block);
        Assert.Equal(expectedSourceText, field.Value);
        Assert.Equal(expectedSpan, field.SourceSpan);
        Assert.Empty(native.EnumerateInlineMetadata("attributes"));
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Inline_GenericAttribute_Tokens_Without_Duplicating_Native_Metadata() {
        const string markdown = "See [docs](old.md){#docs .primary} now\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var link = Assert.Single(result.FinalSyntaxTree.Descendants(), node => node.Kind == MarkdownSyntaxKind.InlineLink);
        var attributes = Assert.Single(link.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        Assert.Equal("{#docs .primary}", attributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 19, 1, 34), attributes.SourceSpan);
        Assert.True(link.SourceSpan!.Value.Contains(attributes.SourceSpan!.Value));
        Assert.Equal(MarkdownSyntaxKind.GenericAttributeBlock, result.FindDeepestFinalNodeAtPosition(1, 23)!.Kind);

        Assert.True(result.TryCreateOriginalSourceSlice(attributes, out var slice));
        Assert.Equal("{#docs .primary}", slice.Text);

        var trailingText = Assert.Single(
            result.FinalSyntaxTree.Descendants(),
            node => node.Kind == MarkdownSyntaxKind.InlineText && node.Literal == " now");
        Assert.Equal(new MarkdownSourceSpan(1, 35, 1, 38), trailingText.SourceSpan);
        Assert.True(result.TryCreateOriginalSourceSlice(trailingText, out var trailingTextSlice));
        Assert.Equal(" now", trailingTextSlice.Text);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeParagraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var nativeLink = Assert.Single(nativeParagraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Link);
        var nativeAttributes = Assert.Single(nativeLink.Metadata, metadata => metadata.Name == "attributes");
        var nativeTrailingText = Assert.Single(nativeParagraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Text && inline.Text == " now");

        Assert.Equal("{#docs .primary}", nativeAttributes.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 19, 1, 34), nativeAttributes.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 35, 1, 38), nativeTrailingText.SourceSpan);
        Assert.True(native.TryCreateOriginalSourceSlice(nativeTrailingText, out var nativeTrailingTextSlice));
        Assert.Equal(" now", nativeTrailingTextSlice.Text);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_Reference_Image_And_Autolink_GenericAttribute_Tokens() {
        const string markdown = "[site][id]{#lnk .primary} ![alt][img]{#img .wide} <https://example.com>{#auto .wide}\n\n[id]: https://example.com\n[img]: img.png\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var links = result.FinalSyntaxTree.Descendants()
            .Where(node => node.Kind == MarkdownSyntaxKind.InlineLink)
            .ToArray();
        var image = Assert.Single(result.FinalSyntaxTree.Descendants(), node => node.Kind == MarkdownSyntaxKind.InlineImage);
        var referenceLink = Assert.Single(links, node => node.Attributes.ElementId == "lnk");
        var angleAutolink = Assert.Single(links, node => node.Attributes.ElementId == "auto");

        AssertGenericAttributeToken(result, referenceLink, "{#lnk .primary}", new MarkdownSourceSpan(1, 11, 1, 25));
        AssertGenericAttributeToken(result, image, "{#img .wide}", new MarkdownSourceSpan(1, 38, 1, 49));
        AssertGenericAttributeToken(result, angleAutolink, "{#auto .wide}", new MarkdownSourceSpan(1, 72, 1, 84));

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(native.Blocks[0]);
        var nativeReferenceLink = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Link && inline.Text == "site");
        var nativeImage = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Image);
        var nativeAutolink = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Link && inline.Text == "https://example.com");

        Assert.Equal("{#lnk .primary}", Assert.Single(nativeReferenceLink.Metadata, metadata => metadata.Name == "attributes").Value);
        Assert.Equal("{#img .wide}", Assert.Single(nativeImage.Metadata, metadata => metadata.Name == "attributes").Value);
        Assert.Equal("{#auto .wide}", Assert.Single(nativeAutolink.Metadata, metadata => metadata.Name == "attributes").Value);
    }

    [Theory]
    [InlineData("^sup^{#sup .high} tail", MarkdownSyntaxKind.InlineSuperscript, MarkdownNativeInlineKind.Superscript, "{#sup .high}", 6, 17, "<p><sup id=\"sup\" class=\"high\">sup</sup> tail</p>")]
    [InlineData("~sub~{#sub .low} tail", MarkdownSyntaxKind.InlineSubscript, MarkdownNativeInlineKind.Subscript, "{#sub .low}", 6, 16, "<p><sub id=\"sub\" class=\"low\">sub</sub> tail</p>")]
    public void ParseWithSyntaxTree_Captures_EmphasisExtra_GenericAttribute_Source_Metadata_And_Writer(
        string markdown,
        MarkdownSyntaxKind syntaxKind,
        MarkdownNativeInlineKind nativeKind,
        string expectedLiteral,
        int startColumn,
        int endColumn,
        string expectedHtml) {
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true,
            Subscript = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var owner = Assert.Single(result.FinalSyntaxTree.Descendants(), node => node.Kind == syntaxKind);
        var expectedSpan = new MarkdownSourceSpan(1, startColumn, 1, endColumn);
        AssertGenericAttributeToken(result, owner, expectedLiteral, expectedSpan);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var inline = Assert.Single(paragraph.InlineRuns, run => run.Kind == nativeKind);
        var metadata = Assert.Single(inline.Metadata, item => item.Name == "attributes");

        Assert.Equal(expectedLiteral, metadata.Value);
        Assert.Equal(expectedSpan, metadata.SourceSpan);
        Assert.True(native.TryCreateOriginalSourceSlice(metadata, out var nativeSlice));
        Assert.Equal(expectedLiteral, nativeSlice.Text);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(metadata, "{#changed .edited}"));
        Assert.Equal(
            markdown.Replace(expectedLiteral, "{#changed .edited}", StringComparison.Ordinal),
            roundtrip.Markdown.TrimEnd('\r', '\n'));

        Assert.Equal(markdown, result.Document.ToMarkdown().TrimEnd('\r', '\n'));
        Assert.Equal(expectedHtml, result.Document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        }));
    }

    [Fact]
    public void ParseWithSyntaxTree_Renders_StrongEmphasis_GenericAttributes_Like_Markdig_Without_Duplicating_Markdown() {
        const string markdown = "***both***{#both .mix} tail";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var emphasis = Assert.Single(result.FinalSyntaxTree.Descendants(), node => node.Kind == MarkdownSyntaxKind.InlineEmphasis && node.Attributes.ElementId == "both");
        var strong = Assert.Single(emphasis.Children, node => node.Kind == MarkdownSyntaxKind.InlineStrong);
        var expectedSpan = new MarkdownSourceSpan(1, 11, 1, 22);

        AssertGenericAttributeToken(result, emphasis, "{#both .mix}", expectedSpan);
        Assert.True(strong.Attributes.IsEmpty);
        Assert.DoesNotContain(strong.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var nativeEmphasis = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Emphasis);
        var nativeStrong = Assert.Single(nativeEmphasis.Children, inline => inline.Kind == MarkdownNativeInlineKind.Strong);
        var emphasisAttributes = Assert.Single(nativeEmphasis.Metadata, item => item.Name == "attributes");

        Assert.Equal("{#both .mix}", emphasisAttributes.Value);
        Assert.Equal(expectedSpan, emphasisAttributes.SourceSpan);
        Assert.DoesNotContain(nativeStrong.Metadata, item => item.Name == "attributes");
        Assert.True(native.TryCreateOriginalSourceSlice(emphasisAttributes, out var emphasisSlice));
        Assert.Equal("{#both .mix}", emphasisSlice.Text);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(emphasisAttributes, "{#changed .edited}"));
        Assert.Equal("***both***{#changed .edited} tail", roundtrip.Markdown.TrimEnd('\r', '\n'));

        Assert.Equal(markdown, result.Document.ToMarkdown().TrimEnd('\r', '\n'));
        Assert.Equal("<p><em id=\"both\" class=\"mix\"><strong id=\"both\" class=\"mix\">both</strong></em> tail</p>", result.Document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        }));
    }

    [Fact]
    public void ParseWithSyntaxTree_Keeps_Strike_Highlight_And_Inserted_GenericAttributes_Literal() {
        const string markdown = "~~gone~~{#s .strike} ==mark=={#m .mark} ++ins++{#i .insert}\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        Assert.DoesNotContain(
            result.FinalSyntaxTree.Descendants(),
            node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);
        Assert.All(
            result.FinalSyntaxTree.Descendants().Where(node =>
                node.Kind == MarkdownSyntaxKind.InlineStrikethrough ||
                node.Kind == MarkdownSyntaxKind.InlineHighlight ||
                node.Kind == MarkdownSyntaxKind.InlineInserted),
            node => Assert.True(node.Attributes.IsEmpty));

        var native = MarkdownNativeDocument.Parse(markdown, options);
        Assert.Empty(native.EnumerateInlineMetadata("attributes"));
    }

    [Fact]
    public void FootnoteReference_GenericAttributes_Are_Consumed_Without_Metadata() {
        const string markdown = "See note[^a]{#ref .wide} tail\n\n[^a]: Footnote\n";
        var options = new MarkdownReaderOptions {
            Footnotes = true,
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
        Assert.DoesNotContain(
            result.FinalSyntaxTree.Descendants(),
            node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        var trailingText = Assert.Single(
            result.FinalSyntaxTree.Descendants(),
            node => node.Kind == MarkdownSyntaxKind.InlineText && node.Literal == " tail");
        Assert.Equal(new MarkdownSourceSpan(1, 25, 1, 29), trailingText.SourceSpan);
        Assert.True(result.TryCreateOriginalSourceSlice(trailingText, out var trailingTextSlice));
        Assert.Equal(" tail", trailingTextSlice.Text);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(native.Blocks[0]);
        var footnote = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.FootnoteRef);
        var nativeTrailingText = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Text && inline.Text == " tail");

        Assert.DoesNotContain(footnote.Metadata, metadata => metadata.Name == "attributes");
        Assert.Equal(new MarkdownSourceSpan(1, 25, 1, 29), nativeTrailingText.SourceSpan);
        Assert.True(native.TryCreateOriginalSourceSlice(nativeTrailingText, out var nativeTrailingTextSlice));
        Assert.Equal(" tail", nativeTrailingTextSlice.Text);
        Assert.DoesNotContain(
            "{#ref .wide}",
            result.Document.ToHtmlFragment(new HtmlOptions {
                Style = HtmlStyle.Plain,
                CssDelivery = CssDelivery.None,
                BodyClass = null,
                GitHubFootnoteHtml = true
            }),
            StringComparison.Ordinal);
    }

    [Fact]
    public void Standalone_GenericAttributes_Before_Footnote_Definition_Are_Consumed_Without_Metadata() {
        const string markdown = "{#fn .wide}\n[^a]: note\n\ntext[^a]\n";
        var options = new MarkdownReaderOptions {
            Footnotes = true,
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);
        Assert.DoesNotContain(
            result.FinalSyntaxTree.Descendants(),
            node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);
        Assert.DoesNotContain(
            result.FinalSyntaxTree.Children,
            node => node.Kind == MarkdownSyntaxKind.Paragraph && string.Equals(node.Literal, "{#fn .wide}", StringComparison.Ordinal));

        var native = MarkdownNativeDocument.Parse(markdown, options);
        Assert.Single(native.Blocks.OfType<MarkdownNativeFootnoteDefinitionBlock>());
        Assert.Empty(native.EnumerateBlockSourceFields("attributes"));
        Assert.Empty(native.EnumerateInlineMetadata("attributes"));

        var html = result.Document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false,
            GitHubFootnoteHtml = true
        });
        Assert.DoesNotContain("{#fn .wide}", html, StringComparison.Ordinal);
        Assert.DoesNotContain("id=\"fn\" class=\"wide\"", html, StringComparison.Ordinal);

        Assert.Equal("[^a]: note\n\ntext[^a]", result.Document.ToMarkdown(new MarkdownWriteOptions { OutputLineEnding = "\n" }).TrimEnd('\n'));
    }

    [Fact]
    public void ParseWithSyntaxTree_Keeps_Blockquote_Block_GenericAttributes_Literal() {
        const string markdown = "> quote {#q .lead}\n> # Heading {#h .wide}\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        Assert.DoesNotContain(
            result.FinalSyntaxTree.Descendants(),
            node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        var html = result.Document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        });

        Assert.Contains("quote {#q .lead}", html, StringComparison.Ordinal);
        Assert.Contains("Heading {#h .wide}", html, StringComparison.Ordinal);
        Assert.DoesNotContain("id=\"q\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("id=\"h\"", html, StringComparison.Ordinal);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        Assert.Empty(native.EnumerateBlockSourceFields("attributes"));
    }

    [Fact]
    public void Standalone_GenericAttributes_Before_Blockquote_Remain_Literal_Paragraph() {
        const string markdown = "{#q .wide}\n> quote\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        Assert.Collection(
            result.Document.Blocks,
            block => {
                var paragraph = Assert.IsType<ParagraphBlock>(block);
                Assert.Equal("{#q .wide}", paragraph.Inlines.RenderMarkdown());
                Assert.True(paragraph.Attributes.IsEmpty);
            },
            block => Assert.IsType<QuoteBlock>(block));

        Assert.DoesNotContain(
            result.FinalSyntaxTree.Descendants(),
            node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);
        Assert.Empty(MarkdownNativeDocument.Parse(markdown, options).EnumerateBlockSourceFields("attributes"));

        var html = result.Document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        });

        Assert.Equal("<p>{#q .wide}</p><blockquote><p>quote</p></blockquote>", html);
    }

    [Fact]
    public void Standalone_GenericAttributes_Before_HtmlBlock_Are_Consumed_Without_Metadata() {
        const string markdown = "{#html .wide}\n<div>raw</div>\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true,
            HtmlBlocks = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        Assert.IsType<HtmlRawBlock>(Assert.Single(result.Document.Blocks));
        Assert.DoesNotContain(
            result.FinalSyntaxTree.Descendants(),
            node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);
        Assert.Empty(MarkdownNativeDocument.Parse(markdown, options).EnumerateBlockSourceFields("attributes"));

        var html = result.Document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        });

        Assert.Equal("<div>raw</div>", html);
    }

    [Fact]
    public void Standalone_GenericAttributes_Before_ThematicBreak_Create_Attributed_Empty_SetextHeading() {
        const string markdown = "{#rule .wide}\n---\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var headingBlock = Assert.IsType<HeadingBlock>(Assert.Single(result.Document.Blocks));
        Assert.Equal(2, headingBlock.Level);
        Assert.Equal(string.Empty, headingBlock.Text);
        Assert.Equal("rule", headingBlock.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, headingBlock.Attributes.Classes);

        var heading = Assert.Single(result.FinalSyntaxTree.Children);
        var attributes = Assert.Single(heading.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);
        var underline = Assert.Single(heading.Children, node => node.Kind == MarkdownSyntaxKind.HeadingSetextUnderlineMarker);

        Assert.Equal(MarkdownSyntaxKind.Heading, heading.Kind);
        Assert.Equal("{#rule .wide}", attributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 13), attributes.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(2, 1, 2, 3), underline.SourceSpan);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeHeading = Assert.IsType<MarkdownNativeHeadingBlock>(Assert.Single(native.Blocks));
        var field = Assert.Single(native.EnumerateBlockSourceFields("attributes"));

        Assert.Equal(string.Empty, nativeHeading.Text);
        Assert.Same(nativeHeading, field.Block);
        Assert.Equal("{#rule .wide}", field.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 13), field.SourceSpan);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(field, "{#line .thin}"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("{#line .thin}\n---\n", roundtrip.Markdown);
    }

    [Fact]
    public void Standalone_GenericAttributes_Before_IndentedCode_Create_Attributed_Paragraph() {
        const string markdown = "{#code .wide}\n    var x = 1;\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true,
            IndentedCodeBlocks = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var paragraphBlock = Assert.IsType<ParagraphBlock>(Assert.Single(result.Document.Blocks));
        Assert.Equal("var x = 1;", paragraphBlock.Inlines.RenderMarkdown());
        Assert.Equal("code", paragraphBlock.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, paragraphBlock.Attributes.Classes);

        var paragraph = Assert.Single(result.FinalSyntaxTree.Children);
        var attributes = Assert.Single(paragraph.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        Assert.Equal(MarkdownSyntaxKind.Paragraph, paragraph.Kind);
        Assert.Equal("{#code .wide}", attributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 13), attributes.SourceSpan);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var nativeParagraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var field = Assert.Single(native.EnumerateBlockSourceFields("attributes"));

        Assert.Equal("var x = 1;", nativeParagraph.Text);
        Assert.Same(nativeParagraph, field.Block);
        Assert.Equal("{#code .wide}", field.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 13), field.SourceSpan);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(field, "{#sample .snippet}"));

        Assert.True(roundtrip.IsLossless);
        Assert.Empty(roundtrip.Diagnostics);
        Assert.Equal("{#sample .snippet}\n    var x = 1;\n", roundtrip.Markdown);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_ListItem_GenericAttribute_Tokens() {
        const string markdown = "- item {#li .selected}\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var listItem = Assert.Single(result.FinalSyntaxTree.Descendants(), node => node.Kind == MarkdownSyntaxKind.ListItem);
        var attributes = Assert.Single(listItem.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        Assert.Equal("{#li .selected}", attributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 8, 1, 22), attributes.SourceSpan);
        Assert.True(listItem.SourceSpan!.Value.Contains(attributes.SourceSpan!.Value));
        Assert.Equal(MarkdownSyntaxKind.GenericAttributeBlock, result.FindDeepestFinalNodeAtPosition(1, 12)!.Kind);

        Assert.True(result.TryCreateOriginalSourceSlice(attributes, out var slice));
        Assert.Equal("{#li .selected}", slice.Text);
    }

    [Fact]
    public void ListItem_GenericAttributes_Preserve_Consumed_Separator_Whitespace_In_Markdown_Writer() {
        const string markdown = "- item  {#li .selected}\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var document = MarkdownReader.Parse(markdown, options);
        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var item = Assert.Single(list.Items);

        Assert.Equal("  ", item.GenericAttributeConsumedWhitespace);
        Assert.Equal(
            "- item  {#li .selected}",
            ((IMarkdownBlock)list).RenderMarkdown().Replace("\r\n", "\n"));
        Assert.Equal(
            "- item  {#li .selected}\n",
            document.ToMarkdown(new MarkdownWriteOptions { OutputLineEnding = "\n" }));
    }

    [Fact]
    public void ListItem_Heading_GenericAttributes_Remain_Literal_While_FencedCode_Attributes_Stay_SourceBacked() {
        const string markdown = "- # Heading {#h .wide}\n- ```cs {#code .wide}\n  x\n  ```\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var heading = Assert.Single(result.FinalSyntaxTree.Descendants(), node => node.Kind == MarkdownSyntaxKind.Heading);
        var headingText = Assert.Single(heading.Descendants(), node => node.Kind == MarkdownSyntaxKind.HeadingText);

        Assert.Equal("Heading {#h .wide}", headingText.Literal);
        Assert.DoesNotContain(heading.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        var code = Assert.Single(result.FinalSyntaxTree.Descendants(), node => node.Kind == MarkdownSyntaxKind.CodeBlock);

        Assert.Equal("code", code.Attributes.ElementId);
        Assert.Equal(new[] { "wide" }, code.Attributes.Classes);

        var html = result.Document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        });

        Assert.Contains("<h1>Heading {#h .wide}</h1>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("id=\"h\"", html, StringComparison.Ordinal);
        Assert.Contains("id=\"code\"", html, StringComparison.Ordinal);
        Assert.Contains("wide", html, StringComparison.Ordinal);
        Assert.Contains("language-cs", html, StringComparison.Ordinal);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var list = Assert.IsType<MarkdownNativeListBlock>(Assert.Single(native.Blocks));
        var nativeHeading = Assert.IsType<MarkdownNativeHeadingBlock>(Assert.Single(list.Items[0].Children));
        var nativeCode = Assert.IsType<MarkdownNativeCodeBlock>(Assert.Single(list.Items[1].Children));
        var fields = native.EnumerateBlockSourceFields("attributes").ToArray();
        var codeAttributeField = Assert.Single(fields);

        Assert.Equal("Heading {#h .wide}", nativeHeading.Text);
        Assert.True(nativeHeading.Heading.Attributes.IsEmpty);
        Assert.Equal("code", nativeCode.ElementId);
        Assert.Equal(new[] { "wide" }, nativeCode.Classes);
        Assert.Same(nativeCode, codeAttributeField.Block);
        Assert.Equal("{#code .wide}", codeAttributeField.Value);
    }

    [Fact]
    public void DefinitionListTerm_GenericAttributes_Are_SourceBacked() {
        const string markdown = "Term {#term .wide}\n:   Definition {#def .wide}\n";
        var options = new MarkdownReaderOptions {
            DefinitionLists = true,
            GenericAttributes = true,
            PreserveTrivia = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var definitionTerm = Assert.Single(result.FinalSyntaxTree.Descendants(), node => node.Kind == MarkdownSyntaxKind.DefinitionTerm);
        var termAttributes = Assert.Single(definitionTerm.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        Assert.Equal("{#term .wide}", termAttributes.Literal);
        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 18), termAttributes.SourceSpan);
        Assert.True(definitionTerm.SourceSpan!.Value.Contains(termAttributes.SourceSpan!.Value));
        Assert.Equal(MarkdownSyntaxKind.GenericAttributeBlock, result.FindDeepestFinalNodeAtPosition(1, 10)!.Kind);

        Assert.True(result.TryCreateOriginalSourceSlice(termAttributes, out var slice));
        Assert.Equal("{#term .wide}", slice.Text);

        var native = MarkdownNativeDocument.Parse(markdown, options);
        var definitionList = Assert.IsType<MarkdownNativeDefinitionListBlock>(Assert.Single(native.Blocks));
        var group = Assert.Single(definitionList.Groups);
        var term = Assert.Single(group.Terms);

        Assert.Equal("Term", term.Text);
        Assert.Equal("Term {#term .wide}", term.Markdown);

        var attributes = native.EnumerateBlockSourceFields("attributes").ToArray();
        var nativeTermAttributes = Assert.Single(
            attributes,
            field => field.Block == definitionList && field.Index == 0 && field.Value == "{#term .wide}");

        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 18), nativeTermAttributes.SourceSpan);

        var roundtrip = native.WriteWithSourceEdit(native.CreateReplaceEdit(nativeTermAttributes, "{#label .tag}"));

        Assert.Contains("Term {#label .tag}", roundtrip.Markdown, StringComparison.Ordinal);
        Assert.Contains(":   Definition {#def .wide}", roundtrip.Markdown, StringComparison.Ordinal);
    }

    private static void AssertGenericAttributeToken(
        MarkdownParseResult result,
        MarkdownSyntaxNode owner,
        string expectedLiteral,
        MarkdownSourceSpan expectedSpan) {
        var attributes = Assert.Single(owner.Children, node => node.Kind == MarkdownSyntaxKind.GenericAttributeBlock);

        Assert.Equal(expectedLiteral, attributes.Literal);
        Assert.Equal(expectedSpan, attributes.SourceSpan);
        if (owner.SourceSpan.HasValue) {
            Assert.True(owner.SourceSpan.Value.Contains(attributes.SourceSpan!.Value));
        }

        Assert.True(result.TryCreateOriginalSourceSlice(attributes, out var slice));
        Assert.Equal(expectedLiteral, slice.Text);
    }
}
