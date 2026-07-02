using System;
using System.Linq;
using OfficeIMO.Markdown;
using OfficeIMO.MarkdownRenderer;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_CurrentHead_Review_Tests {
    [Fact]
    public void ParseWithSyntaxTree_Preserves_TransformDiagnostics() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.PreserveTrivia = true;
        options.DocumentTransforms.Add(new MarkdownCompactHeadingBoundaryTransform());

        var result = MarkdownReader.ParseWithSyntaxTree("previous shutdown was unexpected### Reason", options);

        var diagnostic = Assert.Single(result.TransformDiagnostics);
        Assert.Contains(nameof(MarkdownCompactHeadingBoundaryTransform), diagnostic.TransformName);
    }

    [Fact]
    public void MarkdownToRtf_Does_Not_Duplicate_Loose_List_AdditionalParagraphs() {
        const string markdown = """
            - Lead

              Continuation
            """;

        var document = markdown.ToRtfDocumentFromMarkdown();

        Assert.Equal(1, document.Paragraphs.Count(paragraph => paragraph.ToPlainText() == "Continuation"));
    }

    [Fact]
    public void MarkdownToRtf_Applies_Superscript_And_Subscript_Sequence_Styles() {
        var options = new MarkdownToRtfOptions {
            ReaderOptions = MarkdownReaderOptions.CreateOfficeIMOProfile()
        };

        var document = "Formula ^2^ and H~2~O".ToRtfDocumentFromMarkdown(options);
        var paragraph = Assert.Single(document.Paragraphs);

        Assert.Contains(paragraph.Runs, run => run.Text == "2" && run.VerticalPosition == RtfVerticalPosition.Superscript);
        Assert.Contains(paragraph.Runs, run => run.Text == "2" && run.VerticalPosition == RtfVerticalPosition.Subscript);
    }

    [Fact]
    public void Sequence_RenderMarkdown_Escapes_Own_Delimiters_In_Nested_Text() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();

        var document = MarkdownReader.Parse("++a\\+\\+b++\n\n^x\\^y^\n\n~h\\~i~", options);

        Assert.Collection(
            document.Blocks.Cast<ParagraphBlock>(),
            paragraph => Assert.Equal("++a\\+\\+b++", paragraph.Inlines.RenderMarkdown()),
            paragraph => Assert.Equal("^x\\^y^", paragraph.Inlines.RenderMarkdown()),
            paragraph => Assert.Equal("~h\\~i~", paragraph.Inlines.RenderMarkdown()));
    }

    [Fact]
    public void Sequence_RenderMarkdown_Escapes_Own_Delimiters_In_Nested_Link_And_Image_Text() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();

        var document = MarkdownReader.Parse("++[a\\+\\+b](u)++\n\n^![a\\^b](i)^\n\n~[![a\\~b](i)](u)~", options);

        Assert.Collection(
            document.Blocks.Cast<ParagraphBlock>(),
            paragraph => Assert.Equal("++[a\\+\\+b](u)++", paragraph.Inlines.RenderMarkdown()),
            paragraph => Assert.Equal("^![a\\^b](i)^", paragraph.Inlines.RenderMarkdown()),
            paragraph => Assert.Equal("~[![a\\~b](i)](u)~", paragraph.Inlines.RenderMarkdown()));
    }

    [Fact]
    public void NoPipe_Table_Body_Does_Not_Terminate_On_Disabled_Heading_Syntax() {
        var options = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        options.Headings = false;

        var document = MarkdownReader.Parse("""
            | A |
            |---|
            # value
            """, options);

        var table = Assert.IsType<TableBlock>(Assert.Single(document.Blocks));
        var row = Assert.Single(table.Rows);
        Assert.Equal("# value", Assert.Single(row));
    }

    [Fact]
    public void Native_Inline_SourceSlices_After_Tabs_Use_Visual_Columns() {
        var native = MarkdownNativeDocument.Parse("a\t`b`\n", new MarkdownReaderOptions {
            PreserveTrivia = true
        });
        var paragraph = Assert.IsType<MarkdownNativeParagraphBlock>(Assert.Single(native.Blocks));
        var code = Assert.Single(paragraph.InlineRuns, inline => inline.Kind == MarkdownNativeInlineKind.Code);

        Assert.True(native.TryCreateOriginalSourceSlice(code, out var slice));
        Assert.Equal("`b`", slice.Text);
    }

    [Fact]
    public void MarkdownRenderer_ParseDocument_FastPath_Provides_Syntax_Context_To_Transforms() {
        var sourceSliceCreated = false;
        var sourceText = string.Empty;
        var options = new MarkdownRendererOptions();
        options.DocumentTransforms.Add(new RendererInspectTransform((document, context) => {
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));

            Assert.NotNull(context.SyntaxTree);
            Assert.NotNull(context.TopLevelBlockSourceSpans);
            sourceSliceCreated = context.TryCreateSourceSlice(paragraph, out var slice);
            sourceText = slice.Text;
            return document;
        }));

        OfficeIMO.MarkdownRenderer.MarkdownRenderer.ParseDocument("Alpha\n", options);

        Assert.True(sourceSliceCreated);
        Assert.Equal("Alpha", sourceText);
    }

    [Fact]
    public void MarkdownRenderer_RenderBodyHtml_Provides_Syntax_Context_To_Transforms() {
        var sourceSliceCreated = false;
        var sourceText = string.Empty;
        var options = new MarkdownRendererOptions();
        options.DocumentTransforms.Add(new RendererInspectTransform((document, context) => {
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));

            Assert.NotNull(context.SyntaxTree);
            Assert.NotNull(context.TopLevelBlockSourceSpans);
            sourceSliceCreated = context.TryCreateSourceSlice(paragraph, out var slice);
            sourceText = slice.Text;
            return document;
        }));

        OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("Alpha\n", options);

        Assert.True(sourceSliceCreated);
        Assert.Equal("Alpha", sourceText);
    }

    [Fact]
    public void MarkdownRenderer_ParseDocumentResult_Attaches_Final_ParseResult_To_Transformed_Document() {
        var options = new MarkdownRendererOptions();
        options.DocumentTransforms.Add(new RendererInspectTransform((document, context) => {
            var replacement = MarkdownDoc.Create();
            replacement.Add(new ParagraphBlock(new InlineSequence().Text("Beta")));
            return replacement;
        }));

        var result = OfficeIMO.MarkdownRenderer.MarkdownRenderer.ParseDocumentResult("Alpha\n", options);

        Assert.NotNull(result.Document.ParseResult);
        Assert.Same(result.FinalSyntaxTree, result.Document.ParseResult!.FinalSyntaxTree);
        Assert.Same(result.Document, result.Document.ParseResult.Document);
    }

    [Fact]
    public void GenericAttributes_RenderMarkdown_Uses_KeyForm_For_OneCharacter_Id() {
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.GenericAttributes = true;

        var document = MarkdownReader.Parse("Alpha {id=h}", options);
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));

        var rendered = ((IMarkdownBlock)paragraph).RenderMarkdown();
        Assert.Equal("Alpha {id=\"h\"}", rendered);

        var reparsed = MarkdownReader.Parse(rendered, options);
        var reparsedParagraph = Assert.IsType<ParagraphBlock>(Assert.Single(reparsed.Blocks));
        Assert.Equal("h", reparsedParagraph.Attributes?.ElementId);
    }

    [Fact]
    public void GenericAttributes_Unescapes_Quoted_Attribute_Values() {
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.GenericAttributes = true;

        var document = MarkdownReader.Parse("Alpha {title=\"a\\\"b\" data-path=\"c\\\\d\"}", options);
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));

        Assert.Equal("a\"b", paragraph.Attributes.Attributes["title"]);
        Assert.Equal("c\\d", paragraph.Attributes.Attributes["data-path"]);

        var rendered = ((IMarkdownBlock)paragraph).RenderMarkdown();
        var reparsed = MarkdownReader.Parse(rendered, options);
        var reparsedParagraph = Assert.IsType<ParagraphBlock>(Assert.Single(reparsed.Blocks));

        Assert.Equal("a\"b", reparsedParagraph.Attributes.Attributes["title"]);
        Assert.Equal("c\\d", reparsedParagraph.Attributes.Attributes["data-path"]);
    }

    [Fact]
    public void Native_ParagraphText_SourceField_Excludes_Trailing_GenericAttributes() {
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.GenericAttributes = true;
        options.PreserveTrivia = true;
        var native = MarkdownNativeDocument.Parse("Alpha {#id}\n", options);

        var field = Assert.Single(native.EnumerateBlockSourceFields("paragraphText"));

        Assert.Equal("Alpha", field.Value);
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 5), field.SourceSpan);
        Assert.Equal("Beta {#id}\n", native.CreateReplaceEdit(field, "Beta").Apply(native.SourceMarkdown));
    }

    [Fact]
    public void Toc_Uses_Explicit_Heading_Id_And_Reserves_It_For_Generated_Anchors() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.GenericAttributes = true;
        var document = MarkdownReader.Parse("""
            [TOC]

            # Install {#setup}

            # Setup
            """, options);

        var html = document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        });

        Assert.Contains("href=\"#setup\"", html);
        Assert.Contains("id=\"setup\"", html);
        Assert.Contains("id=\"setup-1\"", html);
        Assert.DoesNotContain("href=\"#install\"", html);
    }

    [Fact]
    public void CodeBlock_Html_Renders_Bare_Fence_Id_And_Classes_Without_Opaque_Options() {
        var document = MarkdownReader.Parse("""
            ```cs linenums #code .wide
            Console.WriteLine();
            ```
            """, MarkdownReaderOptions.CreatePortableProfile());

        var html = document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        });

        Assert.Contains("id=\"code\"", html);
        Assert.Contains("class=\"wide language-cs\"", html);
        Assert.DoesNotContain("linenums", html);
    }

    [Fact]
    public void Standalone_GenericAttributes_Before_Type7_HtmlBlock_Are_Consumed_Without_Metadata() {
        const string markdown = "{#html .wide}\n<custom>\nok\n</custom>\n\n";
        var options = new MarkdownReaderOptions {
            GenericAttributes = true,
            PreserveTrivia = true,
            HtmlBlocks = true
        };

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(result);

        var htmlBlock = Assert.IsType<HtmlRawBlock>(Assert.Single(result.Document.Blocks));
        Assert.True(htmlBlock.Attributes.IsEmpty);
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

        Assert.Equal("<custom>\nok\n</custom>", html);
    }

    private sealed class RendererInspectTransform(Func<MarkdownDoc, MarkdownDocumentTransformContext, MarkdownDoc> inspect) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            Assert.Equal(MarkdownDocumentTransformSource.MarkdownRenderer, context.Source);
            return inspect(document, context);
        }
    }
}
