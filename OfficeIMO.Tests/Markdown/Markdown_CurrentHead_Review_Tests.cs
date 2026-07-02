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

    private sealed class RendererInspectTransform(Func<MarkdownDoc, MarkdownDocumentTransformContext, MarkdownDoc> inspect) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            Assert.Equal(MarkdownDocumentTransformSource.MarkdownRenderer, context.Source);
            return inspect(document, context);
        }
    }
}
