using System;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.MarkdownRenderer;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;
using OfficeIMO.Word.Markdown;
using Xunit;
using PdfTextRun = OfficeIMO.Pdf.TextRun;

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
    public void MarkdownToRtf_Applies_Inserted_Sequence_Style() {
        var options = new MarkdownToRtfOptions {
            ReaderOptions = MarkdownReaderOptions.CreateOfficeIMOProfile()
        };

        var document = "Before ++new++ after".ToRtfDocumentFromMarkdown(options);
        var paragraph = Assert.Single(document.Paragraphs);

        Assert.Contains(paragraph.Runs, run => run.Text == "new" && run.UnderlineStyle != RtfUnderlineStyle.None);
    }

    [Fact]
    public void MarkdownDoc_ToRtfDocument_Applies_Scalar_Inserted_Superscript_And_Subscript_Styles() {
        var markdown = MarkdownDoc.Create();
        markdown.Add(new ParagraphBlock(new InlineSequence()
            .Text("Before ")
            .Inserted("new")
            .Text(" H")
            .Superscript("2")
            .Subscript("n")));

        var document = markdown.ToRtfDocument();
        var paragraph = Assert.Single(document.Paragraphs);

        Assert.Contains(paragraph.Runs, run => run.Text == "new" && run.UnderlineStyle != RtfUnderlineStyle.None);
        Assert.Contains(paragraph.Runs, run => run.Text == "2" && run.VerticalPosition == RtfVerticalPosition.Superscript);
        Assert.Contains(paragraph.Runs, run => run.Text == "n" && run.VerticalPosition == RtfVerticalPosition.Subscript);
    }

    [Fact]
    public void LoadFromMarkdown_Honors_Custom_ReaderOptions_For_Inserted_Superscript_And_Subscript() {
        var readerOptions = MarkdownReaderOptions.CreatePortableProfile();
        readerOptions.Inserted = true;
        readerOptions.Superscript = true;
        readerOptions.Subscript = true;

        using var document = "Before ++new++ and ^up^ plus H~down~O".LoadFromMarkdown(new MarkdownToWordOptions {
            ReaderOptions = readerOptions
        });

        Assert.Contains(document.Paragraphs, run => run.Text == "new" && run.Underline == UnderlineValues.Single);
        Assert.Contains(document.Paragraphs, run => run.Text == "up" && run.VerticalTextAlignment == VerticalPositionValues.Superscript);
        Assert.Contains(document.Paragraphs, run => run.Text == "down" && run.VerticalTextAlignment == VerticalPositionValues.Subscript);
    }

    [Fact]
    public void MarkdownDoc_ToWordDocument_Applies_Scalar_Inserted_Superscript_And_Subscript_Styles() {
        var markdown = MarkdownDoc.Create();
        markdown.Add(new ParagraphBlock(new InlineSequence()
            .Text("Before ")
            .Inserted("new")
            .Text(" H")
            .Superscript("2")
            .Subscript("n")));

        using var document = markdown.ToWordDocument();

        Assert.Contains(document.Paragraphs, run => run.Text == "new" && run.Underline == UnderlineValues.Single);
        Assert.Contains(document.Paragraphs, run => run.Text == "2" && run.VerticalTextAlignment == VerticalPositionValues.Superscript);
        Assert.Contains(document.Paragraphs, run => run.Text == "n" && run.VerticalTextAlignment == VerticalPositionValues.Subscript);
    }

    [Fact]
    public void LoadFromMarkdown_Preserves_Abbreviation_Decoded_Entity_And_SoftBreak_Text_In_Word() {
        var readerOptions = MarkdownReaderOptions.CreateOfficeIMOProfile();
        readerOptions.Abbreviations = true;

        using var parsed = "*[HTML]: Hyper Text Markup Language\n\nHTML and A &amp; B".LoadFromMarkdown(new MarkdownToWordOptions {
            ReaderOptions = readerOptions
        });

        var parsedText = string.Concat(parsed.Paragraphs.Select(run => run.Text));
        Assert.Contains("HTML", parsedText);
        Assert.Contains("A & B", parsedText);

        var typed = MarkdownDoc.Create();
        typed.Add(new ParagraphBlock(new InlineSequence()
            .Text("Alpha")
            .SoftBreak()
            .Text("Beta")));

        using var typedDocument = typed.ToWordDocument();
        Assert.Contains("Alpha Beta", string.Concat(typedDocument.Paragraphs.Select(run => run.Text)));
    }

    [Fact]
    public void MarkdownPdf_Preserves_Inserted_Superscript_And_Subscript_Run_Styles() {
        var readerOptions = MarkdownReaderOptions.CreateOfficeIMOProfile();
        readerOptions.Inserted = true;
        readerOptions.Superscript = true;
        readerOptions.Subscript = true;

        var pdf = "Before ++new++ ^up^ H~down~O".ToPdfDocument(new MarkdownPdfSaveOptions {
            ReaderOptions = readerOptions
        });

        var runs = GetTopLevelPdfTextRuns(pdf);

        Assert.Contains(runs, run => run.Text == "new" && run.Underline);
        Assert.Contains(runs, run => run.Text == "up" && run.Baseline == OfficeIMO.Pdf.PdfTextBaseline.Superscript);
        Assert.Contains(runs, run => run.Text == "down" && run.Baseline == OfficeIMO.Pdf.PdfTextBaseline.Subscript);
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
    public void MarkdownRenderer_RenderBodyHtml_ReaderTransforms_Receive_Syntax_Context() {
        var sourceSliceCreated = false;
        var sourceText = string.Empty;
        var options = new MarkdownRendererOptions();
        options.ReaderOptions.DocumentTransforms.Add(new ReaderInspectTransform((document, context) => {
            var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));

            Assert.Equal(MarkdownDocumentTransformSource.MarkdownReader, context.Source);
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
    public void MarkdownRenderer_RenderBodyHtml_SyntaxExtensions_Force_Syntax_Context() {
        var options = new MarkdownRendererOptions();
        options.ReaderOptions.PreserveTrivia = true;
        options.HtmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        };
        options.HtmlOptions.SyntaxInlineRenderExtensions.Add(MarkdownSyntaxInlineHtmlRenderExtension.CreateContextual(
            "current-head-code-syntax",
            MarkdownSyntaxKind.InlineCodeSpan,
            (inline, syntaxNode, context) => {
                Assert.True(context.TryCreateSourceSlice(syntaxNode, out var slice));
                return $"<kbd data-source=\"{System.Net.WebUtility.HtmlEncode(slice.Text)}\">syntax</kbd>";
            }));

        var html = OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("Use `code` now.\n", options);

        Assert.Contains("<kbd data-source=\"`code`\">syntax</kbd>", html, StringComparison.Ordinal);
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
    public void GenericAttributes_Attach_To_Inserted_Inlines() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.GenericAttributes = true;
        options.Inserted = true;

        var document = MarkdownReader.Parse("++new++{#added .fresh}", options);
        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        var inserted = Assert.IsType<InsertedSequenceInline>(Assert.Single(paragraph.Inlines.Nodes));

        var html = document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            EscapeNonAsciiText = false
        });

        Assert.Contains("<ins", html);
        Assert.Contains("id=\"added\"", html);
        Assert.Contains("class=\"fresh\"", html);
        Assert.Equal("added", inserted.Attributes.ElementId);
        Assert.Equal("fresh", Assert.Single(inserted.Attributes.Classes));
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

    [Fact]
    public void NoPipe_Table_Body_Terminates_Before_Abbreviation_Definition() {
        var options = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        options.Abbreviations = true;
        options.RequireTableBodyRowPipes = false;

        var document = MarkdownReader.Parse("""
            | Name |
            | ---- |
            HTML
            *[HTML]: Hyper Text Markup Language

            HTML
            """, options);

        var table = Assert.IsType<TableBlock>(document.Blocks[0]);
        Assert.Single(table.Rows);
        Assert.Equal("HTML", Assert.Single(table.Rows[0]));
        Assert.IsType<ParagraphBlock>(document.Blocks[1]);

        var html = document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        });

        Assert.DoesNotContain("*[HTML]", html);
        Assert.Contains("<abbr title=\"Hyper Text Markup Language\">HTML</abbr>", html);
    }

    [Fact]
    public void Inline_Autolinks_Win_Before_Abbreviation_Prefixes() {
        var options = MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile();
        options.Abbreviations = true;
        options.AutolinkBareSchemeUrls = true;
        options.AutolinkEmails = true;

        var document = MarkdownReader.Parse("""
            *[https]: protocol
            *[user]: account

            https://example.com user@example.com
            """, options);

        var html = document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        });

        Assert.Contains("<a href=\"https://example.com\">https://example.com</a>", html);
        Assert.Contains("<a href=\"mailto:user@example.com\">user@example.com</a>", html);
        Assert.DoesNotContain("<abbr title=\"protocol\">https</abbr>://example.com", html);
        Assert.DoesNotContain("<abbr title=\"account\">user</abbr>@example.com", html);
    }

    [Fact]
    public void Abbreviation_PreScan_Honors_Disabled_FencedCode() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.Abbreviations = true;
        options.FencedCode = false;

        var document = MarkdownReader.Parse("""
            HTML

            ```
            *[HTML]: Hyper Text Markup Language
            ```
            """, options);

        var html = document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        });

        Assert.Contains("<abbr title=\"Hyper Text Markup Language\">HTML</abbr>", html);
        Assert.DoesNotContain("*[HTML]", html);
    }

    [Fact]
    public void Native_RawHtml_Opening_SourceField_Uses_Remapped_Nested_SourceSpan() {
        const string markdown = "> <script>\n> alert(1)\n> </script>\n";
        var native = MarkdownNativeDocument.Parse(markdown, new MarkdownReaderOptions {
            PreserveTrivia = true,
            HtmlBlocks = true
        });

        var field = Assert.Single(native.EnumerateBlockSourceFields("htmlOpeningTag"));

        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 10), field.SourceSpan);
        Assert.True(native.TryCreateOriginalSourceSlice(field, out var slice));
        Assert.Equal("<script>", slice.Text);
        Assert.Equal("> <section>\n> alert(1)\n> </script>\n", native.CreateReplaceEdit(field, "<section>").Apply(markdown));
    }

    [Fact]
    public void Native_Details_Tag_SourceFields_Use_Remapped_Nested_SourceSpans() {
        const string markdown = "> <details open>\n> <summary>More</summary>\n> body\n> </details>\n";
        var native = MarkdownNativeDocument.Parse(markdown, new MarkdownReaderOptions {
            PreserveTrivia = true,
            HtmlBlocks = true
        });

        var opening = Assert.Single(native.EnumerateBlockSourceFields("detailsOpeningTag"));
        var closing = Assert.Single(native.EnumerateBlockSourceFields("detailsClosingTag"));

        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 16), opening.SourceSpan);
        Assert.Equal(new MarkdownSourceSpan(4, 3, 4, 12), closing.SourceSpan);
        Assert.True(native.TryCreateOriginalSourceSlice(opening, out var openingSlice));
        Assert.True(native.TryCreateOriginalSourceSlice(closing, out var closingSlice));
        Assert.Equal("<details open>", openingSlice.Text);
        Assert.Equal("</details>", closingSlice.Text);
    }

    [Fact]
    public void Nested_Standalone_GenericAttributes_Attach_To_Following_Paragraph() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.GenericAttributes = true;

        var document = MarkdownReader.Parse("""
            - a
              {#para .lead}
              b
            """, options);

        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var item = Assert.Single(list.Items);
        var paragraph = Assert.Single(item.Children.OfType<ParagraphBlock>(), block => block.Attributes.ElementId == "para");

        Assert.Equal("para", paragraph.Attributes.ElementId);
        Assert.Equal("lead", Assert.Single(paragraph.Attributes.Classes));

        var html = document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            EscapeNonAsciiText = false
        });

        Assert.DoesNotContain("{#para", html);
        Assert.Contains("id=\"para\"", html);
        Assert.Contains("class=\"lead\"", html);
    }

    [Fact]
    public void DefinitionLists_Allow_Heading_Looking_Terms_When_Headings_Disabled() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DefinitionLists = true;
        options.Headings = false;

        var document = MarkdownReader.Parse("""
            # term
            :   definition
            """, options);

        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(document.Blocks));
        var html = ((IMarkdownBlock)definitionList).RenderHtml();

        Assert.Contains("<dt># term</dt>", html, StringComparison.Ordinal);
        Assert.Contains("<dd>definition</dd>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void DefinitionLists_Keep_Fence_Looking_Lazy_Continuation_When_FencedCode_Disabled() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DefinitionLists = true;
        options.FencedCode = false;

        var document = MarkdownReader.Parse("""
            Term
            :   ```
                code
                ```
            after
            """, options);

        var definitionList = Assert.IsType<DefinitionListBlock>(Assert.Single(document.Blocks));
        var html = ((IMarkdownBlock)definitionList).RenderHtml();

        Assert.Contains("after", html, StringComparison.Ordinal);
        Assert.Single(document.Blocks);
    }

    [Fact]
    public void Native_Footnote_Token_SourceFields_Use_Remapped_Nested_SourceSpans() {
        const string markdown = "> [^n]: note\n";
        var native = MarkdownNativeDocument.Parse(markdown, new MarkdownReaderOptions {
            PreserveTrivia = true,
            Footnotes = true
        });

        var quote = Assert.IsType<MarkdownNativeQuoteBlock>(Assert.Single(native.Blocks));
        var footnote = Assert.IsType<MarkdownNativeFootnoteDefinitionBlock>(Assert.Single(quote.Children));

        Assert.Equal(new MarkdownSourceSpan(1, 3, 1, 4), footnote.OpeningMarkerSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 5, 1, 5), footnote.LabelSourceSpan);
        Assert.Equal(new MarkdownSourceSpan(1, 6, 1, 7), footnote.SeparatorMarkerSourceSpan);
        Assert.True(native.TryCreateOriginalSourceSlice(footnote.OpeningMarkerSourceSpan!.Value, out var openingSlice));
        Assert.Equal("[^", openingSlice.Text);
    }

    [Fact]
    public void CustomContainer_RenderMarkdown_Preserves_GenericAttributes_For_Reparse() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.CustomContainers = true;
        options.GenericAttributes = true;

        var document = MarkdownReader.Parse("""
            {#box .wide}
            ::: note
            hello
            :::
            """, options);

        var container = Assert.IsType<CustomContainerBlock>(Assert.Single(document.Blocks));
        var rendered = ((IMarkdownBlock)container).RenderMarkdown();
        var reparsed = MarkdownReader.Parse(rendered, options);
        var reparsedContainer = Assert.IsType<CustomContainerBlock>(Assert.Single(reparsed.Blocks));

        Assert.StartsWith("{#box .wide}", rendered, StringComparison.Ordinal);
        Assert.Equal("box", reparsedContainer.Attributes.ElementId);
        Assert.Equal("wide", Assert.Single(reparsedContainer.Attributes.Classes));
    }

    [Fact]
    public void Nested_Standalone_GenericAttributes_Attach_To_Following_CustomContainer() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.CustomContainers = true;
        options.GenericAttributes = true;

        var document = MarkdownReader.Parse("""
            - item
              {#box .wide}
              ::: note
              body
              :::
            """, options);

        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var item = Assert.Single(list.Items);
        var container = Assert.Single(item.Children.OfType<CustomContainerBlock>());

        Assert.Equal("box", container.Attributes.ElementId);
        Assert.Equal("wide", Assert.Single(container.Attributes.Classes));
        Assert.Contains("body", ((IMarkdownBlock)container).RenderHtml(), StringComparison.Ordinal);
    }

    [Fact]
    public void GenericAttributes_Allow_One_Character_Ids() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.GenericAttributes = true;

        var document = MarkdownReader.Parse("""
            # Title {#x}
            Text {#p}
            """, options);

        var heading = Assert.IsType<HeadingBlock>(document.Blocks[0]);
        var paragraph = Assert.IsType<ParagraphBlock>(document.Blocks[1]);
        var headingHtml = ((IMarkdownBlock)heading).RenderHtml();
        var paragraphHtml = ((IMarkdownBlock)paragraph).RenderHtml();

        Assert.Equal("x", heading.Attributes.ElementId);
        Assert.Equal("p", paragraph.Attributes.ElementId);
        Assert.Contains("id=\"x\"", headingHtml, StringComparison.Ordinal);
        Assert.Contains("id=\"p\"", paragraphHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void DefinitionLists_Allow_List_Looking_Terms_When_List_Parser_Disabled() {
        var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
        options.DefinitionLists = true;
        options.UnorderedLists = false;
        options.OrderedLists = false;

        var unordered = Assert.Single(MarkdownReader.Parse("""
            - term
            :   unordered definition
            """, options).Blocks.OfType<DefinitionListBlock>());
        var ordered = Assert.Single(MarkdownReader.Parse("""
            1. term
            :   ordered definition
            """, options).Blocks.OfType<DefinitionListBlock>());

        Assert.Contains("<dt>- term</dt>", ((IMarkdownBlock)unordered).RenderHtml(), StringComparison.Ordinal);
        Assert.Contains("<dd>unordered definition</dd>", ((IMarkdownBlock)unordered).RenderHtml(), StringComparison.Ordinal);
        Assert.Contains("<dt>1. term</dt>", ((IMarkdownBlock)ordered).RenderHtml(), StringComparison.Ordinal);
        Assert.Contains("<dd>ordered definition</dd>", ((IMarkdownBlock)ordered).RenderHtml(), StringComparison.Ordinal);
    }

    [Fact]
    public void Native_TablePipe_SourceFields_Use_Visual_Columns_After_Tabs() {
        const string markdown = "| \t| B |\n| --- | --- |\n| C | D |\n";

        var native = MarkdownNativeDocument.Parse(markdown, new MarkdownReaderOptions {
            Tables = true,
            PreserveTrivia = true
        });

        var table = Assert.IsType<MarkdownNativeTableBlock>(Assert.Single(native.Blocks));
        var firstLinePipes = table.EnumerateSourceFields("tablePipe")
            .Where(pipe => pipe.SourceSpan.StartLine == 1)
            .ToArray();

        Assert.Equal(new[] { 1, 5, 9 }, firstLinePipes.Select(pipe => pipe.SourceSpan.StartColumn!.Value).ToArray());
        Assert.True(native.TryCreateOriginalSourceSlice(firstLinePipes[1].SourceSpan, out var pipeSlice));
        Assert.Equal("|", pipeSlice.Text);
    }

    private sealed class RendererInspectTransform(Func<MarkdownDoc, MarkdownDocumentTransformContext, MarkdownDoc> inspect) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            Assert.Equal(MarkdownDocumentTransformSource.MarkdownRenderer, context.Source);
            return inspect(document, context);
        }
    }

    private sealed class ReaderInspectTransform(Func<MarkdownDoc, MarkdownDocumentTransformContext, MarkdownDoc> inspect) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) => inspect(document, context);
    }

    private static PdfTextRun[] GetTopLevelPdfTextRuns(OfficeIMO.Pdf.PdfDocument document) {
        var blocksProperty = typeof(OfficeIMO.Pdf.PdfDocument).GetProperty("Blocks", BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(blocksProperty);

        var blocks = ((System.Collections.IEnumerable)blocksProperty.GetValue(document)!).Cast<object>().ToArray();
        var block = Assert.Single(blocks);
        var runsProperty = block.GetType().GetProperty("Runs", BindingFlags.Instance | BindingFlags.Public);
        Assert.NotNull(runsProperty);

        return ((System.Collections.IEnumerable)runsProperty.GetValue(block)!).Cast<PdfTextRun>().ToArray();
    }
}
