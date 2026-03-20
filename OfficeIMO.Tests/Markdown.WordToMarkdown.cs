using System;
using System.IO;
using System.Linq;
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void WordToMarkdown_ConvertsElements() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Heading").Style = WordParagraphStyles.Heading1;

            var paragraph = doc.AddParagraph("This is ");
            paragraph.AddText("bold").Bold = true;
            paragraph.AddText(" and ");
            paragraph.AddText("italic").Italic = true;
            paragraph.AddText(" with ");
            paragraph.AddText("strike").Strike = true;
            paragraph.AddText(" and ");
            paragraph.AddText("code").SetFontFamily(FontResolver.Resolve("monospace")!);

            var list = doc.AddList(WordListStyle.Bulleted);
            list.AddItem("Item 1");
            list.AddItem("Item 2");

            var linkParagraph = doc.AddParagraph("Visit ");
            linkParagraph.AddHyperLink("OfficeIMO", new Uri("https://example.com"));

            var table = doc.AddTable(2, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "H1";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "H2";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "C1";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "C2";

            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            doc.AddParagraph().AddImage(imagePath);

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions());

            Assert.Contains("# Heading", markdown);
            Assert.Contains("**bold**", markdown);
            Assert.Contains("*italic*", markdown);
            Assert.Contains("~~strike~~", markdown);
            Assert.Contains("`code`", markdown);
            Assert.Contains("- Item 1", markdown);
            Assert.Matches("\\[OfficeIMO\\]\\(https://example\\.com/?\\)", markdown);
            Assert.Contains("| H1 | H2 |", markdown);
            Assert.Contains("data:image/png;base64", markdown);
        }

        [Fact]
        public void WordToMarkdown_HandlesFootNotes() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Hello").AddFootNote("First note");
            doc.AddParagraph("World").AddFootNote("Second note");

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions());

            Assert.Contains("Hello[^1]", markdown);
            Assert.Contains("World[^2]", markdown);
            Assert.Contains("[^1]: First note", markdown);
            Assert.Contains("[^2]: Second note", markdown);
        }

        [Fact]
        public void WordToMarkdown_ExportsImagesToFiles() {
            using var doc = WordDocument.Create();
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            doc.AddParagraph().AddImage(imagePath);

            string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);

            var options = new WordToMarkdownOptions {
                ImageExportMode = ImageExportMode.File,
                ImageDirectory = tempDir
            };

            string markdown = doc.ToMarkdown(options);

            string fileName = Path.GetFileName(imagePath);
            Assert.Contains($"![", markdown);
            Assert.Contains(fileName, markdown);
            Assert.True(File.Exists(Path.Combine(tempDir, fileName)));
        }

        [Fact]
        public void WordToMarkdown_ToMarkdownDocument_BuildsTypedAstDirectly() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph("Water");
            paragraph.AddText("2").SetVerticalTextAlignment(DocumentFormat.OpenXml.Wordprocessing.VerticalPositionValues.Subscript);
            paragraph.AddText(" and ");
            paragraph.AddText("important").Underline = DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single;

            MarkdownDoc markdown = doc.ToMarkdownDocument(new WordToMarkdownOptions { EnableUnderline = true });
            var block = Assert.IsType<ParagraphBlock>(Assert.Single(markdown.Blocks));

            Assert.Contains(block.Inlines.Nodes, inline => inline is HtmlTagSequenceInline tag && tag.TagName == "sub");
            Assert.Contains(block.Inlines.Nodes, inline => inline is HtmlTagSequenceInline tag && tag.TagName == "u");

            string renderedHtml = markdown.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<sub>2</sub>", renderedHtml, StringComparison.Ordinal);
            Assert.Contains("<u>important</u>", renderedHtml, StringComparison.Ordinal);
        }

        [Fact]
        public void WordToMarkdown_ToMarkdownDocument_Preserves_NonParagraph_Footnote_Blocks() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Body").AddFootNote("Console.WriteLine(1);");
            var footnoteParagraph = doc.FootNotes[0].Paragraphs![1];
            footnoteParagraph.SetStyleId("CodeLang_csharp");

            MarkdownDoc markdown = doc.ToMarkdownDocument(new WordToMarkdownOptions());
            var footnote = Assert.IsType<FootnoteDefinitionBlock>(Assert.Single(markdown.Blocks, block => block is FootnoteDefinitionBlock));
            var code = Assert.IsType<CodeBlock>(Assert.Single(footnote.Blocks));

            Assert.Equal("csharp", code.Language);
            Assert.Equal("Console.WriteLine(1);", code.Content);
            Assert.Empty(footnote.ParagraphBlocks);

            string renderedHtml = markdown.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
            Assert.Contains("<li id=\"fn:1\"><pre><code class=\"language-csharp\">Console.WriteLine(1);", renderedHtml, StringComparison.Ordinal);
            Assert.Contains("<a class=\"footnote-backref\"", renderedHtml, StringComparison.Ordinal);
        }

        [Fact]
        public void WordToMarkdown_ToMarkdownDocument_Preserves_TableCell_Element_Order_And_Nested_Tables() {
            using var doc = WordDocument.Create();
            var outerTable = doc.AddTable(2, 1);
            outerTable.Rows[0].Cells[0].Paragraphs[0].Text = "Container";

            var bodyCell = outerTable.Rows[1].Cells[0];
            bodyCell.Paragraphs[0].Text = "Before nested";

            var nestedTable = bodyCell.AddTable(2, 1);
            nestedTable.Rows[0].Cells[0].Paragraphs[0].Text = "Inner Header";
            nestedTable.Rows[1].Cells[0].Paragraphs[0].Text = "Inner Value";
            bodyCell.Paragraphs[bodyCell.Paragraphs.Count - 1].Text = "After nested";

            MarkdownDoc markdown = doc.ToMarkdownDocument(new WordToMarkdownOptions());
            var table = Assert.IsType<TableBlock>(Assert.Single(markdown.Blocks));
            var markdownCell = Assert.Single(Assert.Single(table.RowCells));

            Assert.Collection(
                markdownCell.Blocks,
                block => Assert.Equal("Before nested", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => {
                    var innerTable = Assert.IsType<TableBlock>(block);
                    Assert.Equal("Inner Header", Assert.Single(innerTable.HeaderCells).Markdown);
                    Assert.Equal("Inner Value", Assert.Single(Assert.Single(innerTable.RowCells)).Markdown);
                },
                block => Assert.Equal("After nested", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));
        }

        [Fact]
        public void WordToMarkdown_ToMarkdownDocument_Expands_TextBox_Content_In_Document_Order() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Intro");

            var textBox = doc.AddTextBox("Alpha");
            textBox.Paragraphs[0].AddParagraph("Beta");

            doc.AddParagraph("Tail");

            MarkdownDoc markdown = doc.ToMarkdownDocument(new WordToMarkdownOptions());

            Assert.Collection(
                markdown.Blocks,
                block => Assert.Equal("Intro", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => Assert.Equal("Alpha", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => Assert.Equal("Beta", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => Assert.Equal("Tail", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));

            string renderedMarkdown = doc.ToMarkdown(new WordToMarkdownOptions());
            Assert.Contains("Intro", renderedMarkdown, StringComparison.Ordinal);
            Assert.Contains("Alpha", renderedMarkdown, StringComparison.Ordinal);
            Assert.Contains("Beta", renderedMarkdown, StringComparison.Ordinal);
            Assert.Contains("Tail", renderedMarkdown, StringComparison.Ordinal);
        }

        [Fact]
        public void WordToMarkdown_ToMarkdownDocument_Exports_Headers_And_Footers_As_Semantic_Blocks_When_Enabled() {
            using var doc = WordDocument.Create();
            doc.AddHeadersAndFooters();
            doc.Header!.Default!.AddParagraph("Header line");
            doc.AddParagraph("Body line");
            doc.Footer!.Default!.AddParagraph("Footer line");

            var options = new WordToMarkdownOptions {
                IncludeHeadersAndFootersAsSemanticBlocks = true
            };

            MarkdownDoc markdown = doc.ToMarkdownDocument(options);

            Assert.Collection(
                markdown.Blocks,
                block => {
                    var semantic = Assert.IsType<SemanticFencedBlock>(block);
                    Assert.Equal(WordMarkdownSemanticBlocks.HeaderSemanticKind, semantic.SemanticKind);
                    Assert.Equal(WordMarkdownSemanticBlocks.HeaderFenceLanguage, semantic.Language);
                    Assert.Equal("Header line", semantic.Content);
                    Assert.True(semantic.FenceInfo.TryGetInt32Attribute("section", out var section));
                    Assert.Equal(1, section);
                    Assert.Equal("default", semantic.FenceInfo.GetAttribute("slot"));
                },
                block => Assert.Equal("Body line", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => {
                    var semantic = Assert.IsType<SemanticFencedBlock>(block);
                    Assert.Equal(WordMarkdownSemanticBlocks.FooterSemanticKind, semantic.SemanticKind);
                    Assert.Equal(WordMarkdownSemanticBlocks.FooterFenceLanguage, semantic.Language);
                    Assert.Equal("Footer line", semantic.Content);
                    Assert.Equal("default", semantic.FenceInfo.GetAttribute("slot"));
                });

            string renderedMarkdown = doc.ToMarkdown(options);
            Assert.Contains("```officeimo-word-header section=1 slot=default", renderedMarkdown, StringComparison.Ordinal);
            Assert.Contains("Header line", renderedMarkdown, StringComparison.Ordinal);
            Assert.Contains("```officeimo-word-footer section=1 slot=default", renderedMarkdown, StringComparison.Ordinal);
            Assert.Contains("Footer line", renderedMarkdown, StringComparison.Ordinal);
        }

        [Fact]
        public void WordToMarkdown_CreateReaderOptions_Parses_Header_And_Footer_Semantic_Blocks() {
            using var doc = WordDocument.Create();
            doc.AddHeadersAndFooters();
            doc.Header!.Default!.AddParagraph("Header line");
            doc.AddParagraph("Body line");
            doc.Footer!.Default!.AddParagraph("Footer line");

            var options = new WordToMarkdownOptions {
                IncludeHeadersAndFootersAsSemanticBlocks = true
            };

            string renderedMarkdown = doc.ToMarkdown(options);
            MarkdownDoc parsed = MarkdownReader.Parse(renderedMarkdown, options.CreateReaderOptions());

            Assert.Collection(
                parsed.Blocks,
                block => {
                    var semantic = Assert.IsType<SemanticFencedBlock>(block);
                    Assert.Equal(WordMarkdownSemanticBlocks.HeaderSemanticKind, semantic.SemanticKind);
                    Assert.Equal("Header line", semantic.Content);
                    Assert.Equal("default", semantic.FenceInfo.GetAttribute("slot"));
                },
                block => Assert.Equal("Body line", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => {
                    var semantic = Assert.IsType<SemanticFencedBlock>(block);
                    Assert.Equal(WordMarkdownSemanticBlocks.FooterSemanticKind, semantic.SemanticKind);
                    Assert.Equal("Footer line", semantic.Content);
                    Assert.Equal("default", semantic.FenceInfo.GetAttribute("slot"));
                });
        }

        [Fact]
        public void WordMarkdownSemanticBlocks_CreateReaderOptions_Preserves_Loose_List_Item_Paragraphs() {
            const string markdown = """
                - first paragraph

                  second paragraph
                - next item
                """;

            MarkdownDoc parsed = MarkdownReader.Parse(markdown, WordMarkdownSemanticBlocks.CreateReaderOptions());
            var list = Assert.IsType<UnorderedListBlock>(Assert.Single(parsed.Blocks));

            Assert.Collection(
                list.Items,
                item => Assert.Equal(new[] { "first paragraph", "second paragraph" }, item.BlockChildren.Select(block => block.RenderMarkdown()).ToArray()),
                item => Assert.Equal(new[] { "next item" }, item.BlockChildren.Select(block => block.RenderMarkdown()).ToArray()));
        }

        [Fact]
        public void MarkdownToWord_LoadFromMarkdown_Restores_Header_And_Footer_Semantic_Blocks() {
            using var source = WordDocument.Create();
            source.AddHeadersAndFooters();
            source.Header!.Default!.AddParagraph("Header line");
            source.AddParagraph("Body line");
            source.Footer!.Default!.AddParagraph("Footer line");

            string markdown = source.ToMarkdown(new WordToMarkdownOptions {
                IncludeHeadersAndFootersAsSemanticBlocks = true
            });

            using var restored = markdown.LoadFromMarkdown();

            Assert.Equal("Header line", restored.Header!.Default!.Paragraphs.Single(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text)).Text.Trim());
            Assert.Equal("Footer line", restored.Footer!.Default!.Paragraphs.Single(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text)).Text.Trim());
            Assert.Contains(restored.Paragraphs, paragraph => paragraph.Text.Trim() == "Body line");
        }

        [Fact]
        public void MarkdownToWord_ToWordDocument_Routes_First_And_Even_HeaderFooter_Semantic_Blocks() {
            var markdown = MarkdownDoc.Create()
                .Add(new SemanticFencedBlock(
                    WordMarkdownSemanticBlocks.HeaderSemanticKind,
                    $"{WordMarkdownSemanticBlocks.HeaderFenceLanguage} section=1 slot=first",
                    "First header"))
                .Add(new SemanticFencedBlock(
                    WordMarkdownSemanticBlocks.HeaderSemanticKind,
                    $"{WordMarkdownSemanticBlocks.HeaderFenceLanguage} section=1 slot=default",
                    "Default header"))
                .Add(new ParagraphBlock(new InlineSequence().Text("Body line")))
                .Add(new SemanticFencedBlock(
                    WordMarkdownSemanticBlocks.FooterSemanticKind,
                    $"{WordMarkdownSemanticBlocks.FooterFenceLanguage} section=1 slot=even",
                    "Even footer"));

            using var document = markdown.ToWordDocument();

            Assert.True(document.DifferentFirstPage);
            Assert.True(document.DifferentOddAndEvenPages);
            Assert.Equal("First header", document.Header!.First!.Paragraphs.Single(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text)).Text.Trim());
            Assert.Equal("Default header", document.Header!.Default!.Paragraphs.Single(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text)).Text.Trim());
            Assert.Equal("Even footer", document.Footer!.Even!.Paragraphs.Single(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text)).Text.Trim());
            Assert.Contains(document.Paragraphs, paragraph => paragraph.Text.Trim() == "Body line");
        }
    }
}

