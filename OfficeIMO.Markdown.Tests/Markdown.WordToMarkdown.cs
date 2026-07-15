using System;
using System.IO;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;
using OfficeIMO.Markdown;
using OfficeIMO.Drawing;
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
            var code = Assert.IsType<CodeBlock>(Assert.Single(footnote.ChildBlocks));

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
                markdownCell.ChildBlocks,
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
        public void WordToMarkdown_Exports_Native_TableOfContents_As_Markdown_Toc_Marker() {
            using var doc = WordDocument.Create();
            WordTableOfContent toc = doc.AddTableOfContent(minLevel: 2, maxLevel: 5);
            toc.Text = "Contents";
            doc.AddParagraph("Report").Style = WordParagraphStyles.Heading1;
            doc.AddParagraph("Region").Style = WordParagraphStyles.Heading2;
            doc.AddParagraph("Pipeline").Style = WordParagraphStyles.Heading3;

            MarkdownDoc markdownDocument = doc.ToMarkdownDocument(new WordToMarkdownOptions());
            var marker = Assert.IsType<TocMarkerBlock>(markdownDocument.Blocks[0]);
            Assert.Equal(2, marker.MinLevel);
            Assert.Equal(5, marker.MaxLevel);
            Assert.Equal("Contents", marker.Title);

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions());
            Assert.Contains("[TOC min=2 max=5 title=\"Contents\" titleLevel=2]", markdown, StringComparison.Ordinal);
            Assert.DoesNotContain("- [Region]", markdown, StringComparison.Ordinal);

            using var restored = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToWordDocument(new MarkdownToWordOptions());
            Assert.NotNull(restored.TableOfContent);
            Assert.Equal(2, restored.TableOfContent!.MinLevel);
            Assert.Equal(5, restored.TableOfContent.MaxLevel);
        }

        [Fact]
        public void WordToMarkdown_Exports_Titleless_Native_TableOfContents_Without_Default_Title() {
            using var doc = WordDocument.Create();
            WordTableOfContent toc = doc.AddTableOfContent(minLevel: 2, maxLevel: 9);
            toc.Text = string.Empty;
            doc.AddParagraph("Deep heading").Style = WordParagraphStyles.Heading9;

            MarkdownDoc markdownDocument = doc.ToMarkdownDocument(new WordToMarkdownOptions());
            var marker = Assert.IsType<TocMarkerBlock>(markdownDocument.Blocks[0]);
            Assert.False(marker.IncludeTitle);
            Assert.Equal(2, marker.MinLevel);
            Assert.Equal(9, marker.MaxLevel);

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions());
            Assert.Contains("[TOC min=2 max=9]", markdown, StringComparison.Ordinal);
            Assert.DoesNotContain("Table of Contents", markdown, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("title=", markdown, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordToMarkdown_Keeps_PageBreakBefore_Outside_List_Item() {
            using var doc = WordDocument.Create();
            WordList list = doc.AddList(WordListStyle.Bulleted);
            WordParagraph item = list.AddItem("Starts new page");
            item.PageBreakBefore = true;

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                PageBreakMode = MarkdownPageBreakMode.SemanticBlock
            });

            int pageBreakIndex = markdown.IndexOf("```officeimo-word-page-break", StringComparison.Ordinal);
            int listItemIndex = markdown.IndexOf("- Starts new page", StringComparison.Ordinal);

            Assert.True(pageBreakIndex >= 0, markdown);
            Assert.True(listItemIndex > pageBreakIndex, markdown);
            Assert.DoesNotContain("- \r\n  ```officeimo-word-page-break", markdown, StringComparison.Ordinal);
            Assert.DoesNotContain("- \n  ```officeimo-word-page-break", markdown, StringComparison.Ordinal);
        }

        [Fact]
        public void WordToMarkdown_Preserves_Ordered_List_Number_After_Lifted_PageBreak() {
            using var doc = WordDocument.Create();
            WordList list = doc.AddList(WordListStyle.Numbered);
            list.AddItem("One");
            list.AddItem("Two");
            WordParagraph third = list.AddItem("Three");
            third.PageBreakBefore = true;

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                PageBreakMode = MarkdownPageBreakMode.SemanticBlock
            });

            int pageBreakIndex = markdown.IndexOf("```officeimo-word-page-break", StringComparison.Ordinal);
            int thirdItemIndex = markdown.IndexOf("3. Three", StringComparison.Ordinal);

            Assert.True(pageBreakIndex >= 0, markdown);
            Assert.True(thirdItemIndex > pageBreakIndex, markdown);
            Assert.DoesNotContain("1. Three", markdown, StringComparison.Ordinal);
        }

        [Fact]
        public void WordToMarkdown_Projects_Equation_To_Math_Fence_In_List_Item() {
            const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></m:oMathPara>";
            using var doc = WordDocument.Create();
            WordList list = doc.AddList(WordListStyle.Bulleted);
            WordParagraph item = list.AddItem("Formula:");
            item.AddEquation(omml);

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                UnsupportedContentMode = MarkdownUnsupportedContentMode.Placeholder
            });

            Assert.Contains("- Formula:", markdown, StringComparison.Ordinal);
            Assert.Contains("```math", markdown, StringComparison.Ordinal);
            Assert.Contains("x=1", markdown, StringComparison.Ordinal);
            Assert.DoesNotContain("Unsupported Word content: equation", markdown, StringComparison.Ordinal);
        }

        [Fact]
        public void WordToMarkdown_Projects_Structured_Equation_To_Latex() {
            const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara>";
            using var doc = WordDocument.Create();
            doc.AddEquation(omml);

            string markdown = doc.ToMarkdown();

            Assert.Contains("```math", markdown, StringComparison.Ordinal);
            Assert.Contains("\\frac{a}{b}", markdown, StringComparison.Ordinal);
        }

        [Fact]
        public void WordToMarkdown_ProjectsEveryEquationFromAMixedParagraph() {
            const string first = "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>x=1</m:t></m:r></m:oMath>";
            const string second = "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>y=2</m:t></m:r></m:oMath>";
            using var doc = WordDocument.Create();
            WordParagraph paragraph = doc.AddParagraph("before ");
            paragraph.AddEquation(first);
            paragraph.AddText(" between ");
            paragraph.AddEquation(second);
            paragraph.AddText(" after");

            string markdown = doc.ToMarkdown();

            Assert.Equal(2, markdown.Split(new[] { "```math" }, StringSplitOptions.None).Length - 1);
            Assert.Contains("x=1", markdown, StringComparison.Ordinal);
            Assert.Contains("y=2", markdown, StringComparison.Ordinal);
        }

        [Fact]
        public void WordToMarkdown_ProjectsEquationInsideVisibleRevisionWrapper() {
            using var doc = WordDocument.Create();
            WordParagraph paragraph = doc.AddParagraph("Formula:");
            paragraph._paragraph.Append(new MoveToRun(new M.OfficeMath(new M.Run(new M.Text("tracked")))) {
                Id = "1",
                Author = "Reviewer"
            });

            string markdown = doc.ToMarkdown();

            Assert.Contains("```math", markdown, StringComparison.Ordinal);
            Assert.Contains("tracked", markdown, StringComparison.Ordinal);
        }

        [Fact]
        public void WordToMarkdown_ProjectsEquationInsideInlineContentControl() {
            using var doc = WordDocument.Create();
            WordParagraph paragraph = doc.AddParagraph("Formula:");
            paragraph._paragraph.Append(new SdtRun(
                new SdtProperties(new SdtId { Val = 2076 }),
                new SdtContentRun(
                    new Run(new Text(" control-prefix ")),
                    new M.OfficeMath(new M.Run(new M.Text("controlled"))),
                    new Run(new Text(" control-suffix")))));

            string markdown = doc.ToMarkdown();

            Assert.Contains("Formula:", markdown, StringComparison.Ordinal);
            Assert.Contains("control-prefix", markdown, StringComparison.Ordinal);
            Assert.Contains("control-suffix", markdown, StringComparison.Ordinal);
            Assert.Contains("```math", markdown, StringComparison.Ordinal);
            Assert.Equal(1, markdown.Split(new[] { "controlled" }, StringSplitOptions.None).Length - 1);
        }

        [Fact]
        public void WordToMarkdown_ProjectsEquationAndSurroundingTextInsideHyperlink() {
            using var doc = WordDocument.Create();
            WordParagraph paragraph = doc.AddParagraph("outer ");
            paragraph._paragraph.Append(new Hyperlink(
                new Run(new RunProperties(new Bold()), new Text("link-prefix ")),
                new M.OfficeMath(new M.Run(new M.Text("linked"))),
                new Run(new RunProperties(new Italic()), new Text(" link-suffix"))) {
                Anchor = "target"
            });

            string markdown = doc.ToMarkdown();

            Assert.Contains("link-prefix", markdown, StringComparison.Ordinal);
            Assert.Contains("link-suffix", markdown, StringComparison.Ordinal);
            Assert.Contains("```math", markdown, StringComparison.Ordinal);
            Assert.Contains("**link-prefix**", markdown, StringComparison.Ordinal);
            Assert.Contains("*link-suffix*", markdown, StringComparison.Ordinal);
            Assert.Equal(1, markdown.Split(new[] { "linked" }, StringSplitOptions.None).Length - 1);
            Assert.Equal(1, markdown.Split(new[] { "link-prefix" }, StringSplitOptions.None).Length - 1);
            Assert.Equal(1, markdown.Split(new[] { "link-suffix" }, StringSplitOptions.None).Length - 1);
        }

        [Fact]
        public void WordToMarkdown_ExportsComplexEqFieldOnceAsMathFence() {
            using var doc = WordDocument.Create();
            WordParagraph paragraph = doc.AddParagraph("before ");
            paragraph._paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" EQ \\f(a,b) ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("(a)/(b)")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            paragraph.AddText(" after");

            string markdown = doc.ToMarkdown();

            Assert.Contains("before", markdown, StringComparison.Ordinal);
            Assert.Contains(" after", markdown, StringComparison.Ordinal);
            Assert.Contains("```math", markdown, StringComparison.Ordinal);
            Assert.Equal(1, markdown.Split(new[] { "(a)/(b)" }, StringSplitOptions.None).Length - 1);
        }

        [Fact]
        public void WordToMarkdown_PageBreakPathDoesNotDuplicateComplexEqCachedResult() {
            using var doc = WordDocument.Create();
            WordParagraph paragraph = doc.AddParagraph("before ");
            paragraph._paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" EQ \\f(a,b) ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("(a)/(b)")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            paragraph.AddBreak(BreakValues.Page);
            paragraph.AddText(" after");

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                PageBreakMode = MarkdownPageBreakMode.SemanticBlock
            });

            Assert.Contains("before", markdown, StringComparison.Ordinal);
            Assert.Contains(" after", markdown, StringComparison.Ordinal);
            Assert.Contains("```math", markdown, StringComparison.Ordinal);
            Assert.Contains("```officeimo-word-page-break", markdown, StringComparison.Ordinal);
            Assert.Equal(1, markdown.Split(new[] { "(a)/(b)" }, StringSplitOptions.None).Length - 1);
        }

        [Fact]
        public void WordToMarkdown_PreservesEquationAndShapeFallbackFromMixedParagraph() {
            const string omml = "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>x=1</m:t></m:r></m:oMath>";
            using var doc = WordDocument.Create();
            WordParagraph paragraph = doc.AddParagraph("Mixed:");
            paragraph.AddEquation(omml);
            paragraph.AddShape(ShapeType.Rectangle, 60, 30);

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                UnsupportedContentMode = MarkdownUnsupportedContentMode.Placeholder
            });

            Assert.Contains("```math", markdown, StringComparison.Ordinal);
            Assert.Contains("x=1", markdown, StringComparison.Ordinal);
            Assert.Contains("Unsupported Word content: shape", markdown, StringComparison.Ordinal);
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
            MarkdownDoc parsed = OfficeIMO.Markdown.MarkdownReader.Parse(renderedMarkdown, options.CreateReaderOptions());

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
        public void WordToMarkdown_Preserves_PageBreaks_As_Semantic_Blocks() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph("Before");
            paragraph.AddBreak(BreakValues.Page);
            paragraph.AddText("After");

            MarkdownDoc markdown = doc.ToMarkdownDocument(new WordToMarkdownOptions());

            Assert.Collection(
                markdown.Blocks,
                block => Assert.Equal("Before", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => {
                    var semantic = Assert.IsType<SemanticFencedBlock>(block);
                    Assert.Equal(WordMarkdownSemanticBlocks.PageBreakSemanticKind, semantic.SemanticKind);
                    Assert.Equal(WordMarkdownSemanticBlocks.PageBreakFenceLanguage, semantic.Language);
                },
                block => Assert.Equal("After", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));

            string renderedMarkdown = markdown.ToMarkdown();
            Assert.Contains("Before", renderedMarkdown, StringComparison.Ordinal);
            Assert.Contains("```officeimo-word-page-break", renderedMarkdown, StringComparison.Ordinal);
            Assert.Contains("After", renderedMarkdown, StringComparison.Ordinal);
        }

        [Fact]
        public void WordToMarkdown_Preserves_Text_Around_Run_Level_PageBreaks() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph();
            paragraph._paragraph.Append(new Run(
                new Text("Before"),
                new Break { Type = BreakValues.Page },
                new Text("After")));

            MarkdownDoc markdown = doc.ToMarkdownDocument(new WordToMarkdownOptions());

            Assert.Collection(
                markdown.Blocks,
                block => Assert.Equal("Before", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => Assert.IsType<SemanticFencedBlock>(block),
                block => Assert.Equal("After", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));
        }

        [Fact]
        public void WordToMarkdown_Detects_PageBreak_After_TextWrapping_Break_In_Same_Run() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph();
            paragraph._paragraph.Append(new Run(
                new Text("Before"),
                new Break(),
                new Text("Middle"),
                new Break { Type = BreakValues.Page },
                new Text("After")));

            MarkdownDoc markdown = doc.ToMarkdownDocument(new WordToMarkdownOptions());

            Assert.Collection(
                markdown.Blocks,
                block => {
                    string text = Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown();
                    Assert.Contains("Before", text, StringComparison.Ordinal);
                    Assert.Contains("Middle", text, StringComparison.Ordinal);
                    Assert.DoesNotContain("\u2028", text, StringComparison.Ordinal);
                },
                block => Assert.IsType<SemanticFencedBlock>(block),
                block => Assert.Equal("After", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));
        }

        [Fact]
        public void WordToMarkdown_Preserves_Paragraph_Level_PageBreaks_Before_Content() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Before");
            var after = doc.AddParagraph("After");
            after.PageBreakBefore = true;

            MarkdownDoc markdown = doc.ToMarkdownDocument(new WordToMarkdownOptions());

            Assert.Collection(
                markdown.Blocks,
                block => Assert.Equal("Before", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => Assert.IsType<SemanticFencedBlock>(block),
                block => Assert.Equal("After", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()));
        }

        [Theory]
        [InlineData(MarkdownPageBreakMode.Html, "<div style=\"page-break-after: always;\"></div>")]
        [InlineData(MarkdownPageBreakMode.HorizontalRule, "---")]
        public void WordToMarkdown_PageBreakMode_Controls_Lossy_Output(MarkdownPageBreakMode mode, string expectedMarker) {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Before");
            doc.AddPageBreak();
            doc.AddParagraph("After");

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                PageBreakMode = mode
            });

            Assert.Contains(expectedMarker, markdown, StringComparison.Ordinal);
            Assert.DoesNotContain(WordMarkdownSemanticBlocks.PageBreakFenceLanguage, markdown, StringComparison.Ordinal);
        }

        [Fact]
        public void WordToMarkdown_PageBreakMode_Can_Omit_PageBreaks() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Before");
            doc.AddPageBreak();
            doc.AddParagraph("After");

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                PageBreakMode = MarkdownPageBreakMode.Omit
            });

            Assert.DoesNotContain(WordMarkdownSemanticBlocks.PageBreakFenceLanguage, markdown, StringComparison.Ordinal);
            Assert.DoesNotContain("<div style=\"page-break-after: always;\"></div>", markdown, StringComparison.Ordinal);
        }

        [Fact]
        public void WordToMarkdown_UnsupportedContentMode_Can_Emit_Placeholders() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Before");
            doc.AddShape(ShapeType.Rectangle, 60, 30);
            doc.AddParagraph("After");
            var warnings = new List<string>();

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                UnsupportedContentMode = MarkdownUnsupportedContentMode.Placeholder,
                OnWarning = warnings.Add
            });

            Assert.Contains("Before", markdown, StringComparison.Ordinal);
            Assert.Contains("Unsupported Word content: shape", markdown, StringComparison.Ordinal);
            Assert.Contains("After", markdown, StringComparison.Ordinal);
            Assert.Contains(warnings, warning => warning.Contains("Unsupported Word shape", StringComparison.Ordinal));
        }

        [Fact]
        public void WordToMarkdown_Result_Reports_Fidelity_Loss() {
            using var document = WordDocument.Create();
            document.AddShape(ShapeType.Rectangle, 60, 30);

            WordToMarkdownResult result = document.ToMarkdownDocumentResult(new WordToMarkdownOptions {
                UnsupportedContentMode = MarkdownUnsupportedContentMode.Placeholder
            });

            Assert.True(result.Succeeded);
            Assert.True(result.HasLoss);
            Assert.Contains(result.Report.Diagnostics, diagnostic =>
                diagnostic.Code == "WordToMarkdownWarning" &&
                diagnostic.LossKind == WordMarkdownConversionLossKind.Approximation);
            Assert.Throws<WordMarkdownConversionException>(() => result.RequireNoLoss());
        }

        [Fact]
        public void WordToMarkdown_UnsupportedContentMode_Emits_Placeholders_For_Mixed_Paragraph_Content() {
            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph("Before");
            paragraph.AddShape(ShapeType.Rectangle, 60, 30);

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                UnsupportedContentMode = MarkdownUnsupportedContentMode.Placeholder
            });

            Assert.Contains("Before", markdown, StringComparison.Ordinal);
            Assert.Contains("Unsupported Word content: shape", markdown, StringComparison.Ordinal);
        }

        [Fact]
        public void WordToMarkdown_UnsupportedContentMode_Can_Emit_HtmlComments() {
            using var doc = WordDocument.Create();
            doc.AddShape(ShapeType.Rectangle, 60, 30);

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                UnsupportedContentMode = MarkdownUnsupportedContentMode.HtmlComment
            });

            Assert.Contains("<!-- Unsupported Word content: shape -->", markdown, StringComparison.Ordinal);
        }

        [Fact]
        public void MarkdownToWord_LoadFromMarkdown_Restores_PageBreak_Semantic_Blocks() {
            const string markdown = """
                Before

                ```officeimo-word-page-break
                ```

                After
                """;

            using var restored = OfficeIMO.Markdown.MarkdownReader.Parse(markdown, WordMarkdownSemanticBlocks.CreateReaderOptions()).ToWordDocument();

            Assert.Contains(restored.Paragraphs, paragraph => paragraph.Text.Trim() == "Before");
            Assert.Contains(restored.Paragraphs, paragraph => paragraph.PageBreak?.BreakType == BreakValues.Page);
            Assert.Contains(restored.Paragraphs, paragraph => paragraph.Text.Trim() == "After");
        }

        [Fact]
        public void WordMarkdownSemanticBlocks_CreateReaderOptions_Preserves_Loose_List_Item_Paragraphs() {
            const string markdown = """
                - first paragraph

                  second paragraph
                - next item
                """;

            MarkdownDoc parsed = OfficeIMO.Markdown.MarkdownReader.Parse(markdown, WordMarkdownSemanticBlocks.CreateReaderOptions());
            var list = Assert.IsType<UnorderedListBlock>(Assert.Single(parsed.Blocks));

            Assert.Collection(
                list.Items,
                item => Assert.Equal(new[] { "first paragraph", "second paragraph" }, item.ChildBlocks.Select(block => block.RenderMarkdown()).ToArray()),
                item => Assert.Equal(new[] { "next item" }, item.ChildBlocks.Select(block => block.RenderMarkdown()).ToArray()));
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

            using var restored = OfficeIMO.Markdown.MarkdownReader.Parse(markdown, WordMarkdownSemanticBlocks.CreateReaderOptions()).ToWordDocument();

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

        [Fact]
        public void MarkdownToWord_LoadFromMarkdownTemplate_Inserts_At_Bookmark_And_Preserves_Template_Content() {
            string templatePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
            try {
                using (var template = WordDocument.Create(templatePath)) {
                    template.AddParagraph("Before template content");
                    template.AddParagraph("PLACEHOLDER").AddBookmark("MainContent");
                    template.AddParagraph("After template content");
                    template.Save();
                }

                const string markdown = """
                    ---
                    title: Hidden metadata
                    ---
                    # Inserted heading

                    Inserted body.

                    - First
                    - Second
                    """;

                var options = new MarkdownToWordTemplateOptions { BookmarkName = "MainContent" };
                MarkdownDoc source = OfficeIMO.Markdown.MarkdownReader.Parse(markdown, options.CreateReaderOptions());
                WordDocument templateDocument = WordDocument.Load(templatePath);
                using var document = source.ToWordDocument(templateDocument, options);

                var paragraphTexts = document.Paragraphs
                    .Select(paragraph => paragraph.Text.Trim())
                    .Where(text => text.Length > 0)
                    .ToArray();

                Assert.DoesNotContain("PLACEHOLDER", paragraphTexts);
                Assert.DoesNotContain(paragraphTexts, text => text.Contains("title: Hidden metadata", StringComparison.Ordinal));
                Assert.True(Array.IndexOf(paragraphTexts, "Before template content") < Array.IndexOf(paragraphTexts, "Inserted heading"));
                Assert.True(Array.IndexOf(paragraphTexts, "Inserted body.") < Array.IndexOf(paragraphTexts, "After template content"));
                Assert.Contains(document.Paragraphs, paragraph => paragraph.Style == WordParagraphStyles.Heading1 && paragraph.Text.Trim() == "Inserted heading");
                Assert.Contains(document.Paragraphs, paragraph => paragraph.Text.Trim() == "First" && paragraph.IsListItem);
            } finally {
                if (File.Exists(templatePath)) {
                    File.Delete(templatePath);
                }
            }
        }

        [Fact]
        public void MarkdownToWord_LoadFromMarkdownTemplate_Realizes_Toc_Placeholders() {
            string templatePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
            try {
                using (var template = WordDocument.Create(templatePath)) {
                    template.AddParagraph("Before template content");
                    template.AddParagraph("PLACEHOLDER").AddBookmark("MainContent");
                    template.Save();
                }

                const string markdown = """
                    # Report

                    [TOC title="Contents" min=1 max=3]

                    ## Details
                    """;

                var options = new MarkdownToWordTemplateOptions { BookmarkName = "MainContent" };
                MarkdownDoc source = OfficeIMO.Markdown.MarkdownReader.Parse(markdown, options.CreateReaderOptions());
                WordDocument templateDocument = WordDocument.Load(templatePath);
                using var document = source.ToWordDocument(templateDocument, options);

                Assert.NotNull(document.TableOfContent);
                Assert.Equal("Contents", document.TableOfContent!.Text);
                Assert.Equal(1, document.TableOfContent.MinLevel);
                Assert.Equal(3, document.TableOfContent.MaxLevel);
            } finally {
                if (File.Exists(templatePath)) {
                    File.Delete(templatePath);
                }
            }
        }

        [Fact]
        public void WordToMarkdown_Emits_ExternalImages_As_Links_Instead_Of_Failing() {
            using var doc = WordDocument.Create();
            doc.AddParagraph().AddImage(new Uri("cid:86dec9c7-5eda-46b3-b8fb-2e2b7b0d6fb8"), 50, 50, description: "Linked image");
            var warnings = new List<string>();

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                ImageExportMode = ImageExportMode.File,
                OnWarning = warnings.Add
            });

            Assert.Contains("![Linked image](cid:86dec9c7-5eda-46b3-b8fb-2e2b7b0d6fb8)", markdown);
            Assert.Contains(warnings, warning => warning.Contains("Externally linked image", StringComparison.Ordinal));
        }

        [Fact]
        public void WordToMarkdown_Exports_Extensionless_ImageNames_With_Detected_Extension() {
            using var doc = WordDocument.Create();
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            using (var stream = File.OpenRead(imagePath)) {
                doc.AddParagraph().AddImage(stream, "Picture 1", width: 32, height: 32, description: "Logo");
            }

            string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);
            try {
                string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                    ImageExportMode = ImageExportMode.File,
                    ImageDirectory = tempDir
                });

                Assert.Contains("![Logo](Picture 1.png)", markdown);
                Assert.True(File.Exists(Path.Combine(tempDir, "Picture 1.png")));
            } finally {
                Directory.Delete(tempDir, recursive: true);
            }
        }

        [Fact]
        public void WordChart_TryGetSnapshot_Reads_Cached_Chart_Data() {
            using var doc = WordDocument.Create();
            var chart = doc.AddChart("Revenue", width: 400, height: 240);
            chart.AddCategories(new System.Collections.Generic.List<string> { "Q1", "Q2" });
            chart.AddBar("Actual", new System.Collections.Generic.List<int> { 10, 20 }, OfficeColor.CornflowerBlue);

            Assert.True(chart.TryGetSnapshot(out var snapshot));
            Assert.Equal("Revenue", snapshot.Title);
            Assert.Equal(WordChartSnapshotKind.ClusteredBar, snapshot.ChartKind);
            Assert.Equal(new[] { "Q1", "Q2" }, snapshot.Data.Categories);
            var series = Assert.Single(snapshot.Data.Series);
            Assert.Equal("Actual", series.Name);
            Assert.Equal(new[] { 10D, 20D }, series.Values);
            Assert.Equal(OfficeColor.CornflowerBlue, series.Color);
            Assert.True(snapshot.WidthPoints > 0);
            Assert.True(snapshot.HeightPoints > 0);
        }

        [Fact]
        public void WordChart_TryGetSnapshot_Uses_Max_Value_Count_For_Fallback_Categories() {
            using var doc = WordDocument.Create();
            var chart = doc.AddChart("Revenue", width: 400, height: 240);
            chart.AddCategories(new System.Collections.Generic.List<string> { "Q1", "Q2" });
            chart.AddBar("Actual", new System.Collections.Generic.List<int> { 10, 20 }, OfficeColor.CornflowerBlue);
            chart.AddBar("Forecast", new System.Collections.Generic.List<int> { 11, 21, 31 }, OfficeColor.SeaGreen);

            var chartPart = Assert.Single(doc.MainDocumentPartRoot.ChartParts);
            foreach (var categoryAxisData in chartPart.ChartSpace!.Descendants<C.CategoryAxisData>().ToList()) {
                categoryAxisData.Remove();
            }

            Assert.True(chart.TryGetSnapshot(out var snapshot));
            Assert.Equal(new[] { "Category 1", "Category 2", "Category 3" }, snapshot.Data.Categories);
            Assert.Equal(new[] { 10D, 20D, 0D }, snapshot.Data.Series[0].Values);
            Assert.Equal(new[] { 11D, 21D, 31D }, snapshot.Data.Series[1].Values);
        }

        [Fact]
        public void WordChart_TryGetSnapshot_Reads_Final_Palette_Series_Colors() {
            using var doc = WordDocument.Create();
            var chart = doc.AddChart("Regional Pipeline", width: 400, height: 240);
            chart.AddCategories(new System.Collections.Generic.List<string> { "Q1", "Q2" });
            chart.AddBar("EMEA", new System.Collections.Generic.List<int> { 10, 20 }, OfficeColor.CornflowerBlue);
            chart.AddBar("APAC", new System.Collections.Generic.List<int> { 8, 14 }, OfficeColor.SeaGreen);
            chart.AddBar("AMER", new System.Collections.Generic.List<int> { 12, 18 }, OfficeColor.Orange);
            chart.ApplyPalette(WordChart.WordChartPalette.ColorBlindSafe);

            Assert.True(chart.TryGetSnapshot(out var snapshot));

            Assert.Equal(OfficeColor.ParseHex("#0072B2"), snapshot.Data.Series[0].Color);
            Assert.Equal(OfficeColor.ParseHex("#E69F00"), snapshot.Data.Series[1].Color);
            Assert.Equal(OfficeColor.ParseHex("#009E73"), snapshot.Data.Series[2].Color);
        }

        [Fact]
        public void WordChart_TryGetSnapshot_Prefers_Later_Real_Categories_Over_Early_Fallback() {
            using var doc = WordDocument.Create();
            var chart = doc.AddChart("Regional Pipeline", width: 400, height: 240);
            chart.AddCategories(new System.Collections.Generic.List<string> { "Q1", "Q2" });
            chart.AddBar("No categories", new System.Collections.Generic.List<int> { 10, 20 }, OfficeColor.CornflowerBlue);
            chart.AddBar("Real categories", new System.Collections.Generic.List<int> { 8, 14 }, OfficeColor.SeaGreen);

            C.BarChartSeries firstSeries = doc._wordprocessingDocument.MainDocumentPart!
                .ChartParts
                .First()
                .ChartSpace
                .GetFirstChild<C.Chart>()!
                .PlotArea!
                .GetFirstChild<C.BarChart>()!
                .Elements<C.BarChartSeries>()
                .First();
            firstSeries.GetFirstChild<C.CategoryAxisData>()?.Remove();

            Assert.True(chart.TryGetSnapshot(out var snapshot));

            Assert.Equal(new[] { "Q1", "Q2" }, snapshot.Data.Categories);
        }

        [Fact]
        public void WordChart_TryGetSnapshot_Clamps_Inflated_Cache_PointCounts() {
            using var doc = WordDocument.Create();
            var chart = doc.AddChart("Inflated cache", width: 400, height: 240);
            chart.AddCategories(new System.Collections.Generic.List<string> { "Q1", "Q2" });
            chart.AddBar("Actual", new System.Collections.Generic.List<int> { 10, 20 }, OfficeColor.CornflowerBlue);

            C.ChartSpace chartSpace = doc._wordprocessingDocument.MainDocumentPart!
                .ChartParts
                .First()
                .ChartSpace;
            foreach (C.PointCount pointCount in chartSpace.Descendants<C.PointCount>()) {
                pointCount.Val = 1_000_000U;
            }

            Assert.True(chart.TryGetSnapshot(out var snapshot));

            Assert.Equal(new[] { "Q1", "Q2" }, snapshot.Data.Categories);
            var series = Assert.Single(snapshot.Data.Series);
            Assert.Equal(new[] { 10D, 20D }, series.Values);
        }

        [Fact]
        public void WordChart_TryGetSnapshot_Rejects_Mixed_Unsupported_Chart_Plots() {
            using var doc = WordDocument.Create();
            var chart = doc.AddChart("Mixed plots", width: 400, height: 240);
            chart.AddCategories(new System.Collections.Generic.List<string> { "Q1", "Q2" });
            chart.AddBar("Actual", new System.Collections.Generic.List<int> { 10, 20 }, OfficeColor.CornflowerBlue);

            C.PlotArea plotArea = doc._wordprocessingDocument.MainDocumentPart!
                .ChartParts
                .First()
                .ChartSpace
                .GetFirstChild<C.Chart>()!
                .PlotArea!;
            plotArea.Append(new C.BubbleChart());

            Assert.False(chart.TryGetSnapshot(out _));
        }

        [Fact]
        public void WordToMarkdown_VisualFallbackMode_Resolves_Theme_Series_Colors() {
            using var doc = WordDocument.Create();
            var chart = doc.AddChart("Theme Revenue", width: 400, height: 240);
            chart.AddCategories(new System.Collections.Generic.List<string> { "Q1", "Q2" });
            chart.AddBar("Actual", new System.Collections.Generic.List<int> { 10, 20 }, OfficeColor.Black);

            C.BarChartSeries seriesElement = doc._wordprocessingDocument.MainDocumentPart!
                .ChartParts
                .First()
                .ChartSpace
                .GetFirstChild<C.Chart>()!
                .PlotArea!
                .GetFirstChild<C.BarChart>()!
                .Elements<C.BarChartSeries>()
                .Single();
            C.ChartShapeProperties shapeProperties = seriesElement.GetFirstChild<C.ChartShapeProperties>()!;
            shapeProperties.RemoveAllChildren<A.SolidFill>();
            var schemeColor = new A.SchemeColor { Val = A.SchemeColorValues.Accent1 };
            schemeColor.Append(new A.LuminanceModulation { Val = 50000 });
            schemeColor.Append(new A.LuminanceOffset { Val = 20000 });
            schemeColor.Append(new A.AlphaModulation { Val = 50000 });
            shapeProperties.Append(new A.SolidFill(schemeColor));

            Assert.True(chart.TryGetSnapshot(out var snapshot));
            var series = Assert.Single(snapshot.Data.Series);
            Assert.Equal(OfficeColor.FromRgba(0x39, 0x63, 0xB1, 0x80), series.Color);

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                VisualFallbackMode = MarkdownVisualFallbackMode.SvgDataUri
            });

            const string prefix = "data:image/svg+xml;base64,";
            int sourceStart = markdown.IndexOf(prefix, StringComparison.Ordinal);
            Assert.True(sourceStart >= 0, markdown);
            int payloadStart = sourceStart + prefix.Length;
            int payloadEnd = markdown.IndexOf(')', payloadStart);
            Assert.True(payloadEnd > payloadStart, markdown);
            string svg = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(markdown.Substring(payloadStart, payloadEnd - payloadStart)));

            Assert.Contains("fill=\"#3963B1\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("fill-opacity=\"0.502\"", svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordToMarkdown_VisualFallbackMode_Resolves_Theme_Alias_Series_Colors() {
            using var doc = WordDocument.Create();
            var chart = doc.AddChart("Theme Alias Revenue", width: 400, height: 240);
            chart.AddCategories(new System.Collections.Generic.List<string> { "Q1", "Q2" });
            chart.AddBar("Actual", new System.Collections.Generic.List<int> { 10, 20 }, OfficeColor.Black);

            C.BarChartSeries seriesElement = doc._wordprocessingDocument.MainDocumentPart!
                .ChartParts
                .First()
                .ChartSpace
                .GetFirstChild<C.Chart>()!
                .PlotArea!
                .GetFirstChild<C.BarChart>()!
                .Elements<C.BarChartSeries>()
                .Single();
            C.ChartShapeProperties shapeProperties = seriesElement.GetFirstChild<C.ChartShapeProperties>()!;
            shapeProperties.RemoveAllChildren<A.SolidFill>();
            shapeProperties.Append(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Text2 }));

            Assert.True(chart.TryGetSnapshot(out var snapshot));
            var series = Assert.Single(snapshot.Data.Series);
            Assert.Equal(OfficeColor.ParseHex("#44546A"), series.Color);

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                VisualFallbackMode = MarkdownVisualFallbackMode.SvgDataUri
            });

            const string prefix = "data:image/svg+xml;base64,";
            int sourceStart = markdown.IndexOf(prefix, StringComparison.Ordinal);
            Assert.True(sourceStart >= 0, markdown);
            int payloadStart = sourceStart + prefix.Length;
            int payloadEnd = markdown.IndexOf(')', payloadStart);
            Assert.True(payloadEnd > payloadStart, markdown);
            string svg = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(markdown.Substring(payloadStart, payloadEnd - payloadStart)));

            Assert.Contains("fill=\"#44546A\"", svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordToMarkdown_VisualFallbackMode_Keeps_Sparse_Explicit_Chart_Colors_Aligned() {
            using var doc = WordDocument.Create();
            var chart = doc.AddChart("Sparse Color", width: 400, height: 240);
            chart.AddCategories(new System.Collections.Generic.List<string> { "Q1", "Q2" });
            chart.AddBar("Default", new System.Collections.Generic.List<int> { 10, 20 }, OfficeColor.Black);
            chart.AddBar("Explicit", new System.Collections.Generic.List<int> { 8, 14 }, OfficeColor.ParseHex("#CC3366"));

            var seriesElements = doc._wordprocessingDocument.MainDocumentPart!
                .ChartParts
                .First()
                .ChartSpace
                .GetFirstChild<C.Chart>()!
                .PlotArea!
                .GetFirstChild<C.BarChart>()!
                .Elements<C.BarChartSeries>()
                .ToList();
            C.ChartShapeProperties firstShapeProperties = seriesElements[0].GetFirstChild<C.ChartShapeProperties>()!;
            firstShapeProperties.RemoveAllChildren<A.SolidFill>();
            firstShapeProperties.RemoveAllChildren<A.Outline>();

            Assert.True(chart.TryGetSnapshot(out var snapshot));
            Assert.Null(snapshot.Data.Series[0].Color);
            Assert.Equal(OfficeColor.ParseHex("#CC3366"), snapshot.Data.Series[1].Color);

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                VisualFallbackMode = MarkdownVisualFallbackMode.SvgDataUri
            });

            const string prefix = "data:image/svg+xml;base64,";
            int sourceStart = markdown.IndexOf(prefix, StringComparison.Ordinal);
            Assert.True(sourceStart >= 0, markdown);
            int payloadStart = sourceStart + prefix.Length;
            int payloadEnd = markdown.IndexOf(')', payloadStart);
            Assert.True(payloadEnd > payloadStart, markdown);
            string svg = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(markdown.Substring(payloadStart, payloadEnd - payloadStart)));

            string defaultColor = "#" + OfficeChartDrawingRenderer.GetSeriesColor(0).ToRgbHex();
            Assert.False(string.Equals("#CC3366", defaultColor, StringComparison.OrdinalIgnoreCase));
            Assert.Contains("fill=\"" + defaultColor + "\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("fill=\"#CC3366\"", svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordToMarkdown_VisualFallbackMode_Uses_Line_Outline_Color() {
            using var doc = WordDocument.Create();
            var chart = doc.AddChart("Trend", width: 400, height: 240);
            chart.AddCategories(new System.Collections.Generic.List<string> { "Q1", "Q2", "Q3" });
            chart.AddLine("Actual", new System.Collections.Generic.List<int> { 10, 20, 15 }, OfficeColor.Black);

            C.LineChartSeries seriesElement = doc._wordprocessingDocument.MainDocumentPart!
                .ChartParts
                .First()
                .ChartSpace
                .GetFirstChild<C.Chart>()!
                .PlotArea!
                .GetFirstChild<C.LineChart>()!
                .Elements<C.LineChartSeries>()
                .Single();
            C.ChartShapeProperties shapeProperties = seriesElement.GetFirstChild<C.ChartShapeProperties>()!;
            shapeProperties.RemoveAllChildren<A.SolidFill>();
            shapeProperties.RemoveAllChildren<A.Outline>();
            shapeProperties.Append(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "CC3366" })));

            Assert.True(chart.TryGetSnapshot(out var snapshot));
            var series = Assert.Single(snapshot.Data.Series);
            Assert.Equal(OfficeColor.ParseHex("#CC3366"), series.Color);

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                VisualFallbackMode = MarkdownVisualFallbackMode.SvgDataUri
            });

            const string prefix = "data:image/svg+xml;base64,";
            int sourceStart = markdown.IndexOf(prefix, StringComparison.Ordinal);
            Assert.True(sourceStart >= 0, markdown);
            int payloadStart = sourceStart + prefix.Length;
            int payloadEnd = markdown.IndexOf(')', payloadStart);
            Assert.True(payloadEnd > payloadStart, markdown);
            string svg = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(markdown.Substring(payloadStart, payloadEnd - payloadStart)));

            Assert.Contains("stroke=\"#CC3366\"", svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordToMarkdown_VisualFallbackMode_Uses_Pie_Point_Colors() {
            using var doc = WordDocument.Create();
            var chart = doc.AddChart("Rules outcome", width: 400, height: 240);
            chart.AddPie("Passed", 42);
            chart.AddPie("Failed", 30);
            chart.AddPie("Skipped", 5);
            chart.ApplyPalette(WordChart.WordChartPalette.Professional, semanticOutcomes: true, applyToPies: true, applyToSeries: false);

            Assert.True(chart.TryGetSnapshot(out var snapshot));
            var series = Assert.Single(snapshot.Data.Series);
            Assert.NotNull(series.PointColors);
            Assert.Equal(OfficeColor.ParseHex("#2fb344"), series.PointColors![0]);
            Assert.Equal(OfficeColor.ParseHex("#f76707"), series.PointColors[1]);
            Assert.Equal(OfficeColor.ParseHex("#868e96"), series.PointColors[2]);

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                VisualFallbackMode = MarkdownVisualFallbackMode.SvgDataUri
            });

            const string prefix = "data:image/svg+xml;base64,";
            int sourceStart = markdown.IndexOf(prefix, StringComparison.Ordinal);
            Assert.True(sourceStart >= 0, markdown);
            int payloadStart = sourceStart + prefix.Length;
            int payloadEnd = markdown.IndexOf(')', payloadStart);
            Assert.True(payloadEnd > payloadStart, markdown);
            string svg = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(markdown.Substring(payloadStart, payloadEnd - payloadStart)));

            Assert.Contains("fill=\"#2FB344\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("fill=\"#F76707\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("fill=\"#868E96\"", svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordToMarkdown_VisualFallbackMode_Renders_Charts_As_Svg_DataUri_Images() {
            using var doc = WordDocument.Create();
            var chart = doc.AddChart("Revenue", width: 400, height: 240);
            chart.AddCategories(new System.Collections.Generic.List<string> { "Q1", "Q2" });
            chart.AddBar("Actual", new System.Collections.Generic.List<int> { 10, 20 }, OfficeColor.CornflowerBlue);
            var warnings = new System.Collections.Generic.List<string>();

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                VisualFallbackMode = MarkdownVisualFallbackMode.SvgDataUri,
                OnWarning = warnings.Add
            });

            const string prefix = "data:image/svg+xml;base64,";
            int sourceStart = markdown.IndexOf(prefix, StringComparison.Ordinal);
            Assert.True(sourceStart >= 0, markdown);
            int payloadStart = sourceStart + prefix.Length;
            int payloadEnd = markdown.IndexOf(')', payloadStart);
            Assert.True(payloadEnd > payloadStart, markdown);
            string svg = System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(markdown.Substring(payloadStart, payloadEnd - payloadStart)));

            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("<rect", svg, StringComparison.Ordinal);
            Assert.Contains("fill=\"#6495ED\"", svg, StringComparison.Ordinal);
            Assert.Contains(warnings, warning => warning.Contains("SVG Markdown image fallback", StringComparison.Ordinal));
        }

        [Fact]
        public void SaveAsMarkdown_VisualFallbackMode_Writes_Chart_Svg_Sidecar_Resources() {
            string tempDir = Path.Combine(Path.GetTempPath(), "OfficeIMO-Markdown-ChartResources-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            try {
                string markdownPath = Path.Combine(tempDir, "Report.md");
                using var doc = WordDocument.Create();
                var chart = doc.AddChart("Regional Pipeline", width: 400, height: 240);
                chart.AddCategories(new System.Collections.Generic.List<string> { "Q1", "Q2" });
                chart.AddBar("Actual", new System.Collections.Generic.List<int> { 10, 20 }, OfficeColor.CornflowerBlue);

                doc.SaveAsMarkdown(markdownPath, new WordToMarkdownOptions {
                    VisualFallbackMode = MarkdownVisualFallbackMode.SvgFile
                });

                string markdown = File.ReadAllText(markdownPath);
                string resourcePath = Path.Combine(tempDir, "Report.assets", "01-regional-pipeline.svg");
                Assert.Contains("![Regional Pipeline](Report.assets/01-regional-pipeline.svg)", markdown);
                Assert.DoesNotContain("data:image/svg+xml;base64", markdown, StringComparison.Ordinal);
                Assert.True(File.Exists(resourcePath), "Expected chart SVG resource at " + resourcePath);
                string svg = File.ReadAllText(resourcePath);
                Assert.Contains("<svg", svg, StringComparison.Ordinal);
                Assert.Contains("Regional Pipeline", svg, StringComparison.Ordinal);
            } finally {
                Directory.Delete(tempDir, recursive: true);
            }
        }

        [Fact]
        public void SaveAsMarkdown_VisualFallbackMode_Derives_Sidecar_Resources_Per_Save_When_Options_Are_Reused() {
            string tempDir = Path.Combine(Path.GetTempPath(), "OfficeIMO-Markdown-ReusedChartResources-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDir);
            try {
                using var doc = WordDocument.Create();
                var chart = doc.AddChart("Reusable Options", width: 400, height: 240);
                chart.AddCategories(new System.Collections.Generic.List<string> { "Q1", "Q2" });
                chart.AddBar("Actual", new System.Collections.Generic.List<int> { 10, 20 }, OfficeColor.CornflowerBlue);

                var options = new WordToMarkdownOptions {
                    VisualFallbackMode = MarkdownVisualFallbackMode.SvgFile
                };
                string firstPath = Path.Combine(tempDir, "First.md");
                string secondPath = Path.Combine(tempDir, "Second.md");

                doc.SaveAsMarkdown(firstPath, options);
                doc.SaveAsMarkdown(secondPath, options);

                Assert.Null(options.VisualFallbackDirectory);
                Assert.Null(options.VisualFallbackPathPrefix);
                Assert.True(File.Exists(Path.Combine(tempDir, "First.assets", "01-reusable-options.svg")));
                Assert.True(File.Exists(Path.Combine(tempDir, "Second.assets", "01-reusable-options.svg")));
                Assert.Contains("First.assets/01-reusable-options.svg", File.ReadAllText(firstPath), StringComparison.Ordinal);
                Assert.Contains("Second.assets/01-reusable-options.svg", File.ReadAllText(secondPath), StringComparison.Ordinal);
            } finally {
                Directory.Delete(tempDir, recursive: true);
            }
        }

        [Fact]
        public void WordToMarkdown_VisualFallbackMode_Default_Does_Not_Render_Charts() {
            using var doc = WordDocument.Create();
            var chart = doc.AddChart("Revenue", width: 400, height: 240);
            chart.AddCategories(new System.Collections.Generic.List<string> { "Q1", "Q2" });
            chart.AddBar("Actual", new System.Collections.Generic.List<int> { 10, 20 }, OfficeColor.CornflowerBlue);

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions());

            Assert.DoesNotContain("data:image/svg+xml;base64", markdown, StringComparison.Ordinal);
        }
    }
}
