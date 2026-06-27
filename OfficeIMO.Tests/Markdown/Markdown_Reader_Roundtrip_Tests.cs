using System;
using System.Linq;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite {
    public class Markdown_Reader_Roundtrip_Tests {
        [Fact]
        public void Reader_Roundtrips_Basic_Document() {
            var md = MarkdownDoc.Create()
                .FrontMatter(new { title = "Doc", tags = new [] { "a", "b" }, published = true })
                .H1("Doc")
                .P("Intro paragraph with a link to docs.")
                .H2("Install")
                .Code("bash", "dotnet tool install -g Example")
                .Caption("Install the global tool")
                .H2("Features")
                .Ul(ul => ul
                    .Item("Tables from sequences")
                    .Item("Callouts and TOC")
                    .ItemTask("Works offline", true))
                .Image("https://example.com/logo.png", alt: "Logo", title: "Example").Caption("Our logo")
                .Table(t => t.Headers("Col1", "Col2").Row("A", "1").Row("B", "2").Align(ColumnAlignment.Left, ColumnAlignment.Right))
                .Callout("info", "Heads up", "Early access APIs may change.")
                .Dl(d => d.Item("Term", "Definition"));

            var text = md.ToMarkdown();
            var parsed = MarkdownReader.Parse(text);

            // Basic block shape assertions
            Assert.True(parsed.Blocks.Count >= 7);
            Assert.IsType<HeadingBlock>(parsed.Blocks[0] is FrontMatterBlock ? parsed.Blocks[1] : parsed.Blocks[0]);

            // Find code block and validate caption
            var code = parsed.Blocks.OfType<CodeBlock>().FirstOrDefault();
            Assert.NotNull(code);
            Assert.Equal("bash", code!.Language);
            Assert.Equal("Install the global tool", code.Caption);

            // Validate image + caption
            var img = parsed.Blocks.OfType<ImageBlock>().FirstOrDefault();
            Assert.NotNull(img);
            Assert.Contains("logo.png", img!.Path);
            Assert.Equal("Logo", img.Alt);
            Assert.Equal("Example", img.Title);
            Assert.Equal("Our logo", img.Caption);
            Assert.Null(img.LinkUrl);

            // Validate table header/alignments
            var table = parsed.Blocks.OfType<TableBlock>().FirstOrDefault();
            Assert.NotNull(table);
            Assert.Equal(new [] { "Col1", "Col2" }, table!.Headers);
            Assert.Equal(new [] { ColumnAlignment.Left, ColumnAlignment.Right }, table.Alignments);

            // Validate list kinds: unordered with a task item
            var ul = parsed.Blocks.OfType<UnorderedListBlock>().FirstOrDefault();
            Assert.NotNull(ul);
            Assert.True(ul!.Items.Count >= 3);
        }

        [Fact]
        public void Reader_Roundtrips_Linked_Image_Block() {
            var md = MarkdownDoc.Create()
                .Add(new ImageBlock("https://example.com/logo.png", "Logo", "Example", linkUrl: "https://example.com/docs", linkTitle: "Documentation"))
                .Caption("Our linked logo");

            var text = md.ToMarkdown();
            var parsed = MarkdownReader.Parse(text);

            var img = Assert.IsType<ImageBlock>(Assert.Single(parsed.Blocks));
            Assert.Equal("https://example.com/logo.png", img.Path);
            Assert.Equal("Logo", img.Alt);
            Assert.Equal("Example", img.Title);
            Assert.Equal("https://example.com/docs", img.LinkUrl);
            Assert.Equal("Documentation", img.LinkTitle);
            Assert.Equal("Our linked logo", img.Caption);
        }

        [Fact]
        public void Document_Add_Treats_FrontMatter_As_Document_Header_Block() {
            var md = MarkdownDoc.Create()
                .Add(FrontMatterBlock.FromObject(new { title = "Doc", published = true }))
                .H1("Doc");

            var text = md.ToMarkdown().Replace("\r", "");

            Assert.StartsWith("---\n", text);
            Assert.Contains("title: Doc", text);
            Assert.Contains("published: true", text);
            Assert.Contains("\n\n# Doc", text);
            Assert.Single(md.Blocks);
            Assert.IsType<HeadingBlock>(md.Blocks[0]);
        }

        [Fact]
        public void Reader_Exposes_FrontMatter_Entries_As_Structured_Data() {
            const string markdown = """
---
title: Doc
published: true
tags: [a, b]
---

# Heading
""";

            var parsed = MarkdownReader.Parse(markdown);
            var frontMatter = Assert.IsType<FrontMatterBlock>(parsed.DocumentHeader!);

            Assert.Equal(frontMatter.Entries, parsed.FrontMatterEntries);

            Assert.Collection(
                frontMatter.Entries,
                entry => {
                    Assert.Equal("title", entry.Key);
                    Assert.Equal("Doc", Assert.IsType<string>(entry.Value));
                },
                entry => {
                    Assert.Equal("published", entry.Key);
                    Assert.True(Assert.IsType<bool>(entry.Value));
                },
                entry => {
                    Assert.Equal("tags", entry.Key);
                    var tags = Assert.IsAssignableFrom<IEnumerable<string>>(entry.Value);
                    Assert.Equal(new[] { "a", "b" }, tags.ToArray());
                });

            var published = frontMatter.FindEntry("published");
            Assert.NotNull(published);
            Assert.True(Assert.IsType<bool>(published!.Value));

            Assert.True(frontMatter.TryGetValue<string>("title", out var title));
            Assert.Equal("Doc", title);

            Assert.True(frontMatter.TryGetValue<bool>("published", out var isPublished));
            Assert.True(isPublished);

            Assert.True(frontMatter.TryGetValue<IEnumerable<string>>("tags", out var tagValues));
            Assert.Equal(new[] { "a", "b" }, tagValues!.ToArray());

            Assert.True(frontMatter.HasEntry("published"));
            Assert.True(frontMatter.HasEntry("Published"));
            Assert.False(frontMatter.HasEntry("Published", StringComparison.Ordinal));
            Assert.False(frontMatter.TryGetValue<int>("title", out _));
            Assert.Null(frontMatter.FindEntry("missing"));
            Assert.False(frontMatter.HasEntry("missing"));

            Assert.True(parsed.HasDocumentHeader);
            Assert.True(parsed.HasFrontMatterEntry("published"));
            Assert.True(parsed.HasFrontMatterEntry("Published"));
            Assert.False(parsed.HasFrontMatterEntry("Published", StringComparison.Ordinal));
            Assert.Equal("published", parsed.FindFrontMatterEntry("published")!.Key);
            Assert.True(parsed.TryGetFrontMatterValue<string>("title", out var documentTitle));
            Assert.Equal("Doc", documentTitle);
            Assert.False(parsed.TryGetFrontMatterValue<int>("title", out _));
            Assert.False(parsed.HasFrontMatterEntry("missing"));
        }

        [Fact]
        public void Reader_Exposes_TopLevel_Blocks_In_Document_Order() {
            const string markdown = """
---
title: Doc
---

# Heading

Paragraph
""";

            var parsed = MarkdownReader.Parse(markdown);

            Assert.Collection(
                parsed.TopLevelBlocks,
                block => Assert.IsType<FrontMatterBlock>(block),
                block => Assert.IsType<HeadingBlock>(block),
                block => Assert.IsType<ParagraphBlock>(block));

            Assert.Collection(
                parsed.Blocks,
                block => Assert.IsType<HeadingBlock>(block),
                block => Assert.IsType<ParagraphBlock>(block));

            Assert.Collection(
                parsed.TopLevelBlocksOfType<HeadingBlock>(),
                block => Assert.Equal("Heading", block.Text));

            Assert.Collection(
                parsed.TopLevelBlocksOfType<FrontMatterBlock>(),
                block => Assert.Equal("Doc", block.Entries[0].Value));

            var topLevelHeading = parsed.FindFirstTopLevelBlockOfType<HeadingBlock>();
            Assert.NotNull(topLevelHeading);
            Assert.Equal("Heading", topLevelHeading!.Text);

            Assert.True(parsed.HasTopLevelBlockOfType<HeadingBlock>());
            Assert.True(parsed.HasTopLevelBlockOfType<FrontMatterBlock>());
            Assert.False(parsed.HasTopLevelBlockOfType<QuoteBlock>());
        }

        [Fact]
        public void Reader_Enumerates_Blocks_Depth_First() {
            const string markdown = """
---
title: Doc
---

> Quote
>
> - item

Paragraph
""";

            var parsed = MarkdownReader.Parse(markdown);
            var kinds = parsed.DescendantsAndSelf().Select(block => block.GetType()).ToArray();

            Assert.Equal(
                new[] {
                    typeof(FrontMatterBlock),
                    typeof(QuoteBlock),
                    typeof(ParagraphBlock),
                    typeof(UnorderedListBlock),
                    typeof(ParagraphBlock),
                    typeof(ParagraphBlock)
                },
                kinds);

            Assert.Equal(
                new[] { "Quote", "item", "Paragraph" },
                parsed.DescendantsOfType<ParagraphBlock>()
                    .Select(block => block.Inlines.RenderMarkdown())
                    .ToArray());

            var firstNestedList = parsed.FindFirstDescendantOfType<UnorderedListBlock>();
            Assert.NotNull(firstNestedList);
            Assert.Equal("item", firstNestedList!.Items[0].Content.RenderMarkdown());

            Assert.True(parsed.HasDescendantOfType<UnorderedListBlock>());
            Assert.True(parsed.HasDescendantOfType<ParagraphBlock>());
            Assert.False(parsed.HasDescendantOfType<TableBlock>());
        }

        [Fact]
        public void Reader_Enumerates_List_Items_In_Document_Order() {
            const string markdown = """
- outer
  - inner

> - quoted
""";

            var parsed = MarkdownReader.Parse(markdown);
            var items = parsed.DescendantListItems().ToArray();

            Assert.Equal(3, items.Length);
            Assert.Equal("outer", items[0].Content.RenderMarkdown());
            Assert.Equal("inner", items[1].Content.RenderMarkdown());
            Assert.Equal("quoted", items[2].Content.RenderMarkdown());

            var topList = Assert.IsType<UnorderedListBlock>(parsed.Blocks[0]);
            Assert.Equal(topList.Items, topList.ListItems);

            var quotedList = Assert.IsType<UnorderedListBlock>(Assert.IsType<QuoteBlock>(parsed.Blocks[1]).ChildBlocks[0]);
            Assert.Equal(quotedList.Items, quotedList.ListItems);
        }

        [Fact]
        public void Reader_Enumerates_Typed_Table_Cell_Blocks_Depth_First() {
            const string markdown = """
| Col1 | Col2 |
| --- | --- |
| A | B |
""";

            var parsed = MarkdownReader.Parse(markdown);
            var kinds = parsed.DescendantsAndSelf().Select(block => block.GetType()).ToArray();

            Assert.Equal(
                new[] {
                    typeof(TableBlock),
                    typeof(ParagraphBlock),
                    typeof(ParagraphBlock),
                    typeof(ParagraphBlock),
                    typeof(ParagraphBlock)
                },
                kinds);

            Assert.Equal(
                new[] { "Col1", "Col2", "A", "B" },
                parsed.DescendantsOfType<ParagraphBlock>()
                    .Select(block => block.Inlines.RenderMarkdown())
                    .ToArray());
        }

        [Fact]
        public void Reader_Roundtrips_Structured_Table_Cell_Block_Content() {
            const string markdown = """
| Section | Notes |
| --- | --- |
| Alpha | Intro<br><br>> Quoted<br><br>- first<br>- second |
""";

            var parsed = MarkdownReader.Parse(markdown);
            var table = Assert.IsType<TableBlock>(Assert.Single(parsed.Blocks));
            Assert.Collection(table.RowCells[0][1].Blocks,
                block => Assert.Equal("Intro", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => Assert.IsType<QuoteBlock>(block),
                block => {
                    var list = Assert.IsType<UnorderedListBlock>(block);
                    Assert.Equal(new[] { "first", "second" }, list.Items.Select(item => item.Content.RenderMarkdown()).ToArray());
                });

            var roundtrip = parsed.ToMarkdown().Replace("\r\n", "\n");
            Assert.Contains("Intro<br><br>> Quoted<br><br>- first<br>- second", roundtrip, StringComparison.Ordinal);

            var reparsed = MarkdownReader.Parse(roundtrip);
            var reparsedTable = Assert.IsType<TableBlock>(Assert.Single(reparsed.Blocks));
            Assert.Collection(reparsedTable.RowCells[0][1].Blocks,
                block => Assert.Equal("Intro", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => Assert.IsType<QuoteBlock>(block),
                block => {
                    var list = Assert.IsType<UnorderedListBlock>(block);
                    Assert.Equal(new[] { "first", "second" }, list.Items.Select(item => item.Content.RenderMarkdown()).ToArray());
                });
        }

        [Fact]
        public void Reader_Roundtrips_Structured_Table_Cell_Code_Block_Content() {
            const string markdown = """
| Section | Notes |
| --- | --- |
| Alpha | Intro<br><br>```text<br>code line 1<br>code line 2<br>``` |
""";

            var parsed = MarkdownReader.Parse(markdown);
            var table = Assert.IsType<TableBlock>(Assert.Single(parsed.Blocks));
            Assert.Collection(table.RowCells[0][1].Blocks,
                block => Assert.Equal("Intro", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => {
                    var code = Assert.IsType<CodeBlock>(block);
                    Assert.Equal("text", code.Language);
                    Assert.Contains("code line 1", code.Content, StringComparison.Ordinal);
                    Assert.Contains("code line 2", code.Content, StringComparison.Ordinal);
                });

            var roundtrip = parsed.ToMarkdown().Replace("\r\n", "\n");
            Assert.Contains("```text<br>code line 1<br>code line 2<br>```", roundtrip, StringComparison.Ordinal);

            var reparsed = MarkdownReader.Parse(roundtrip);
            var reparsedTable = Assert.IsType<TableBlock>(Assert.Single(reparsed.Blocks));
            Assert.Collection(reparsedTable.RowCells[0][1].Blocks,
                block => Assert.Equal("Intro", Assert.IsType<ParagraphBlock>(block).Inlines.RenderMarkdown()),
                block => {
                    var code = Assert.IsType<CodeBlock>(block);
                    Assert.Equal("text", code.Language);
                    Assert.Contains("code line 1", code.Content, StringComparison.Ordinal);
                    Assert.Contains("code line 2", code.Content, StringComparison.Ordinal);
                });
        }

        [Fact]
        public void Reader_Roundtrips_SingleLine_Structured_Table_Cell_Content() {
            const string markdown = """
| Notes | Extra |
| --- | --- |
| ## Important | - first |
""";

            var parsed = MarkdownReader.Parse(markdown);
            var table = Assert.IsType<TableBlock>(Assert.Single(parsed.Blocks));
            Assert.IsType<HeadingBlock>(Assert.Single(table.RowCells[0][0].Blocks));
            Assert.IsType<UnorderedListBlock>(Assert.Single(table.RowCells[0][1].Blocks));

            var roundtrip = parsed.ToMarkdown().Replace("\r\n", "\n");
            Assert.Contains("| ## Important | - first |", roundtrip, StringComparison.Ordinal);

            var reparsed = MarkdownReader.Parse(roundtrip);
            var reparsedTable = Assert.IsType<TableBlock>(Assert.Single(reparsed.Blocks));
            Assert.IsType<HeadingBlock>(Assert.Single(reparsedTable.RowCells[0][0].Blocks));
            Assert.IsType<UnorderedListBlock>(Assert.Single(reparsedTable.RowCells[0][1].Blocks));
        }

        [Fact]
        public void MarkdownRoundtripWriter_Preserves_Unchanged_OriginalMarkdown_When_PreserveTrivia_Is_Enabled() {
            const string markdown = "# Title\r\n\r\nParagraph one\rSecond para";
            var options = new MarkdownReaderOptions {
                PreserveTrivia = true
            };

            var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

            var roundtrip = MarkdownRoundtripWriter.WriteUnchanged(result);

            Assert.True(roundtrip.IsLossless);
            Assert.Empty(roundtrip.Diagnostics);
            Assert.Equal(markdown, roundtrip.Markdown);
        }

        [Fact]
        public void MarkdownRoundtripWriter_Reports_Fallback_When_OriginalMarkdown_Was_Not_Preserved() {
            const string markdown = "# Title\r\n";
            var result = MarkdownReader.ParseWithSyntaxTree(markdown);

            var roundtrip = MarkdownRoundtripWriter.WriteUnchanged(result);

            Assert.False(roundtrip.IsLossless);
            var diagnostic = Assert.Single(roundtrip.Diagnostics);
            Assert.Equal("roundtrip.preserve-trivia-required", diagnostic.Id);
            Assert.Equal(result.Document.ToMarkdown(), roundtrip.Markdown);
        }

        [Fact]
        public void MarkdownRoundtripWriter_Reports_Fallback_When_DocumentTransforms_Changed_Result() {
            const string markdown = "previous shutdown was unexpected### Reason";
            var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
            options.PreserveTrivia = true;
            options.DocumentTransforms.Add(new MarkdownCompactHeadingBoundaryTransform());

            var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, options);

            var roundtrip = MarkdownRoundtripWriter.WriteUnchanged(result);

            Assert.False(roundtrip.IsLossless);
            var diagnostic = Assert.Single(roundtrip.Diagnostics);
            Assert.Equal("roundtrip.document-transformed", diagnostic.Id);
            Assert.Equal(result.TransformDiagnostics[0].AffectedSourceSpan, diagnostic.SourceSpan);
            Assert.Equal(result.Document.ToMarkdown(), roundtrip.Markdown);
            Assert.NotEqual(result.OriginalMarkdown, roundtrip.Markdown);
        }

        [Fact]
        public void MarkdownRoundtripWriter_Reports_Transform_RelatedSourceSpans_When_UnchangedWrite_Falls_Back() {
            var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
            options.PreserveTrivia = true;
            options.DocumentTransforms.Add(new UppercaseParagraphsTransform());

            var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
alpha

beta
""", options);

            var roundtrip = MarkdownRoundtripWriter.WriteUnchanged(result);

            Assert.False(roundtrip.IsLossless);
            var diagnostic = Assert.Single(roundtrip.Diagnostics);
            Assert.Equal("roundtrip.document-transformed", diagnostic.Id);
            Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 4), diagnostic.SourceSpan);
            Assert.Equal(new[] { 1, 3 }, diagnostic.RelatedSourceSpans.Select(span => span.StartLine).ToArray());
            Assert.Equal(new[] { 1, 3 }, diagnostic.RelatedSourceSpans.Select(span => span.EndLine).ToArray());
        }

        [Fact]
        public void MarkdownRoundtripWriter_Preserves_Unchanged_OriginalMarkdown_When_NoOpTransform_Ran() {
            const string markdown = "# Title\r\n\r\nBody\r\n";
            var options = new MarkdownReaderOptions {
                PreserveTrivia = true
            };
            options.DocumentTransforms.Add(new NoOpTransform());

            var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, options);

            var transformDiagnostic = Assert.Single(result.TransformDiagnostics);
            Assert.False(transformDiagnostic.HasChangedBlocks);
            var roundtrip = MarkdownRoundtripWriter.WriteUnchanged(result);

            Assert.True(roundtrip.IsLossless);
            Assert.Empty(roundtrip.Diagnostics);
            Assert.Equal(markdown, roundtrip.Markdown);
        }

        [Fact]
        public void MarkdownRoundtripWriter_Applies_SourceEdit_To_OriginalMarkdown_When_PreserveTrivia_Is_Enabled() {
            const string markdown = "# Old **Title**\r\n\r\nBody\r\n";
            var options = new MarkdownReaderOptions {
                PreserveTrivia = true
            };
            var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, options);
            var native = MarkdownNativeDocument.FromParseResult(result);
            var heading = Assert.IsType<MarkdownNativeHeadingBlock>(native.Blocks[0]);
            var edit = native.CreateReplaceEdit(heading.TextSourceSpan!.Value, "New Title");

            var roundtrip = MarkdownRoundtripWriter.WriteWithSourceEdit(result, edit);

            Assert.True(roundtrip.IsLossless);
            Assert.Empty(roundtrip.Diagnostics);
            Assert.Equal("# New Title\r\n\r\nBody\r\n", roundtrip.Markdown);
        }

        [Fact]
        public void MarkdownRoundtripWriter_Applies_SourceEdit_To_OriginalMarkdown_When_NoOpTransform_Ran() {
            const string markdown = "# Old **Title**\r\n\r\nBody\r\n";
            var options = new MarkdownReaderOptions {
                PreserveTrivia = true
            };
            options.DocumentTransforms.Add(new NoOpTransform());
            var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, options);
            var native = MarkdownNativeDocument.FromParseResult(result);
            var heading = Assert.IsType<MarkdownNativeHeadingBlock>(native.Blocks[0]);
            var edit = native.CreateReplaceEdit(heading.TextSourceSpan!.Value, "New Title");

            var roundtrip = MarkdownRoundtripWriter.WriteWithSourceEdit(result, edit);

            Assert.True(roundtrip.IsLossless);
            Assert.Empty(roundtrip.Diagnostics);
            Assert.Equal("# New Title\r\n\r\nBody\r\n", roundtrip.Markdown);
        }

        [Fact]
        public void MarkdownRoundtripWriter_Applies_SourceEdits_To_Multiline_Setext_Heading_Tokens() {
            const string markdown = "Foo *bar\r\nbaz*\r\n====\r\n\r\nBody\r\n";
            var options = new MarkdownReaderOptions {
                PreserveTrivia = true
            };
            var native = MarkdownNativeDocument.Parse(markdown, options);
            var heading = Assert.IsType<MarkdownNativeHeadingBlock>(native.Blocks[0]);

            Assert.Equal(1, heading.Level);
            Assert.Equal("Foo bar baz", heading.Text);
            Assert.Equal(new MarkdownSourceSpan(3, 1, 3, 4), heading.LevelSourceSpan);
            Assert.Equal(new MarkdownSourceSpan(1, 1, 2, 4), heading.TextSourceSpan);

            var textEdit = native.CreateReplaceEdit(heading.TextSourceSpan!.Value, "New **Title**");
            var textRoundtrip = native.WriteWithSourceEdit(textEdit);
            Assert.True(textRoundtrip.IsLossless);
            Assert.Empty(textRoundtrip.Diagnostics);
            Assert.Equal("New **Title**\r\n====\r\n\r\nBody\r\n", textRoundtrip.Markdown);

            var levelEdit = native.CreateReplaceEdit(heading.LevelSourceSpan!.Value, "---");
            var levelRoundtrip = native.WriteWithSourceEdit(levelEdit);
            Assert.True(levelRoundtrip.IsLossless);
            Assert.Empty(levelRoundtrip.Diagnostics);
            Assert.Equal("Foo *bar\r\nbaz*\r\n---\r\n\r\nBody\r\n", levelRoundtrip.Markdown);
        }

        [Fact]
        public void MarkdownNativeDocument_Applies_Shuffled_SourceEdits_To_OriginalMarkdown() {
            const string markdown = "# Old **Title**\r\n\r\nSee [docs](old.md \"Old title\") and `code`.\r\n";
            var options = new MarkdownReaderOptions {
                PreserveTrivia = true
            };
            var native = MarkdownNativeDocument.Parse(markdown, options);
            var heading = Assert.IsType<MarkdownNativeHeadingBlock>(native.Blocks[0]);
            var link = Assert.Single(native.EnumerateInlines(), inline => inline.Kind == MarkdownNativeInlineKind.Link);
            var target = Assert.Single(link.Metadata, metadata => metadata.Name == "target");
            var title = Assert.Single(link.Metadata, metadata => metadata.Name == "title");

            var edits = new[] {
                native.CreateReplaceEdit(title, "New docs title"),
                native.CreateReplaceEdit(heading.TextSourceSpan!.Value, "New **Title**"),
                native.CreateReplaceEdit(target, "new/location.md")
            };

            var roundtrip = native.WriteWithSourceEdits(edits);

            Assert.True(roundtrip.IsLossless);
            Assert.Empty(roundtrip.Diagnostics);
            Assert.Equal("# New **Title**\r\n\r\nSee [docs](new/location.md \"New docs title\") and `code`.\r\n", roundtrip.Markdown);
        }

        [Fact]
        public void MarkdownRoundtripWriter_Applies_SourceEdit_To_NormalizedMarkdown_When_OriginalMarkdown_Was_Not_Preserved() {
            const string markdown = "# Old **Title**\r\n\r\nBody\r\n";
            var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown);
            var native = MarkdownNativeDocument.FromParseResult(result);
            var heading = Assert.IsType<MarkdownNativeHeadingBlock>(native.Blocks[0]);
            var edit = native.CreateReplaceEdit(heading.TextSourceSpan!.Value, "New Title");

            var roundtrip = MarkdownRoundtripWriter.WriteWithSourceEdit(result, edit);

            Assert.False(roundtrip.IsLossless);
            var diagnostic = Assert.Single(roundtrip.Diagnostics);
            Assert.Equal("roundtrip.preserve-trivia-required", diagnostic.Id);
            Assert.Equal(edit.SourceSpan, diagnostic.SourceSpan);
            Assert.Equal("# New Title\n\nBody\n", roundtrip.Markdown);
        }

        [Fact]
        public void MarkdownRoundtripWriter_Reports_Transform_SourceSpan_When_SourceEdits_Fallback_To_NormalizedMarkdown() {
            const string markdown = "previous shutdown was unexpected### Reason";
            var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
            options.PreserveTrivia = true;
            options.DocumentTransforms.Add(new MarkdownCompactHeadingBoundaryTransform());
            var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, options);
            var native = MarkdownNativeDocument.FromParseResult(result);
            var affectedSourceSpan = result.TransformDiagnostics[0].AffectedSourceSpan;
            Assert.True(affectedSourceSpan.HasValue);
            var edit = native.CreateReplaceEdit(affectedSourceSpan.Value, "Updated");

            var roundtrip = MarkdownRoundtripWriter.WriteWithSourceEdit(result, edit);

            Assert.False(roundtrip.IsLossless);
            var diagnostic = Assert.Single(roundtrip.Diagnostics);
            Assert.Equal("roundtrip.document-transformed", diagnostic.Id);
            Assert.Equal(affectedSourceSpan, diagnostic.SourceSpan);
            Assert.Equal("Updated", roundtrip.Markdown);
        }

        [Fact]
        public void MarkdownRoundtripWriter_Reports_Transform_RelatedSourceSpans_When_SourceEdit_Falls_Back() {
            var options = MarkdownReaderOptions.CreateOfficeIMOProfile();
            options.PreserveTrivia = true;
            options.DocumentTransforms.Add(new UppercaseParagraphsTransform());
            var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics("""
alpha

beta
""", options);
            var native = MarkdownNativeDocument.FromParseResult(result);
            var edit = native.CreateReplaceEdit(Assert.IsType<MarkdownNativeParagraphBlock>(native.Blocks[0]), "Updated");

            var roundtrip = MarkdownRoundtripWriter.WriteWithSourceEdit(result, edit);

            Assert.False(roundtrip.IsLossless);
            var diagnostic = Assert.Single(roundtrip.Diagnostics);
            Assert.Equal("roundtrip.document-transformed", diagnostic.Id);
            Assert.Equal(new MarkdownSourceSpan(1, 1, 3, 4), diagnostic.SourceSpan);
            Assert.Equal(new[] { 1, 3 }, diagnostic.RelatedSourceSpans.Select(span => span.StartLine).ToArray());
            Assert.Equal(new[] { 1, 3 }, diagnostic.RelatedSourceSpans.Select(span => span.EndLine).ToArray());
            Assert.Equal("Updated", roundtrip.Markdown);
        }

        [Fact]
        public void MarkdownRoundtripWriter_Reports_Fallback_When_SourceEdit_Cannot_Map_To_OriginalMarkdown() {
            const string markdown = "# Ol\u200Bd\r\n";
            var options = new MarkdownReaderOptions {
                PreserveTrivia = true,
                InputNormalization = new MarkdownInputNormalizationOptions {
                    NormalizeZeroWidthSpacingArtifacts = true
                }
            };
            var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, options);
            var native = MarkdownNativeDocument.FromParseResult(result);
            var heading = Assert.IsType<MarkdownNativeHeadingBlock>(native.Blocks[0]);
            var edit = native.CreateReplaceEdit(heading.TextSourceSpan!.Value, "New");

            var roundtrip = MarkdownRoundtripWriter.WriteWithSourceEdit(result, edit);

            Assert.False(roundtrip.IsLossless);
            var diagnostic = Assert.Single(roundtrip.Diagnostics);
            Assert.Equal("roundtrip.original-source-slice-unavailable", diagnostic.Id);
            Assert.Equal(edit.SourceSpan, diagnostic.SourceSpan);
            Assert.Equal("# New\n", roundtrip.Markdown);
        }

        [Fact]
        public void MarkdownRoundtripWriter_Reports_Fallback_For_Overlapping_SourceEdits() {
            const string markdown = "# Old **Title**\n";
            var options = new MarkdownReaderOptions {
                PreserveTrivia = true
            };
            var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, options);
            var native = MarkdownNativeDocument.FromParseResult(result);
            var heading = Assert.IsType<MarkdownNativeHeadingBlock>(native.Blocks[0]);
            var edit = native.CreateReplaceEdit(heading.TextSourceSpan!.Value, "New Title");

            var roundtrip = MarkdownRoundtripWriter.WriteWithSourceEdits(result, new[] { edit, edit });

            Assert.False(roundtrip.IsLossless);
            var diagnostic = Assert.Single(roundtrip.Diagnostics);
            Assert.Equal("roundtrip.overlapping-edits", diagnostic.Id);
            Assert.Equal(edit.SourceSpan, diagnostic.SourceSpan);
            Assert.Equal(markdown, roundtrip.Markdown);
        }

        [Fact]
        public void MarkdownRoundtripWriter_Applies_Nested_Inline_SourceEdit_To_OriginalMarkdown() {
            const string markdown = "Start **bold and _em_** end\r\n";
            var options = new MarkdownReaderOptions {
                PreserveTrivia = true
            };
            var result = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, options);
            var native = MarkdownNativeDocument.FromParseResult(result);
            var strong = Assert.Single(native.EnumerateInlines(), inline => inline.Kind == MarkdownNativeInlineKind.Strong);
            var emphasis = Assert.Single(strong.Children, inline => inline.Kind == MarkdownNativeInlineKind.Emphasis);
            Assert.Equal(new MarkdownSourceSpan(1, 19, 1, 20), emphasis.SourceSpan);
            var edit = native.CreateReplaceEdit(emphasis, "updated");

            var roundtrip = MarkdownRoundtripWriter.WriteWithSourceEdit(result, edit);

            Assert.True(roundtrip.IsLossless);
            Assert.Empty(roundtrip.Diagnostics);
            Assert.Equal("Start **bold and _updated_** end\r\n", roundtrip.Markdown);
        }

        [Fact]
        public void Reader_Enumerates_Headings_And_Resolved_Anchors() {
            const string markdown = """
# Title

## Repeat

## Repeat
""";

            var parsed = MarkdownReader.Parse(markdown);
            var headings = parsed.DescendantHeadings().ToArray();

            Assert.Equal(new[] { "Title", "Repeat", "Repeat" }, headings.Select(h => h.Text).ToArray());
            Assert.Equal("title", parsed.GetHeadingAnchor(headings[0]));
            Assert.Equal("repeat", parsed.GetHeadingAnchor(headings[1]));
            Assert.Equal("repeat-1", parsed.GetHeadingAnchor(headings[2]));

            var infos = parsed.GetHeadingInfos();
            Assert.Equal(new[] { "Title", "Repeat", "Repeat" }, infos.Select(info => info.Text).ToArray());
            Assert.Equal(new[] { "title", "repeat", "repeat-1" }, infos.Select(info => info.Anchor).ToArray());
            Assert.Equal(new[] { 1, 2, 2 }, infos.Select(info => info.Level).ToArray());
            Assert.Same(headings[1], infos[1].Block);

            var byAnchor = parsed.FindHeadingByAnchor("#repeat-1");
            Assert.NotNull(byAnchor);
            Assert.Equal("Repeat", byAnchor!.Text);
            Assert.Equal("repeat-1", byAnchor.Anchor);

            Assert.True(parsed.HasHeadingAnchor("#repeat"));
            Assert.True(parsed.HasHeadingAnchor("repeat-1"));
            Assert.Null(parsed.FindHeadingByAnchor("missing"));
            Assert.False(parsed.HasHeadingAnchor("missing"));

            var firstByText = parsed.FindHeading("repeat");
            Assert.NotNull(firstByText);
            Assert.Equal("repeat", firstByText!.Anchor);

            Assert.True(parsed.HasHeading("repeat"));
            Assert.True(parsed.HasHeading("Repeat", StringComparison.Ordinal));

            var caseSensitiveMiss = parsed.FindHeading("repeat", StringComparison.Ordinal);
            Assert.Null(caseSensitiveMiss);
            Assert.False(parsed.HasHeading("repeat", StringComparison.Ordinal));

            Assert.Null(parsed.FindHeading(string.Empty));
            Assert.False(parsed.HasHeading(string.Empty));

            var byText = parsed.FindHeadings("repeat");
            Assert.Equal(2, byText.Count);
            Assert.Equal(new[] { "repeat", "repeat-1" }, byText.Select(info => info.Anchor).ToArray());
        }

        private sealed class NoOpTransform : IMarkdownDocumentTransform {
            public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) => document;
        }

        private sealed class UppercaseParagraphsTransform : IMarkdownDocumentTransform {
            public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
                var transformed = MarkdownDoc.Create();
                foreach (var block in document.Blocks) {
                    if (block is ParagraphBlock paragraph) {
                        transformed.Add(new ParagraphBlock(new InlineSequence().Text(paragraph.Inlines.RenderMarkdown().ToUpperInvariant())));
                    } else {
                        transformed.Add(block);
                    }
                }

                return transformed;
            }
        }
    }
}
