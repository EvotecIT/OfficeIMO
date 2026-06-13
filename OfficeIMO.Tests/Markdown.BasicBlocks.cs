using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void MarkdownToWord_ConvertsVariousElements() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string md = $@"# Heading 1

Paragraph with **bold** and *italic* and [link](https://example.com).

- Item 1
- Item 2

```c
code
```

|A|B|
|-|-|
|1|2|

> Quote line

---

![Alt]({imagePath})
";
            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions { FontFamily = "Calibri" });

            Assert.Equal(WordParagraphStyles.Heading1, doc.Paragraphs[0].Style);
            var quoteParagraph = doc.Paragraphs.First(p => p.Text.Contains("Quote line"));
            Assert.True(quoteParagraph.IndentationBefore > 0);

            var codeParagraph = doc.Paragraphs.First(p => p.Text.Contains("code"));
            // New Markdown engine does not assign a language-specific style to code blocks.
            // It uses monospace font on runs instead of paragraph style.
            Assert.Null(codeParagraph.StyleId);

            using MemoryStream ms = new();
            doc.Save(ms);
            ms.Position = 0;
            using WordprocessingDocument docx = WordprocessingDocument.Open(ms, false);
            var body = docx.MainDocumentPart!.Document.Body!;

            var codeRun = body.Descendants<Run>().First(r => r.InnerText.Contains("code"));
            Assert.Equal(FontResolver.Resolve("monospace"), codeRun.RunProperties!.RunFonts!.Ascii);
        }

        [Fact]
        public void Markdown_BlockQuote_Nesting_RoundTrip() {
            string md = @"> Level 1\n> > Level 2";
            var doc = md.LoadFromMarkdown();

            string markdown = doc.ToMarkdown();
            Assert.Contains("> Level 1", markdown);
            Assert.Contains("> > Level 2", markdown);
        }

        [Fact]
        public void MarkdownToWord_Renders_DetailsBlock_As_Structured_Paragraphs() {
            const string md = """
                <details open>
                <summary>More info</summary>

                Hidden text
                </details>
                """;

            using var doc = md.LoadFromMarkdown(new MarkdownToWordOptions());

            var summaryIndex = doc.Paragraphs.FindIndex(p => string.Equals(p.Text.Trim(), "More info", StringComparison.Ordinal));
            var bodyIndex = doc.Paragraphs.FindIndex(p => string.Equals(p.Text.Trim(), "Hidden text", StringComparison.Ordinal));
            Assert.True(summaryIndex >= 0);
            Assert.True(bodyIndex >= 0);

            var summaryParagraph = doc.Paragraphs[summaryIndex];
            Assert.Contains(summaryParagraph.GetRuns(), run => run.Bold);
            Assert.True(bodyIndex > summaryIndex);
        }

        [Fact]
        public void MarkdownToWord_Renders_FrontMatter_Header_Before_Body() {
            const string md = """
                ---
                title: Sample
                tags: [docs, ast]
                ---

                Body text
                """;

            using var doc = md.LoadFromMarkdown(new MarkdownToWordOptions());

            var paragraphs = doc.Paragraphs
                .Select(p => p.Text.Trim())
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();

            Assert.True(paragraphs.Count >= 5);
            Assert.Equal("---", paragraphs[0]);
            Assert.Equal("title: Sample", paragraphs[1]);
            Assert.Equal("tags: [docs, ast]", paragraphs[2]);
            Assert.Equal("---", paragraphs[3]);
            Assert.Contains("Body text", paragraphs);
        }

        [Fact]
        public void MarkdownToWord_Renders_Toc_Marker_As_Native_Word_TableOfContents() {
            const string markdown = """
                [TOC min=2 max=4 title="Contents"]

                # Report

                ## Region

                ### Pipeline
                """;

            using var document = markdown.LoadFromMarkdown(new MarkdownToWordOptions());

            Assert.NotNull(document.TableOfContent);
            Assert.Equal(2, document.TableOfContent!.MinLevel);
            Assert.Equal(4, document.TableOfContent.MaxLevel);
            Assert.Equal("Contents", document.TableOfContent.Text);
            Assert.True(document.Settings.UpdateFieldsOnOpen);
        }

        [Fact]
        public void MarkdownToWord_Renders_Titleless_Toc_Marker_Without_Default_Word_Title() {
            const string markdown = """
                [TOC min=2 max=4]

                # Report

                ## Region
                """;

            using var document = markdown.LoadFromMarkdown(new MarkdownToWordOptions());

            Assert.NotNull(document.TableOfContent);
            Assert.Equal(2, document.TableOfContent!.MinLevel);
            Assert.Equal(4, document.TableOfContent.MaxLevel);
            Assert.Equal(string.Empty, document.TableOfContent.Text);
        }

        [Fact]
        public void TocMarkerBlock_Renders_Word_Toc_Levels_Above_Markdown_Heading_Six() {
            var marker = new OfficeIMO.Markdown.TocMarkerBlock {
                MinLevel = 7,
                MaxLevel = 9,
                IncludeTitle = true,
                Title = "Deep contents",
                TitleLevel = 9
            };

            string markdown = ((OfficeIMO.Markdown.IMarkdownBlock)marker).RenderMarkdown();

            Assert.Equal("[TOC min=7 max=9 title=\"Deep contents\" titleLevel=6]", markdown);
        }

        [Fact]
        public void MarkdownToWord_Preserves_Toc_Marker_Levels_Above_Markdown_Heading_Six() {
            const string markdown = """
                [TOC min=7 max=9]

                ####### Deep region
                """;

            using var document = markdown.LoadFromMarkdown(new MarkdownToWordOptions());

            Assert.NotNull(document.TableOfContent);
            Assert.Equal(7, document.TableOfContent!.MinLevel);
            Assert.Equal(9, document.TableOfContent.MaxLevel);
        }

        [Fact]
        public void TocMarkerBlock_RoundTrips_Escaped_Title_Attributes() {
            const string title = "A \"quoted\" \\ title";
            var marker = new OfficeIMO.Markdown.TocMarkerBlock {
                MinLevel = 2,
                MaxLevel = 4,
                IncludeTitle = true,
                Title = title
            };

            string markdown = ((OfficeIMO.Markdown.IMarkdownBlock)marker).RenderMarkdown();

            using var document = markdown.LoadFromMarkdown(new MarkdownToWordOptions());

            Assert.Contains("title=\"A \\\"quoted\\\" \\\\ title\"", markdown, StringComparison.Ordinal);
            Assert.NotNull(document.TableOfContent);
            Assert.Equal(title, document.TableOfContent!.Text);
        }

        [Fact]
        public void TocMarkerBlock_ToHtmlFragment_Generates_Toc_From_Typed_Ast() {
            var markdown = OfficeIMO.Markdown.MarkdownDoc.Create()
                .Add(new OfficeIMO.Markdown.TocMarkerBlock {
                    MinLevel = 2,
                    MaxLevel = 3,
                    IncludeTitle = true,
                    Title = "Contents"
                })
                .H1("Report")
                .H2("Region")
                .H3("Pipeline");

            string html = markdown.ToHtmlFragment(new OfficeIMO.Markdown.HtmlOptions {
                Style = OfficeIMO.Markdown.HtmlStyle.Plain,
                CssDelivery = OfficeIMO.Markdown.CssDelivery.None,
                BodyClass = null
            });

            Assert.Contains("<h2>Contents</h2>", html, StringComparison.Ordinal);
            Assert.Contains("<a href=\"#region\">Region</a>", html, StringComparison.Ordinal);
            Assert.Contains("<a href=\"#pipeline\">Pipeline</a>", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<a href=\"#report\">Report</a>", html, StringComparison.Ordinal);
        }

        [Fact]
        public void MarkdownToWord_Renders_TocBuilder_Title_Only_As_Native_Toc_Title() {
            var markdown = OfficeIMO.Markdown.MarkdownDoc.Create()
                .H1("Report")
                .Toc(options => {
                    options.IncludeTitle = true;
                    options.Title = "Contents";
                    options.TitleLevel = 2;
                    options.MinLevel = 2;
                    options.MaxLevel = 4;
                }, placeAtTop: true)
                .H2("Region");

            using var document = markdown.ToWordDocument();

            var body = document._wordprocessingDocument.MainDocumentPart!.Document.Body!;
            int topLevelTitleParagraphs = body
                .Elements<Paragraph>()
                .Count(paragraph => string.Equals(paragraph.InnerText.Trim(), "Contents", StringComparison.Ordinal));

            Assert.NotNull(document.TableOfContent);
            Assert.Equal("Contents", document.TableOfContent!.Text);
            Assert.Equal(0, topLevelTitleParagraphs);
        }

        [Fact]
        public void MarkdownToWord_Renders_DocumentToc_RequireTopLevel_As_Native_Toc_From_LevelOne() {
            var markdown = OfficeIMO.Markdown.MarkdownDoc.Create()
                .Toc(options => {
                    options.IncludeTitle = false;
                    options.MinLevel = 2;
                    options.MaxLevel = 4;
                }, placeAtTop: true)
                .H1("Report")
                .H2("Region");

            using var document = markdown.ToWordDocument();

            Assert.NotNull(document.TableOfContent);
            Assert.Equal(1, document.TableOfContent!.MinLevel);
            Assert.Equal(4, document.TableOfContent.MaxLevel);
        }

        [Fact]
        public void MarkdownToWord_Renders_Typed_TocMarkerBlock_As_Native_Word_TableOfContents() {
            var markdown = OfficeIMO.Markdown.MarkdownDoc.Create()
                .Add(new OfficeIMO.Markdown.TocMarkerBlock {
                    MinLevel = 2,
                    MaxLevel = 4,
                    IncludeTitle = true,
                    Title = "Contents"
                })
                .H1("Report")
                .H2("Region");

            using var document = markdown.ToWordDocument();

            Assert.NotNull(document.TableOfContent);
            Assert.Equal(2, document.TableOfContent!.MinLevel);
            Assert.Equal(4, document.TableOfContent.MaxLevel);
            Assert.Equal("Contents", document.TableOfContent.Text);
        }

        [Fact]
        public void MarkdownToWord_Renders_Scoped_Toc_From_Realized_Entries() {
            var markdown = OfficeIMO.Markdown.MarkdownDoc.Create()
                .H1("Section")
                .TocHere(options => {
                    options.IncludeTitle = true;
                    options.Title = "Contents";
                    options.Scope = OfficeIMO.Markdown.TocScope.PreviousHeading;
                    options.MinLevel = 2;
                    options.MaxLevel = 3;
                })
                .H2("Inside")
                .H1("Other")
                .H2("Outside");

            using var document = markdown.ToWordDocument();
            string text = string.Join("\n", document.Paragraphs.Select(paragraph => paragraph.Text));

            Assert.Null(document.TableOfContent);
            Assert.Contains("Contents", text, StringComparison.Ordinal);
            Assert.Contains("[Inside](#inside)", text, StringComparison.Ordinal);
            Assert.DoesNotContain("[Outside](#outside)", text, StringComparison.Ordinal);
        }

        [Fact]
        public void MarkdownToWord_Renders_EmptyScopedToc_Without_NativeDocumentToc() {
            var markdown = OfficeIMO.Markdown.MarkdownDoc.Create()
                .H1("Empty Section")
                .TocHere(options => {
                    options.Scope = OfficeIMO.Markdown.TocScope.PreviousHeading;
                    options.MinLevel = 2;
                    options.MaxLevel = 3;
                })
                .H1("Other")
                .H2("Outside");

            using var document = markdown.ToWordDocument();
            string text = string.Join("\n", document.Paragraphs.Select(paragraph => paragraph.Text));

            Assert.Null(document.TableOfContent);
            Assert.DoesNotContain("[Outside](#outside)", text, StringComparison.Ordinal);
        }
    }
}
