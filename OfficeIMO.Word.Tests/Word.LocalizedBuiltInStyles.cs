using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Word saves built-in styles under their English built-in name but a translated identifier. These templates use
    /// the identifiers French, German and Polish Word write for the same built-in styles.
    /// </summary>
    public class WordLocalizedBuiltInStyles {
        public static TheoryData<string, string, string, string, string> Cultures => new() {
            { "Titre1", "Titre2", "Titre", "Citation", "TM1" },
            { "berschrift1", "berschrift2", "Titel", "Zitat", "Verzeichnis1" },
            { "Nagwek1", "Nagwek2", "Tytu", "Cytat", "Spistreci1" }
        };

        private static WordDocument LoadLocalizedTemplate(string heading1, string heading2, string title, string quote, string toc1) {
            using var stream = new MemoryStream();
            using (WordDocument template = WordDocument.Create(stream)) {
                foreach (string english in new[] { "Heading1", "Heading2", "Title", "Quote", "TOC1" }) {
                    if (template.Styles.Contains(english)) template.Styles.Remove(english);
                }
                template.Styles.AddParagraphStyle(heading1, "heading 1").Bold = true;
                template.Styles.AddParagraphStyle(heading2, "heading 2").Italic = true;
                template.Styles.AddParagraphStyle(title, "Title").FontSize = OfficeLength.Points(26);
                template.Styles.AddParagraphStyle(quote, "Quote").Italic = true;
                template.Styles.AddParagraphStyle(toc1, "toc 1");
                template.Save();
            }
            stream.Position = 0;
            return WordDocument.Load(stream);
        }

        private static int StylesNamed(WordDocument document, string name) =>
            document.Styles.Count(style => string.Equals(style.Name, name, StringComparison.OrdinalIgnoreCase));

        [Theory]
        [MemberData(nameof(Cultures))]
        public void LocalizedBuiltInStyles_ReadAsBuiltInAndAreReusedWhenApplied(string heading1, string heading2, string title, string quote, string toc1) {
            using WordDocument document = LoadLocalizedTemplate(heading1, heading2, title, quote, toc1);
            WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 2);
            document.AddParagraph("Report").SetStyleId(title);
            document.AddParagraph("Overview").SetStyleId(heading1);
            document.AddParagraph("Details").Style = WordParagraphStyles.Heading2;
            document.AddParagraph("Quoted").SetStyleId("Quote");

            Assert.Equal(WordParagraphStyles.Title, document.Paragraphs.Single(paragraph => paragraph.Text == "Report").Style);
            Assert.Equal(WordParagraphStyles.Heading1, document.Paragraphs.Single(paragraph => paragraph.Text == "Overview").Style);
            Assert.Equal(heading2, document.Paragraphs.Single(paragraph => paragraph.Text == "Details").StyleId);
            Assert.Equal(quote, document.Paragraphs.Single(paragraph => paragraph.Text == "Quoted").StyleId);
            Assert.Equal(heading1, document.Styles.GetOrAdd(WordParagraphStyles.Heading1).StyleId);

            WordTableOfContentRefreshReport report = toc.RefreshEntries();

            Assert.Equal(new[] { "Overview", "Details" }, report.Entries.Select(entry => entry.Text).ToArray());
            Assert.Equal(new[] { 1, 2 }, report.Entries.Select(entry => entry.Level).ToArray());
            Assert.Contains(toc.SdtBlock.Descendants<ParagraphStyleId>(), style => style.Val == toc1);
            foreach (string name in new[] { "heading 1", "heading 2", "Title", "Quote", "toc 1" }) {
                Assert.Equal(1, StylesNamed(document, name));
            }
            Assert.True(document.DocumentIsValid);
        }

        [Theory]
        [MemberData(nameof(Cultures))]
        public void LocalizedBuiltInStyles_ConvertToRealHeadingsAndQuotes(string heading1, string heading2, string title, string quote, string toc1) {
            using WordDocument document = LoadLocalizedTemplate(heading1, heading2, title, quote, toc1);
            document.AddParagraph("Overview").SetStyleId(heading1);
            document.AddParagraph("Body text");
            document.AddParagraph("Details").SetStyleId(heading2);
            document.AddParagraph("Quoted text").SetStyleId(quote);

            string markdown = document.ToMarkdown().ToMarkdown();
            Assert.Contains("# Overview", markdown, StringComparison.Ordinal);
            Assert.Contains("> Quoted text", markdown, StringComparison.Ordinal);
            Assert.Contains("## Details", markdown, StringComparison.Ordinal);

            string html = document.ToHtml();
            Assert.Contains("<h1", html, StringComparison.Ordinal);
            Assert.Contains("<h2", html, StringComparison.Ordinal);
            Assert.Contains("<blockquote", html, StringComparison.Ordinal);

            using var pdf = new MemoryStream();
            document.SaveAsPdf(pdf, new WordPdfExportOptions {
                IncludePageNumbers = false,
                PdfOptions = new PdfCreateOptions { TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers }
            });
            byte[] bytes = pdf.ToArray();
            var logical = PdfDocumentReadResult.Load(bytes);
            Assert.Equal(1, Assert.Single(logical.Headings, heading => heading.Text.Contains("Overview")).Level);
            Assert.Equal(2, Assert.Single(logical.Headings, heading => heading.Text.Contains("Details")).Level);
            Assert.Contains(PdfInspector.Inspect(bytes).Outlines, outline => outline.Title == "Overview");
        }

        [Fact]
        public void LocalizedNormalStyleKeepsTheDocumentDefaultBodyFontInHtml() {
            using WordDocument document = WordDocument.Create();
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            Style normal = styles.Elements<Style>().Single(style => style.StyleId == "Normal");
            normal.StyleId = "Standard";
            normal.StyleRunProperties?.RemoveAllChildren<RunFonts>();
            foreach (Style basedOnNormal in styles.Elements<Style>().Where(style => style.BasedOn?.Val == "Normal")) basedOnNormal.BasedOn!.Val = "Standard";
            RunPropertiesBaseStyle defaults = styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!;
            defaults.RemoveAllChildren<RunFonts>();
            defaults.PrependChild(new RunFonts { Ascii = "Garamond", HighAnsi = "Garamond" });
            document.InvalidateBuiltInStyleIndex();
            document.AddParagraph("Body").SetStyleId("Standard");

            Assert.Contains("--officeimo-font-body:", document.ToHtml(), StringComparison.Ordinal);
            Assert.Contains("Garamond", document.ToHtml(), StringComparison.Ordinal);
        }

        [Fact]
        public void BuiltInParagraphStylesIgnoreTableStylesWithTheBuiltInName() {
            using WordDocument document = WordDocument.Create();
            document.Styles.Remove("Heading1");
            document.Styles.AddTableStyle("HeadingTable", "heading 1");

            WordParagraph heading = document.AddParagraph("Overview");
            heading.Style = WordParagraphStyles.Heading1;

            Assert.Equal("Heading1", heading.StyleId);
            Assert.Equal(WordStyleType.Paragraph, document.Styles["Heading1"].Type);
        }

        [Fact]
        public void TableOfContentsRefreshKeepsAnEnglishStyleTheDocumentDefines() {
            using WordDocument document = WordDocument.Create();
            document.Styles.AddParagraphStyle("Citation", "Quote");
            document.Styles.AddParagraphStyle("Quote", "Pull quote");
            WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 1);
            document.AddParagraph("Overview").Style = WordParagraphStyles.Heading1;
            toc.RefreshEntries();
            ParagraphStyleId title = toc.SdtBlock.Descendants<ParagraphStyleId>().First();
            title.Val = "Quote";

            toc.RefreshEntries();

            Assert.Contains(toc.SdtBlock.Descendants<ParagraphStyleId>(), style => style.Val == "Quote");
            Assert.DoesNotContain(toc.SdtBlock.Descendants<ParagraphStyleId>(), style => style.Val == "Citation");
        }

        [Theory]
        [InlineData(0, 720)]
        [InlineData(360, 1080)]
        public void BlockquoteKeepingANonQuoteStyleStaysAQuoteThroughMarkdown(int noteIndentTwips, int expectedIndentTwips) {
            using WordDocument document = WordDocument.Create();
            WordStyle note = document.Styles.AddParagraphStyle("Note", "Note");
            if (noteIndentTwips > 0) note.IndentationBefore = OfficeLength.Points(noteIndentTwips / 20D);

            document.AddHtml("<blockquote class=\"Note\"><p>Class quote</p></blockquote><blockquote><p class=\"Note\">Own quote</p></blockquote>",
                new HtmlWordExportOptions { PreferDocumentStyles = true });

            foreach (string text in new[] { "Class quote", "Own quote" }) {
                WordParagraph quoted = document.Paragraphs.Single(paragraph => paragraph.Text == text);
                Assert.Equal("Note", quoted.StyleId);
                Assert.Equal(expectedIndentTwips, checked((int?)quoted.IndentationBefore?.ToTwips()));
            }
            string markdown = document.ToMarkdown().ToMarkdown();
            Assert.Contains("> Class quote", markdown, StringComparison.Ordinal);
            Assert.Contains("> Own quote", markdown, StringComparison.Ordinal);
            Assert.DoesNotContain("> > ", markdown, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void HtmlQuoteIndentStartsFromTheDefaultParagraphStyle(bool preferDocumentStyles) {
            using WordDocument document = WordDocument.Create();
            document.Styles.AddParagraphStyle("CorpsTexte", "Corps de texte").IndentationBefore = OfficeLength.Points(18);

            document.AddHtml("<p>Body</p><blockquote><p>Quoted</p></blockquote>",
                new HtmlWordExportOptions { PreferDocumentStyles = preferDocumentStyles, DefaultParagraphStyleId = "CorpsTexte" });

            WordParagraph body = document.Paragraphs.Single(paragraph => paragraph.Text == "Body");
            WordParagraph quoted = document.Paragraphs.Single(paragraph => paragraph.Text == "Quoted");
            Assert.Equal("CorpsTexte", body.StyleId);
            Assert.Null(body.IndentationBefore);
            Assert.Equal("CorpsTexte", quoted.StyleId);
            // The body style's 360 twips stay under the quotation's 720.
            Assert.Equal(1080, checked((int?)quoted.IndentationBefore?.ToTwips()));
            string markdown = document.ToMarkdown().ToMarkdown();
            Assert.Contains("> Quoted", markdown, StringComparison.Ordinal);
            Assert.DoesNotContain("> Body", markdown, StringComparison.Ordinal);
        }

        [Fact]
        public void HtmlQuoteKeepsAClassStyleThatReplacedTheDefaultStyle() {
            using WordDocument document = WordDocument.Create();
            document.Styles.AddParagraphStyle("CorpsTexte", "Corps de texte");
            document.Styles.AddParagraphStyle("Citation", "Quote");
            document.Styles.AddParagraphStyle("Note", "Note").IndentationBefore = OfficeLength.Points(18);

            document.AddHtml("<blockquote><p class=\"Note\">Noted</p><p>Plain</p></blockquote>",
                new HtmlWordExportOptions { PreferDocumentStyles = true, DefaultParagraphStyleId = "CorpsTexte" });

            WordParagraph noted = document.Paragraphs.Single(paragraph => paragraph.Text == "Noted");
            Assert.Equal("Note", noted.StyleId);
            Assert.Equal(1080, checked((int?)noted.IndentationBefore?.ToTwips()));
            Assert.Equal("Citation", document.Paragraphs.Single(paragraph => paragraph.Text == "Plain").StyleId);
            Assert.Contains("> Noted", document.ToMarkdown().ToMarkdown(), StringComparison.Ordinal);
        }

        [Theory]
        [MemberData(nameof(Cultures))]
        public void LocalizedBuiltInStyles_AreClassifiedWithoutLoadingTheStylesPart(string heading1, string heading2, string title, string quote, string toc1) {
            byte[] package;
            using (WordDocument template = LoadLocalizedTemplate(heading1, heading2, title, quote, toc1)) {
                template.AddParagraph("Overview").SetStyleId(heading1);
                template.AddParagraph("Quoted").SetStyleId(quote);
                using var output = new MemoryStream();
                template.Save(output);
                package = output.ToArray();
            }
            using var input = new MemoryStream(package, writable: false);
            using WordDocument document = WordDocument.Load(input, new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });

            Assert.Equal(WordParagraphStyles.Heading1, document.Paragraphs.Single(paragraph => paragraph.Text == "Overview").Style);
            Assert.Equal("Quote", document.Paragraphs.Single(paragraph => paragraph.Text == "Quoted").BuiltInStyleId);
            Assert.Equal(heading2, document.FindBuiltInStyleId("Heading2"));
            Assert.Null(document.FindBuiltInStyleId("Hyperlink"));

            Assert.False(document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.IsRootElementLoaded);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void BuiltInNameLookupsSkipEarlierStylesOfAnotherType(bool loadStyles) {
            byte[] package;
            using (var output = new MemoryStream()) {
                using (WordDocument source = WordDocument.Create(output)) {
                    source.Styles.Remove("Heading1");
                    source.Styles.Remove("Heading2");
                    // A malformed file: styles of another type share names, and identifiers, with the paragraph styles,
                    // and a paragraph style has no identifier.
                    Styles styles = source._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                    styles.Append(
                        new Style(new StyleName { Val = "heading 1" }) { Type = StyleValues.Table, StyleId = "HeadingTable" },
                        new Style(new StyleName { Val = "Body note" }) { Type = StyleValues.Character, StyleId = "BodyNoteRun" },
                        new Style(new StyleName { Val = "Body note" }) { Type = StyleValues.Paragraph },
                        new Style(new StyleName { Val = "heading 1" }) { Type = StyleValues.Paragraph, StyleId = "Titre1" },
                        new Style(new StyleName { Val = "Body note" }) { Type = StyleValues.Paragraph, StyleId = "BodyNote" },
                        new Style(new StyleName { Val = "heading 2" }) { Type = StyleValues.Table, StyleId = "Titre2" },
                        new Style(new StyleName { Val = "heading 2" }) { Type = StyleValues.Paragraph, StyleId = "Titre2" });
                    source.AddParagraph("Overview").SetStyleId("Titre1");
                    source.Save();
                }
                package = output.ToArray();
            }
            using var input = new MemoryStream(package, writable: false);
            using WordDocument document = WordDocument.Load(input, new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            if (loadStyles) _ = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles;

            Assert.Equal("Titre1", document.FindBuiltInStyleId("Heading1"));
            Assert.Equal(WordParagraphStyles.Heading1, document.Paragraphs.Single(paragraph => paragraph.Text == "Overview").Style);
            Assert.Equal("BodyNote", document.Styles.FindParagraphStyleId("Body note"));
            // Titre2 resolves to the earlier table style, so no paragraph style stands for heading 2.
            Assert.Null(document.FindBuiltInStyleId("Heading2"));
            Assert.Equal(loadStyles, document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.IsRootElementLoaded);
        }

        [Fact]
        public void TableOfContentsKeepsEntriesInTocPrefixedStylesWithoutBuiltInNames() {
            using var stream = new MemoryStream();
            using (WordDocument document = WordDocument.Create(stream)) {
                WordTableOfContent toc = document.AddTableOfContent(minLevel: 1, maxLevel: 1);
                document.AddParagraph("Overview").Style = WordParagraphStyles.Heading1;
                toc.RefreshEntries();
                document.Styles.AddParagraphStyle("TOC1_0", "toc 1_0");
                foreach (ParagraphStyleId entryStyle in toc.SdtBlock.Descendants<ParagraphStyleId>().Where(style => style.Val == "TOC1")) {
                    entryStyle.Val = "TOC1_0";
                }
                document.Save();
            }
            stream.Position = 0;
            using WordDocument reloaded = WordDocument.Load(stream);

            Assert.Contains(reloaded.TableOfContent!.SdtBlock.Descendants<ParagraphStyleId>(), style => style.Val == "TOC1_0");
        }

        [Fact]
        public void MarkdownHeadingSlugsSkipHeadingsInsideTableCells() {
            using WordDocument document = WordDocument.Create();
            WordTable table = document.AddTable(1, 1);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Overview";
            table.Rows[0].Cells[0].Paragraphs[0].Style = WordParagraphStyles.Heading1;
            WordParagraph heading = document.AddParagraph("Overview");
            heading.Style = WordParagraphStyles.Heading1;
            heading.AddBookmark("BodyOverview");
            document.AddParagraph("See ").AddHyperLink("overview", "BodyOverview");

            string markdown = document.ToMarkdown().ToMarkdown();

            Assert.Contains("(#overview)", markdown, StringComparison.Ordinal);
            Assert.DoesNotContain("(#overview-1)", markdown, StringComparison.Ordinal);
        }

        [Theory]
        [MemberData(nameof(Cultures))]
        public void LocalizedBuiltInStyles_AreReusedByHtmlAndMarkdownImports(string heading1, string heading2, string title, string quote, string toc1) {
            using WordDocument document = LoadLocalizedTemplate(heading1, heading2, title, quote, toc1);
            document.AddHtml("<h1>From HTML</h1><h2>HTML detail</h2><blockquote><p>HTML quote</p></blockquote>");
            var options = new MarkdownWordExportOptions { Template = document, PreferDocumentStyles = true };
            MarkdownDocument.Parse("# From Markdown\n\n## Markdown detail\n\n> Markdown quote").ToWord(options);

            string? StyleOf(string text) => document.BodyRoot.Elements<Paragraph>()
                .Single(paragraph => paragraph.InnerText == text).ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            Assert.Equal(heading1, StyleOf("From HTML"));
            Assert.Equal(heading2, StyleOf("HTML detail"));
            Assert.Equal(quote, StyleOf("HTML quote"));
            Assert.Equal(heading1, StyleOf("From Markdown"));
            Assert.Equal(heading2, StyleOf("Markdown detail"));
            Assert.Equal(quote, StyleOf("Markdown quote"));
            foreach (string name in new[] { "heading 1", "heading 2", "Quote" }) {
                Assert.Equal(1, StylesNamed(document, name));
            }
        }
    }
}
