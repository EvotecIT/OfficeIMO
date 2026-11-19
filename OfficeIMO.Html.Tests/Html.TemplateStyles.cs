using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlTemplateStyles {
        // A template as a French Word installation saves it: built-in heading and quote styles keep their
        // English names under translated identifiers, next to the template's own body and term styles.
        private static WordDocument LoadTemplate() {
            using var stream = new MemoryStream();
            using (WordDocument template = WordDocument.Create(stream)) {
                WordStyle body = template.Styles.AddParagraphStyle("CorpsTexte", "Corps de texte");
                body.FontFamily = "Garamond";
                body.FontSize = OfficeLength.Points(12);
                template.Styles.Remove("Heading1");
                WordStyle heading = template.Styles.AddParagraphStyle("Titre1", "heading 1");
                heading.FontFamily = "Garamond";
                heading.Bold = true;
                template.Styles.AddParagraphStyle("Citation", "Quote").Italic = true;
                template.Styles.AddParagraphStyle("TermeDefini", "Defined term").Bold = true;
                template.AddParagraph("Existing template text");
                template.Save();
            }
            stream.Position = 0;
            return WordDocument.Load(stream);
        }

        [Fact]
        public void AddHtml_WithDocumentStyles_UsesTemplateStylesWithoutDirectFormatting() {
            using WordDocument document = LoadTemplate();
            const string html = """
                <style>p { font-family: Georgia; line-height: 2; } .note { font-weight: bold; font-size: 48pt; color: #336699; }</style>
                <h1 style="font-family: 'Comic Sans MS'; color: red">Report</h1>
                <p class="note" style="font-size: 20px; margin-left: 30px; text-align: center">Body <strong>bold</strong> text</p>
                <div style="background: yellow">Loose text</div>
                <blockquote style="margin-left: 50px"><p>Quoted</p></blockquote>
                <dl><dt class="TermeDefini">Term</dt><dd>Description</dd></dl>
                <ul><li>Item</li></ul>
                """;
            var options = new HtmlWordExportOptions {
                PreferDocumentStyles = true,
                DefaultParagraphStyleId = "Corps de texte",
                IgnoreCssFormatting = true,
                FontFamily = "Arial"
            };

            document.AddHtml(html, options);
            List<Paragraph> paragraphs = document.BodyRoot.Elements<Paragraph>()
                .Where(paragraph => paragraph.InnerText.Trim().Length > 0 && paragraph.InnerText != "Existing template text").ToList();

            Assert.Equal("Titre1", StyleOf(paragraphs, "Report"));
            Assert.Equal("CorpsTexte", StyleOf(paragraphs, "Body bold text"));
            Assert.Equal("CorpsTexte", StyleOf(paragraphs, "Loose text"));
            Assert.Equal("Citation", StyleOf(paragraphs, "Quoted"));
            Assert.Equal("TermeDefini", StyleOf(paragraphs, "Term"));
            Assert.Equal("HtmlDefinitionDescription", StyleOf(paragraphs, "Description"));
            Assert.NotEqual("CorpsTexte", StyleOf(paragraphs, "Item"));
            Assert.False(document.Styles.Contains("Heading1"));
            Assert.Equal("Garamond", document.Styles["CorpsTexte"].FontFamily);

            foreach (Paragraph paragraph in paragraphs) {
                bool listItem = paragraph.ParagraphProperties?.NumberingProperties != null;
                bool definition = paragraph.InnerText == "Description";
                Assert.True(paragraph.ParagraphProperties == null || paragraph.ParagraphProperties.ChildElements.All(child =>
                        child is ParagraphStyleId ||
                        child is ParagraphMarkRunProperties { HasChildren: false } ||
                        listItem && child is NumberingProperties or Indentation ||
                        definition && child is Indentation),
                    $"'{paragraph.InnerText}' carries direct paragraph formatting: {paragraph.ParagraphProperties?.OuterXml}");
                foreach (Run run in paragraph.Descendants<Run>()) {
                    bool strong = run.InnerText == "bold";
                    Assert.True(run.RunProperties == null || run.RunProperties.ChildElements.All(child => strong && child is Bold or BoldComplexScript),
                        $"Run '{run.InnerText}' carries direct formatting: {run.RunProperties?.OuterXml}");
                }
            }
            AssertValid(document);
        }

        [Fact]
        public void AddHtml_ByDefault_LetsClassAndQuoteStylesOwnTheirLayout() {
            using WordDocument document = WordDocument.Create();
            document.Styles.AddParagraphStyle("Quote").IndentationBefore = OfficeLength.Points(48);
            var options = new HtmlWordExportOptions();
            options.ClassStyles["term"] = WordParagraphStyles.Heading4;
            options.ClassStyles["detail"] = WordParagraphStyles.ListParagraph;

            document.AddHtml("<blockquote><p>Quoted</p></blockquote><dl><dt class=\"term\">Term</dt><dd class=\"detail\">Detail</dd><dt>Plain</dt><dd>Plain detail</dd></dl>", options);

            WordParagraph quoted = document.Paragraphs.Single(paragraph => paragraph.Text == "Quoted");
            Assert.Equal("Quote", quoted.StyleId);
            Assert.Null(quoted.IndentationBefore);
            Assert.Equal("Heading4", document.Paragraphs.Single(paragraph => paragraph.Text == "Term").StyleId);
            WordParagraph detail = document.Paragraphs.Single(paragraph => paragraph.Text == "Detail");
            Assert.Equal("ListParagraph", detail.StyleId);
            Assert.Null(detail.IndentationBefore);
            Assert.Equal("HtmlDefinitionTerm", document.Paragraphs.Single(paragraph => paragraph.Text == "Plain").StyleId);
            string html = document.ToHtml();
            Assert.Contains("<dt>Term</dt>", html, StringComparison.Ordinal);
            Assert.Contains("<dd>Detail</dd>", html, StringComparison.Ordinal);
            Assert.Equal(720, checked((int?)document.Paragraphs.Single(paragraph => paragraph.Text == "Plain detail").IndentationBefore?.ToTwips()));
        }

        [Fact]
        public void AddHtml_KeepsCharacterStylesOutOfParagraphRoles() {
            using WordDocument document = WordDocument.Create();
            document.Styles.AddCharacterStyle("CodeRun", "Quote");
            if (!document.Styles.Contains("Hyperlink")) document.Styles.AddCharacterStyle("Hyperlink");

            Assert.Throws<InvalidOperationException>(() =>
                document.AddHtml("<p>Body</p>", new HtmlWordExportOptions { DefaultParagraphStyleId = "Hyperlink" }));
            document.AddHtml("<blockquote><p>Quoted</p></blockquote>");

            WordParagraph quoted = document.Paragraphs.Single(paragraph => paragraph.Text == "Quoted");
            Assert.Null(quoted.StyleId);
            Assert.Equal(720, checked((int?)quoted.IndentationBefore?.ToTwips()));
        }

        [Fact]
        public void AddHtml_ClassNamesSelectDocumentStylesExactlyAcrossHtmlWhitespace() {
            using WordDocument document = WordDocument.Create();
            document.Styles.AddParagraphStyle("Body", "Upper body");
            document.Styles.AddParagraphStyle("body", "Lower body");
            document.Styles.AddParagraphStyle("Note", "Margin note");
            document.Styles.AddParagraphStyle("HtmlDefinitionDescription", "Term look");

            document.AddHtml("<style>.Note { font-size: 24pt; font-weight: bold; }</style>" +
                "<p class=\"Body\">Upper</p><p class=\"body\">Lower</p><p class=\"aux\tNote\">Noted</p>" +
                "<dl><dt class=\"HtmlDefinitionDescription\">Term</dt><dd>Meaning</dd></dl>",
                new HtmlWordExportOptions { PreferDocumentStyles = true });

            Assert.Equal("Body", document.Paragraphs.Single(paragraph => paragraph.Text == "Upper").StyleId);
            Assert.Equal("body", document.Paragraphs.Single(paragraph => paragraph.Text == "Lower").StyleId);
            // The document style outranks the heading the stylesheet's bold 24pt rule would infer.
            Assert.Equal("Note", document.Paragraphs.Single(paragraph => paragraph.Text == "Noted").StyleId);
            // The imported role, not the style identifier it happens to share, decides the exported element.
            Assert.Contains("<dt>Term</dt>", document.ToHtml(), StringComparison.Ordinal);
        }

        [Fact]
        public void ToWord_IgnoringCssFormattingLeavesThePageColor() {
            using WordDocument document = OfficeIMO.Html.HtmlConversionDocument
                .Parse("<body style=\"background-color:#101010;color:white\"><p>Text</p></body>")
                .ToWord(new HtmlWordExportOptions { IgnoreCssFormatting = true });

            Assert.Null(document.Background.Color);
        }

        [Fact]
        public void AddHtml_AfterRawStyleEditsUsesTheCurrentStyles() {
            using WordDocument document = WordDocument.Create();
            document.Styles.AddParagraphStyle("Citation", "Quote");
            document.AddHtml("<blockquote><p>First</p></blockquote>");
            Assert.Equal("Citation", document.Paragraphs.Single(paragraph => paragraph.Text == "First").StyleId);

            document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>()
                .Single(style => style.StyleId == "Citation").Remove();
            document.AddHtml("<blockquote><p>Second</p></blockquote>");

            WordParagraph second = document.Paragraphs.Single(paragraph => paragraph.Text == "Second");
            Assert.Null(second.StyleId);
            Assert.Equal(720, checked((int?)second.IndentationBefore?.ToTwips()));
        }

        [Fact]
        public void AddHtml_ClassMappedBlockquoteKeepsItsAuthoredMargin() {
            using WordDocument document = WordDocument.Create();
            var options = new HtmlWordExportOptions();
            options.ClassStyles["pull"] = WordParagraphStyles.ListParagraph;

            document.AddHtml("<blockquote class=\"pull\" style=\"margin-left: 36pt\"><p>Pulled</p></blockquote><blockquote class=\"pull\"><p>Plain</p></blockquote>", options);

            WordParagraph pulled = document.Paragraphs.Single(paragraph => paragraph.Text == "Pulled");
            Assert.Equal("ListParagraph", pulled.StyleId);
            Assert.Equal(720, checked((int?)pulled.IndentationBefore?.ToTwips()));
            // The class style keeps its own 36pt indent and the quotation adds 36pt on top.
            Assert.Equal(1440, checked((int?)document.Paragraphs.Single(paragraph => paragraph.Text == "Plain").IndentationBefore?.ToTwips()));
        }

        [Fact]
        public void AddHtml_WithUnknownDefaultParagraphStyle_ThrowsBeforeChangingTheDocument() {
            using WordDocument document = LoadTemplate();
            int blocks = document.BodyRoot.ChildElements.Count;

            var exception = Assert.Throws<InvalidOperationException>(() =>
                document.AddHtml("<p>Body</p>", new HtmlWordExportOptions { DefaultParagraphStyleId = "MissingBody" }));

            Assert.Contains("MissingBody", exception.Message, StringComparison.Ordinal);
            Assert.Equal(blocks, document.BodyRoot.ChildElements.Count);
        }

        private static string? StyleOf(IEnumerable<Paragraph> paragraphs, string text) =>
            Assert.Single(paragraphs, paragraph => paragraph.InnerText.Trim() == text).ParagraphProperties?.ParagraphStyleId?.Val?.Value;

        private static void AssertValid(WordDocument document) {
            using var stream = new MemoryStream(document.ToBytes());
            using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
            var errors = new OpenXmlValidator().Validate(package).ToList();
            Assert.True(errors.Count == 0, OpenXmlValidationFormatting.FormatValidationErrors(errors));
        }
    }
}
