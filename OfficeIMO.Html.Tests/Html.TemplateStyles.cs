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
        public void AddHtml_ClassMappedBlockquoteKeepsItsAuthoredMargin() {
            using WordDocument document = WordDocument.Create();
            var options = new HtmlWordExportOptions();
            options.ClassStyles["pull"] = WordParagraphStyles.ListParagraph;

            document.AddHtml("<blockquote class=\"pull\" style=\"margin-left: 36pt\"><p>Pulled</p></blockquote><blockquote class=\"pull\"><p>Plain</p></blockquote>", options);

            WordParagraph pulled = document.Paragraphs.Single(paragraph => paragraph.Text == "Pulled");
            Assert.Equal("ListParagraph", pulled.StyleId);
            Assert.Equal(720, checked((int?)pulled.IndentationBefore?.ToTwips()));
            Assert.Null(document.Paragraphs.Single(paragraph => paragraph.Text == "Plain").IndentationBefore);
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
