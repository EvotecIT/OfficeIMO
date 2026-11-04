using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void AddHtmlToBody_InsertsAtBeginning() {
            using WordDocument document = WordDocument.Create();
            IReadOnlyList<WordElement> created = document.AddHtml("<p>First</p>");
            document.AddParagraph("Second");

            Assert.Equal("First", document.Paragraphs[0].Text);
            Assert.Equal("Second", document.Paragraphs[1].Text);
            Assert.Equal("First", Assert.Single(created.OfType<WordParagraph>(), paragraph => paragraph.Text.Length > 0).Text);
        }

        [Fact]
        public void AddHtmlToBody_InsertsInMiddle() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Start");
            IReadOnlyList<WordElement> created = document.AddHtml(OfficeIMO.Html.HtmlConversionDocument.Parse("<p>Middle</p>"));
            document.AddParagraph("End");

            Assert.Equal("Start", document.Paragraphs[0].Text);
            Assert.Equal("Middle", document.Paragraphs[1].Text);
            Assert.Equal("End", document.Paragraphs[2].Text);
            WordParagraph middle = Assert.Single(created.OfType<WordParagraph>(), paragraph => paragraph.Text.Length > 0);
            middle.Text = "Middle edited";
            Assert.Equal("Middle edited", document.Paragraphs[1].Text);
        }

        [Fact]
        public async Task AddHtmlToBody_InsertsAtEnd() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Start");
            document.AddParagraph("Middle");
            IReadOnlyList<WordElement> initial = await document.AddHtmlAsync(OfficeIMO.Html.HtmlConversionDocument.Parse("<p>End</p>"));

            Assert.Equal("Start", document.Paragraphs[0].Text);
            Assert.Equal("Middle", document.Paragraphs[1].Text);
            Assert.Equal("End", document.Paragraphs[2].Text);
            Assert.Single(initial.OfType<WordParagraph>(), paragraph => paragraph.Text == "End").Text = "End edited";

            IReadOnlyList<WordElement> created = await document.AddHtmlAsync("""
                <section><p>Before</p><table><tr><td>Table value</td></tr></table>
                <p><a href="https://example.test/">Linked</a></p></section>
                <section><ul><li>List<table><tr><td>Wrapped value</td></tr></table></li></ul>
                <p>After</p></section>
                """);
            Assert.Equal(3, document.Sections.Count);
            Assert.Equal(new[] { "Before", "[table]", "Linked", "List", "[sdt]", "After" },
                created.Select(element => element switch {
                    WordParagraph paragraph => paragraph.Text,
                    WordTable => "[table]",
                    WordStructuredDocumentTag => "[sdt]",
                    _ => throw new InvalidOperationException("Unexpected imported element.")
                }).Where(text => text.Length > 0));
            Assert.True(Assert.IsAssignableFrom<ICollection<WordElement>>(created).IsReadOnly);

            WordParagraph before = Assert.Single(created.OfType<WordParagraph>(), paragraph => paragraph.Text == "Before");
            WordParagraph linked = Assert.Single(created.OfType<WordParagraph>(), paragraph => paragraph.Text == "Linked");
            WordParagraph after = Assert.Single(created.OfType<WordParagraph>(), paragraph => paragraph.Text == "After");
            WordTable table = Assert.Single(created.OfType<WordTable>());
            WordStructuredDocumentTag control = Assert.Single(created.OfType<WordStructuredDocumentTag>());
            Assert.Same(document.Sections[1], before.Parent);
            Assert.Same(document.Sections[2], after.Parent);
            Assert.NotNull(linked._hyperlink);
            Assert.Same(document.OpenXmlDocument.MainDocumentPart!.Document.Body, table._table.Parent);
            Assert.Same(document.OpenXmlDocument.MainDocumentPart!.Document.Body, control.SdtBlock!.Parent);
            before.Text = "Before edited";
            linked.Text = "Linked edited";
            after.Text = "After edited";
            table.Rows[0].Cells[0].Paragraphs.Single(paragraph => paragraph.Text == "Table value").Text = "Table edited";
            control.Tag = "created-control-edited";

            using var stream = new MemoryStream();
            document.Save(stream);
            using WordDocument reloaded = WordDocument.Load(new MemoryStream(stream.ToArray()));
            Assert.Equal(3, reloaded.Sections.Count);
            Assert.Equal(new[] { "Start", "Middle", "End edited" },
                reloaded.Sections[0].Paragraphs.Where(paragraph => paragraph.Text.Length > 0).Select(paragraph => paragraph.Text));
            Assert.Contains(reloaded.Sections[1].Paragraphs, paragraph => paragraph.Text == "Before edited");
            Assert.Contains(reloaded.Sections[2].Paragraphs, paragraph => paragraph.Text == "After edited");
            Assert.Contains(reloaded.Tables.SelectMany(importedTable => importedTable.Paragraphs), paragraph => paragraph.Text == "Table edited");
            Hyperlink savedLink = Assert.Single(reloaded.OpenXmlDocument.MainDocumentPart!.Document.Body!.Descendants<Hyperlink>());
            Assert.Equal("Linked edited", savedLink.InnerText);
            Assert.Contains(reloaded.OpenXmlDocument.MainDocumentPart.HyperlinkRelationships,
                relationship => relationship.Id == savedLink.Id?.Value && relationship.Uri.AbsoluteUri == "https://example.test/");
            Assert.Contains(reloaded.OpenXmlDocument.MainDocumentPart.Document.Body.Descendants<SdtBlock>(),
                block => block.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value == "created-control-edited");
        }
    }
}

