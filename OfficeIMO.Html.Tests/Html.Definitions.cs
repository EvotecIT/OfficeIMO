using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public class HtmlDefinitions {
        [Fact]
        public void DfnIsItalicAndRoundsTrip() {
            const string html = "<p>A <dfn>term</dfn> appears.</p>";
            using var doc = html.ToWordDocument();
            var runs = doc.Paragraphs[0].GetRuns().ToList();
            Assert.Equal("HtmlDfn", runs[1].CharacterStyleId);
            Assert.True(runs[1].Italic);
            Assert.Equal("term", runs[1].Text);

            string roundTrip = doc.ToHtml();
            Assert.Contains("<dfn", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("</dfn>", roundTrip, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void DefinitionListImportsValidDocxAndRoundTripsAsDefinitionList() {
            const string html = "<dl><dt>Term</dt><dd>Definition</dd></dl>";
            using var doc = html.ToWordDocument();

            Assert.Equal("HtmlDefinitionTerm", doc.Paragraphs[0].StyleId);
            Assert.Equal("HtmlDefinitionDescription", doc.Paragraphs[1].StyleId);
            Assert.Equal(720, doc.Paragraphs[1].IndentationBefore);

            using MemoryStream packageStream = doc.ToStream();
            packageStream.Position = 0;
            using WordprocessingDocument package = WordprocessingDocument.Open(packageStream, false);
            var errors = new OpenXmlValidator().Validate(package).ToList();
            Assert.True(errors.Count == 0, OpenXmlValidationFormatting.FormatValidationErrors(errors));

            string roundTrip = doc.ToHtml();
            Assert.Contains("<dl><dt>Term</dt><dd>Definition</dd></dl>", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<blockquote>Definition</blockquote>", roundTrip, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void WordToHtmlDefinitionMarkersExportConsecutiveParagraphsAsDefinitionList() {
            using var doc = WordDocument.Create();
            doc.AddParagraph("Term").SetStyleId("HtmlDefinitionTerm");
            doc.AddParagraph("Definition").SetStyleId("HtmlDefinitionDescription");
            doc.AddParagraph("Next paragraph");

            string html = doc.ToHtml();

            Assert.Contains("<dl><dt>Term</dt><dd>Definition</dd></dl>", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<p>Next paragraph</p>", html, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void DefinitionListInTableCellRoundsTripAsDefinitionList() {
            const string html = "<table><tr><td><dl><dt>Metric</dt><dd>Value</dd></dl></td></tr></table>";
            using var doc = html.ToWordDocument();
            var cell = doc.Tables[0].Rows[0].Cells[0];

            Assert.Contains(cell.Paragraphs, paragraph => paragraph.Text == "Metric" && paragraph.StyleId == "HtmlDefinitionTerm");
            Assert.Contains(cell.Paragraphs, paragraph => paragraph.Text == "Value" && paragraph.StyleId == "HtmlDefinitionDescription");

            string roundTrip = doc.ToHtml();

            Assert.Contains("<dl><dt>Metric</dt><dd>Value</dd></dl>", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("<blockquote>Value</blockquote>", roundTrip, StringComparison.OrdinalIgnoreCase);
        }
    }
}

