using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using DocumentFormat.OpenXml.Packaging;
using System;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_TableCaptionAbove() {
            string html = "<table><caption>Above</caption><tr><td>A</td></tr></table>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions { TableCaptionPosition = TableCaptionPosition.Above });
            var bodyXml = doc._wordprocessingDocument.MainDocumentPart.Document.Body.OuterXml;
            Assert.True(bodyXml.IndexOf("Above", StringComparison.Ordinal) < bodyXml.IndexOf("<w:tbl", StringComparison.Ordinal));
        }

        [Fact]
        public void HtmlToWord_TableCaptionBelow() {
            string html = "<table><caption>Below</caption><tr><td>A</td></tr></table>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions { TableCaptionPosition = TableCaptionPosition.Below });
            var bodyXml = doc._wordprocessingDocument.MainDocumentPart.Document.Body.OuterXml;
            Assert.True(bodyXml.IndexOf("Below", StringComparison.Ordinal) > bodyXml.IndexOf("<w:tbl", StringComparison.Ordinal));
        }
    }
}
