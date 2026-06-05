using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Html;
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void Test_Html_Ruby_ImportsAsWordRuby() {
        var html = "<p><ruby lang=\"ja-JP\"><rb>漢字</rb><rp>(</rp><rt>かんじ</rt><rp>)</rp></ruby></p>";

        using var document = html.LoadFromHtml(new HtmlToWordOptions { FontFamily = "Calibri" });

        var paragraph = Assert.Single(document._wordprocessingDocument.MainDocumentPart!.Document.Body!.Elements<Paragraph>());
        var ruby = Assert.Single(paragraph.Descendants<Ruby>());
        Assert.Equal("漢字", ruby.RubyBase!.InnerText);
        Assert.Equal("かんじ", ruby.RubyContent!.InnerText);
        Assert.Equal(RubyAlignValues.Center, ruby.RubyProperties!.RubyAlign!.Val!.Value);
    }

    [Fact]
    public void Test_Html_Ruby_FallbackWithoutAnnotationKeepsText() {
        var html = "<p><ruby>漢字</ruby></p>";

        using var document = html.LoadFromHtml(new HtmlToWordOptions());

        var paragraph = Assert.Single(document._wordprocessingDocument.MainDocumentPart!.Document.Body!.Elements<Paragraph>());
        Assert.Empty(paragraph.Elements<Ruby>());
        Assert.Equal("漢字", paragraph.InnerText);
    }

    [Fact]
    public void Test_Html_Ruby_DocxValidatesWithOpenXmlValidator() {
        var html = "<p>Word <ruby><rb>東</rb><rt>とう</rt></ruby> ruby</p>";
        var path = Path.Combine(Path.GetTempPath(), "officeimo-html-ruby-" + Guid.NewGuid().ToString("N") + ".docx");

        try {
            using (var document = html.LoadFromHtml(new HtmlToWordOptions())) {
                document.Save(path);
            }

            using var package = WordprocessingDocument.Open(path, false);
            var errors = new OpenXmlValidator().Validate(package).ToList();
            Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors.Select(error => $"{error.Description} Path={error.Path?.XPath} Node={error.Node?.OuterXml}")));
            Assert.Single(package.MainDocumentPart!.Document.Descendants<Ruby>());
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void Test_Html_Ruby_RoundTripsToHtml() {
        var html = "<p>Word <ruby><rb>東</rb><rt>とう</rt></ruby> ruby</p>";

        using var document = html.LoadFromHtml(new HtmlToWordOptions());

        string roundTrip = document.ToHtml();

        Assert.Contains("<ruby>", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<rb>東</rb>", roundTrip, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<rt>とう</rt>", roundTrip, StringComparison.OrdinalIgnoreCase);
    }
}
