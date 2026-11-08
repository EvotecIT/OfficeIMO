using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void LegacyDoc_LoadReportsProhibitedStyleHighlightAndRetainsSupportedFormatting() {
        byte[] bytes = LegacyDocTestBuilder.CreateUnicodeDocWithStyleRunFormattingAndDirectExplicitOffRunFormatting(
            includeProhibitedStyleHighlight: true);

        using WordDocument document = WordDocument.Load(new MemoryStream(bytes), options: null, out OfficeConversionReport report);

        OfficeDiagnostic warning = Assert.Single(report.Diagnostics, diagnostic => diagnostic.Code == "doc-stylesheet-invalid");
        Assert.Equal(OfficeDiagnosticSeverity.Warning, warning.Severity);
        Assert.Equal(OfficeConversionLossKind.Approximation, warning.LossKind);
        Assert.Contains("highlight", warning.Message);
        WordParagraph run = Assert.Single(document.Paragraphs);
        Assert.Equal("direct off", run.Text);
        Assert.Equal(WordHighlightColor.None, run.Highlight);
        Style style = Assert.Single(document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>(),
            candidate => candidate.StyleId == run.StyleId);
        Assert.NotNull(style.StyleRunProperties!.Bold);
        Assert.Equal(UnderlineValues.Single, style.StyleRunProperties.GetFirstChild<Underline>()?.Val?.Value);
        Assert.Null(style.StyleRunProperties.GetFirstChild<Highlight>());
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void LegacyDoc_SaveRejectsStyleHighlightBeforeWritingDestination(bool customStyle, bool explicitNone) {
        using WordDocument document = WordDocument.Create();
        WordParagraph run = document.AddParagraph("Direct highlight remains supported");
        run.Highlight = WordHighlightColor.Yellow;
        Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        Style style;
        if (customStyle) {
            style = new Style(new StyleName { Val = "Highlighted style" }, new BasedOn { Val = "Normal" },
                new StyleRunProperties(new Bold())) {
                Type = StyleValues.Paragraph, StyleId = "HighlightedStyle", CustomStyle = true
            };
            styles.Append(style);
            run._paragraph.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "HighlightedStyle" });
        } else {
            style = Assert.Single(styles.Elements<Style>(), candidate => candidate.StyleId == "Normal");
        }

        using (WordDocument valid = WordDocument.Load(new MemoryStream(document.ToBytes(new WordSaveOptions { Format = WordFileFormat.Doc })))) {
            Assert.Equal(WordHighlightColor.Yellow, Assert.Single(valid.Paragraphs).Highlight);
        }
        style.StyleRunProperties ??= new StyleRunProperties();
        style.StyleRunProperties.Append(new Highlight { Val = explicitNone ? HighlightColorValues.None : HighlightColorValues.Yellow });
        byte[] original = { 17, 34, 51, 68 };
        using var destination = new MemoryStream();
        destination.Write(original, 0, original.Length);
        destination.Position = 2;

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
            document.Save(destination, new WordSaveOptions { Format = WordFileFormat.Doc }));

        Assert.Contains("highlight is not permitted in DOC styles", exception.Message);
        Assert.Equal(original, destination.ToArray());
        Assert.Equal(2, destination.Position);
    }
}
