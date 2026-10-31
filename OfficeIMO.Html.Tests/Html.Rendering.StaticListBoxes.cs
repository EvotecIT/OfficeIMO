using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false, 1, "Beta|Gamma|Delta", "14px/20px Georgia", 5)]
    [InlineData(true, 1, "Beta|Gamma|Delta", "14px/20px Georgia", 5)]
    [InlineData(false, 4, "Gamma|Delta|Epsilon", "14px/20px Georgia", 5)]
    [InlineData(false, 1, "Alpha|Beta|Gamma", "16px/1.4 Georgia", 3)]
    public void HtmlPdf_StaticListBoxesKeepVisibleOptionsAndHighlightSelection(bool multiple, int selected, string expected,
        string font, int count) {
        string[] labels = new[] { "Alpha", "Beta", "Gamma", "Delta", "Epsilon" }.Take(count).ToArray();
        string html = "<select name='choices' size='3' style='font:" + font + "' " + (multiple ? "multiple" : string.Empty) + ">"
            + string.Concat(labels.Select((label, index) => "<option" + (index == selected ? " selected" : string.Empty) + ">" + label + "</option>"))
            + "</select>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(source);
        HtmlRenderText selectedText = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(),
            text => text.Text == labels[selected]);
        Assert.Equal(OfficeColor.White, selectedText.Color);
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillColor == OfficeColor.FromRgb(0, 95, 184));

        using PdfCore.PdfDocument pdf = source.ToPdf(new HtmlToPdfOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        });
        byte[] bytes = pdf.ToBytes();
        Assert.Empty(PdfCore.PdfInspector.Inspect(bytes).FormFields);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(expected.Split('|'), independent.GetPage(1).GetWords().Select(word => word.Text));
    }

    [Fact]
    public void HtmlPdf_StaticTextareaRetainsEveryAuthoredRowWithFractionalLineHeight() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<textarea rows='3' style='font:16px/1.4 Georgia'>Alpha\nBeta\nGamma</textarea>");
        using PdfCore.PdfDocument pdf = source.ToPdf(new HtmlToPdfOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        });
        using var independent = UglyToad.PdfPig.PdfDocument.Open(pdf.ToBytes());

        Assert.Equal(new[] { "Alpha", "Beta", "Gamma" }, independent.GetPage(1).GetWords().Select(word => word.Text));
    }
}
