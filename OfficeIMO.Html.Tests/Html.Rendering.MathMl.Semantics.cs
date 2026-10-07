using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("", "x divided by two")]
    [InlineData("display='block'", "x divided by two")]
    [InlineData("style='transform:translate(12px,8px)'", "x divided by two")]
    public void HtmlMathMlPdf_PreservesFormulaRoleDescriptionAndLogicalText(string attributes, string expectedDescription) {
        string html = "<html lang='en'><body><p>BEFORE <math " + attributes
            + " alttext='x divided by two'><mfrac><mi>x</mi><mn>2</mn></mfrac></math> AFTER</p>"
            + "<a href='https://example.test/math-link'>MATH-LINK</a></body></html>";
        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes();
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(PdfCore.PdfInspector.Inspect(bytes).TaggedContent);
        PdfCore.PdfStructureElementInfo formula = Assert.Single(tagged.StructureElements, element => element.StructureType == "Formula");
        Assert.Equal(expectedDescription, formula.AlternateText);
        Assert.NotEmpty(formula.ChildElementObjectNumbers);
        Assert.DoesNotContain(tagged.StructureElements, element => element.StructureType == "Figure");
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("BEFORE", text, StringComparison.Ordinal);
        Assert.Contains("AFTER", text, StringComparison.Ordinal);
        Assert.Equal(1, text.Split(new[] { "(x)/(2)" }, StringSplitOptions.None).Length - 1);
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(bytes));
    }

    [Theory]
    [InlineData("dir='rtl'", "שלום")]
    [InlineData("style='writing-mode:vertical-rl'", "BEFORE")]
    [InlineData("style='margin-top:140px'", "BEFORE")]
    public void HtmlMathMlPdf_KeepsFormulaDescriptionThroughReorderedAndPagedContainers(string attributes, string prefix) {
        string html = "<html lang='en'><body style='margin:0'><div " + attributes + ">" + prefix
            + " <math alttext='x squared'><msup><mi>x</mi><mn>2</mn></msup></math> AFTER</div></body></html>";
        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions {
            PageSize = new OfficeIMO.Drawing.OfficePageSize(3D, 1D),
            HonorCssPageRules = false, Margins = HtmlRenderMargins.All(0D)
        });
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(PdfCore.PdfInspector.Inspect(bytes).TaggedContent);
        PdfCore.PdfStructureElementInfo formula = Assert.Single(tagged.StructureElements, element => element.StructureType == "Formula");
        Assert.Equal("x squared", formula.AlternateText);
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains(prefix, text, StringComparison.Ordinal);
        Assert.Contains("AFTER", text, StringComparison.Ordinal);
        Assert.Equal(1, text.Split(new[] { "x^(2)" }, StringSplitOptions.None).Length - 1);
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(bytes));
    }

    [Fact]
    public void HtmlMathMlPdf_UsesLogicalDescriptionFallbackAndOmitsHiddenFormula() {
        const string html = "<html lang='en'><body><math><mi>x</mi></math>"
            + "<math style='display:none' alttext='hidden formula'><mi>y</mi></math></body></html>";
        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes();
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(PdfCore.PdfInspector.Inspect(bytes).TaggedContent);
        PdfCore.PdfStructureElementInfo formula = Assert.Single(tagged.StructureElements, element => element.StructureType == "Formula");
        Assert.Equal("x", formula.AlternateText);
    }
    [Theory]
    [InlineData("<math><mtext> </mtext></math>")]
    [InlineData("<math><mrow/></math>")]
    public void HtmlMathMlPdf_EmptyPlaceholderDoesNotRejectTheSurroundingDocument(string math) {
        byte[] bytes = HtmlConversionDocument.Parse("<p>BEFORE " + math + " AFTER</p>").ToPdfBytes();
        string text = PdfCore.PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("BEFORE", text, StringComparison.Ordinal);
        Assert.Contains("AFTER", text, StringComparison.Ordinal);
    }

}
