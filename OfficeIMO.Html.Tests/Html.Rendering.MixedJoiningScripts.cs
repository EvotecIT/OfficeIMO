using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlImages_ParagraphBidiRetainsArabicJoiningBesideHebrew() {
        const string logical = "العربية نص עברית";
        const string arabic = "العربية نص";
        const string markup = "<style>@page{size:320px 120px;margin:0}body{margin:0}"
            + "p{font:20px 'OfficeIMO Shaping Test'}</style>"
            + "<p dir='rtl'><span>עברית </span><span style='color:red'>" + logical + "</span></p>";
        var options = new HtmlToPdfOptions {
            AllowSystemFontFallback = false,
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(
                (logical + OfficeArabicTextShaper.Shape(logical)).Select(character => (int)character).ToArray()));
        HtmlConversionDocument document = HtmlConversionDocument.Parse(markup);
        HtmlRenderResult retained = HtmlRenderEngine.Execute(document,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Png, options));
        string painted = string.Concat(EnumerateRenderVisuals(retained.Document.Pages[0].Scene)
            .OfType<HtmlRenderText>().Select(text => text.Text));

        Assert.Equal(OfficeArabicTextShaper.Shape(arabic).Where(IsArabicForm).OrderBy(character => character),
            painted.Where(IsArabicForm).OrderBy(character => character));
        Assert.DoesNotContain(painted, character => character >= '\u0620' && character <= '\u064A');
        Assert.DoesNotContain(retained.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported);
        byte[] pdf = document.ToPdfBytes(options);
        var read = OfficeIMO.Pdf.PdfReadDocument.Open(pdf);
        Assert.Contains(logical, read.ExtractText(), StringComparison.Ordinal);
        var expected = OfficeIMO.Pdf.PdfImageExportExtensions.ExportImages(read, OfficeImageExportFormat.Png,
            new OfficeIMO.Pdf.PdfImageExportOptions { Scale = 96D / 72D });
        Assert.Equal(Assert.Single(expected).Bytes, retained.ExportImage().Bytes);
    }

    private static bool IsArabicForm(char value) => value >= '\uFB50' && value <= '\uFDFF'
        || value >= '\uFE70' && value <= '\uFEFF';
}
