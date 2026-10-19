using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("raster", "")]
    [InlineData("raster", "opacity:.8")]
    [InlineData("raster-crop", "")]
    [InlineData("svg", "")]
    [InlineData("svg-crop", "")]
    [InlineData("text", "")]
    [InlineData("checkbox", "")]
    [InlineData("radio", "")]
    public void HtmlPdf_TaggedSourceOrderRetainsAtomicContentBeforeParagraphAndTable(string kind, string effect) {
        bool form = kind is "text" or "checkbox" or "radio";
        string atomic = form
            ? "<input type='" + kind + "' name='source' value='FieldValue'>"
            : SourceOrderImage(kind);
        string html = SourceOrderStyles(400)
            + "<section style='" + effect + "'><h1>Heading</h1>" + atomic
            + "<p>Summary</p>" + SourceOrderTable(2) + "<h2>End</h2></section>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(SourceOrderOptions());
        var root = SourceOrderRoot(pdf, out var elements);

        // Searchable field text has its own existing Span carrier. The widget's
        // Form still belongs at the same source position, before the paragraph.
        Assert.Equal(new[] { "H1", form ? "Form" : "Figure", "P", "Table", "H2" },
            root.ChildElementObjectNumbers.Select(id => elements[id].StructureType).Where(role => role != "Span"));
        Assert.Single(elements.Values, item => item.StructureType == (form ? "Form" : "Figure"));
        Assert.Single(elements.Values, item => item.StructureType == "Table");
        Assert.Equal(1, PdfCore.PdfInspector.Inspect(pdf).PageCount);
    }

    [Fact]
    public void HtmlPdf_TaggedSourceOrderSurvivesRepeatedTableFragments() {
        string html = SourceOrderStyles(160) + "<section><h1>Heading</h1><p>Before</p>"
            + SourceOrderImage("raster") + SourceOrderTable(24) + "<h2>After</h2></section>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(SourceOrderOptions());
        var root = SourceOrderRoot(pdf, out var elements);

        Assert.True(PdfCore.PdfInspector.Inspect(pdf).PageCount >= 4);
        Assert.Equal(new[] { "H1", "P", "Figure", "Table", "H2" },
            root.ChildElementObjectNumbers.Select(id => elements[id].StructureType));
        Assert.Single(elements.Values, item => item.StructureType == "Table");
        Assert.Single(elements.Values, item => item.StructureType == "TH");
        Assert.Equal(24, elements.Values.Count(item => item.StructureType == "TD"));
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
        foreach (string marker in new[] { "Heading", "Before", "Row24", "After" })
            Assert.Equal(1, text.Split(new[] { marker }, StringSplitOptions.None).Length - 1);
    }

    [Fact]
    public void HtmlPdf_TaggedSourceOrderKeepsFlattenedAndReorderedFlexScopesTogether() {
        string html = SourceOrderStyles(400)
            + "<section><h1>Heading</h1><div style='display:flex'>"
            + "<section style='display:contents'><p>First</p>" + SourceOrderTable(1) + "</section>"
            + "<section style='order:-1'><p>Second</p></section></div><p>After</p></section>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(SourceOrderOptions());
        var root = SourceOrderRoot(pdf, out var elements);

        Assert.Equal(new[] { "H1", "Div", "P" }, root.ChildElementObjectNumbers.Select(id => elements[id].StructureType));
        var flex = elements[root.ChildElementObjectNumbers[1]];
        Assert.Equal(new[] { "Sect", "Div" }, flex.ChildElementObjectNumbers.Select(id => elements[id].StructureType));
        var first = elements[flex.ChildElementObjectNumbers[0]];
        Assert.Equal(new[] { "Div", "Div" }, first.ChildElementObjectNumbers.Select(id => elements[id].StructureType));
        Assert.Equal(new[] { "P", "Table" }, first.ChildElementObjectNumbers
            .Select(id => elements[elements[id].ChildElementObjectNumbers.Single()].StructureType));
    }

    private static HtmlToPdfOptions SourceOrderOptions() => new() {
        AutoFitWidePrintContent = false,
        ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
    };

    private static string SourceOrderStyles(int height) =>
        "<!doctype html><html lang='en-US'><style>@page{size:400px " + height + "px;margin:12px}"
        + "body,p,h1,h2,td,th{margin:0;font:14px/20px Arial}img{display:block;width:24px;height:24px}"
        + "table{width:100%;border-collapse:collapse}td,th{padding:0;border:1px solid black}th{background:#dceffc}</style>";

    private static string SourceOrderTable(int rows) =>
        "<table><thead><tr><th scope='col'>Header</th></tr></thead><tbody>"
        + string.Concat(Enumerable.Range(1, rows).Select(index => "<tr><td>Row" + index.ToString("D2") + "</td></tr>"))
        + "</tbody></table>";

    private static string SourceOrderImage(string kind) {
        string source = kind.StartsWith("svg", StringComparison.Ordinal)
            ? "data:image/svg+xml;base64," + Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(
                "<svg xmlns='http://www.w3.org/2000/svg' width='48' height='24'><rect width='48' height='24' fill='blue'/></svg>"))
            : "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(4, 2));
        return "<img alt='Source icon' src='" + source + "'"
            + (kind.EndsWith("-crop", StringComparison.Ordinal) ? " style='object-fit:cover'" : string.Empty) + ">";
    }

    private static PdfCore.PdfStructureElementInfo SourceOrderRoot(byte[] pdf,
        out Dictionary<int, PdfCore.PdfStructureElementInfo> elements) {
        var tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(PdfCore.PdfInspector.Inspect(pdf).TaggedContent);
        elements = tagged.StructureElements.ToDictionary(item => item.ObjectNumber);
        var document = Assert.Single(tagged.StructureElements, item => item.StructureType == "Document");
        return elements[Assert.Single(document.ChildElementObjectNumbers)];
    }
}
