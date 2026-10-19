using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
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

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged)]
    public void HtmlPdf_TaggedSourceOrderSurvivesFormProjection(HtmlRenderIntentProfile profile) {
        string html = SourceOrderStyles(400)
            + "<section><h1>Heading</h1><input name='source' value='FieldValue'>"
            + "<p>Summary</p>" + SourceOrderTable(2) + "</section>";
        byte[] pdf = HtmlConversionDocument.Parse(html).RenderToPdfResult(
            HtmlRenderRequest.Create(profile, HtmlRenderEncoder.Pdf, SourceOrderOptions())).ToBytes();
        var root = SourceOrderRoot(pdf, out var elements);
        Assert.Equal(new[] { "H1", "Form", "P", "Table" }, root.ChildElementObjectNumbers
            .Select(id => elements[id].StructureType).Where(role => role != "Span"));
    }

    [Theory]
    [InlineData("Before")]
    [InlineData("<div>Before</div>")]
    [InlineData("Before<span class='rtl'>א</span>After")]
    [InlineData("<div style='writing-mode:vertical-rl;height:100px'>Before</div>")]
    public void HtmlPdf_TaggedSourceOrderRetainsAnonymousParagraphs(string before) {
        string html = SourceOrderStyles(400) + "<section><h1>Heading</h1>" + before
            + SourceOrderTable(2) + "<p>After</p></section>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(SourceOrderOptions());
        var root = SourceOrderRoot(pdf, out var elements);
        Assert.Equal(new[] { "H1", "P", "Table", "P" },
            root.ChildElementObjectNumbers.Select(id => elements[id].StructureType));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlPdf_TaggedSourceOrderRetainsAnonymousAndGeneratedFlexItems(bool generated) {
        string html = SourceOrderStyles(400)
            + (generated ? "<style>.items::before{content:'F'}.items::after{content:'Z'}</style>" : "")
            + "<section><div class='items' style='display:flex'>A<span>B</span>C</div></section>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(SourceOrderOptions());
        var root = SourceOrderRoot(pdf, out var elements);
        var order = new List<int>();
        Collect(root);
        Assert.Equal(order.OrderBy(value => value), order);
        Assert.Equal(generated ? 5 : 3, order.Count);

        void Collect(PdfCore.PdfStructureElementInfo item) {
            order.AddRange(item.MarkedContentReferences.Select(reference => reference.MarkedContentId));
            foreach (int child in item.ChildElementObjectNumbers) Collect(elements[child]);
        }
    }

    [Fact]
    public void HtmlPdf_TaggedSourceOrderKeepsDiagnosedEmptyAnonymousContent() {
        string html = SourceOrderStyles(400) + "<section><div>&#xf000;</div>" + SourceOrderTable(1) + "</section>";
        var result = HtmlConversionDocument.Parse(html).ToPdfDocumentResult(SourceOrderOptions());
        Assert.Contains(result.Report.Warnings,
            warning => warning.Code == HtmlPdfDiagnosticCodes.UnavailablePrivateUseGlyphOmitted);
        Assert.DoesNotContain("\uf000", PdfCore.PdfReadDocument.Open(result.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_TaggedSourceOrderRetainsPaginatedBidiEnvelopes() {
        string html = SourceOrderStyles(120) + "<section><h1>Heading</h1><div dir='rtl'>"
            + string.Join(" ", Enumerable.Repeat("Before <span class='rtl'>א</span> After", 16))
            + "</div>" + SourceOrderTable(2) + "<h2>End</h2></section>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(SourceOrderOptions());
        var root = SourceOrderRoot(pdf, out var elements);
        string?[] roles = root.ChildElementObjectNumbers.Select(id => elements[id].StructureType).ToArray();
        int table = Array.IndexOf(roles, "Table");
        Assert.True(PdfCore.PdfInspector.Inspect(pdf).PageCount > 1);
        Assert.Equal("H1", roles[0]);
        Assert.True(table > 1);
        Assert.All(roles.Skip(1).Take(table - 1), role => Assert.Equal("P", role));
        Assert.Equal(new[] { "Table", "H2" }, roles.Skip(table));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlPdf_TaggedSourceOrderRetainsInlineImageBetweenText(bool bidi) {
        string middle = SourceOrderImage("raster").Replace("<img ", "<img style='display:inline' ");
        string html = SourceOrderStyles(400) + "<section><p>Before" + middle
            + (bidi ? "<span class='rtl'>א</span>" : "") + "After</p></section>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(SourceOrderOptions());
        var root = SourceOrderRoot(pdf, out var elements);
        var paragraph = elements[Assert.Single(root.ChildElementObjectNumbers)];
        Assert.Equal("P", paragraph.StructureType);
        string?[] roles = paragraph.ChildElementObjectNumbers.Select(id => elements[id].StructureType).ToArray();
        Assert.Equal("Span", roles[0]);
        Assert.Contains("Figure", roles);
        Assert.True(roles[1] == "Figure", string.Join(",", roles));
        Assert.All(roles.Skip(2), role => Assert.Equal("Span", role));
        Assert.True(roles.Length >= 3);
    }

    private static HtmlToPdfOptions SourceOrderOptions() {
        byte[] hebrew = ManagedTextShapingTestAssets.CreateFont(0x05D0);
        var options = new HtmlToPdfOptions {
            AutoFitWidePrintContent = false,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        };
        options.Fonts.Add("Source Hebrew", hebrew);
        options.PdfOptions.RegisterNamedFontFamily(new PdfCore.PdfEmbeddedFontFamily("Source Hebrew", hebrew));
        return options;
    }

    private static string SourceOrderStyles(int height) =>
        "<!doctype html><html lang='en-US'><style>@page{size:400px " + height + "px;margin:12px}"
        + "body,p,h1,h2,td,th{margin:0;font:14px/20px Arial}img{display:block;width:24px;height:24px}"
        + "table{width:100%;border-collapse:collapse}td,th{padding:0;border:1px solid black}th{background:#dceffc}"
        + ".rtl{font-family:'Source Hebrew'}</style>";

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
