using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("")]
    [InlineData("opacity:.85")]
    [InlineData("overflow:hidden")]
    [InlineData("flex-direction:row-reverse")]
    [InlineData("", "order:2", "order:1")]
    public void HtmlFlexPdf_ColumnStructureContainsLaterParagraphBeforeEarlierPaintedCaption(
        string containerStyle, string primaryStyle = "", string sidebarStyle = "") {
        string html = CreateFlexLogicalOrderFixture(containerStyle, primaryStyle, sidebarStyle);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions {
            AutoFitWidePrintContent = false,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        });
        PdfCore.PdfReadDocument document = PdfCore.PdfReadDocument.Open(pdf);
        Assert.Equal(2, document.Pages.Count);
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(PdfCore.PdfInspector.Inspect(pdf).TaggedContent);
        PdfCore.PdfStructureElementInfo section = Assert.Single(tagged.StructureElements, item => item.StructureType == "Sect");
        var elements = tagged.StructureElements.ToDictionary(item => item.ObjectNumber);
        PdfCore.PdfStructureElementInfo[] columns = section.ChildElementObjectNumbers.Select(id => elements[id]).ToArray();
        Assert.Equal(new[] { "Div", "Div" }, columns.Select(item => item.StructureType));
        Assert.Equal(new[] { "P", "P" }, columns[0].ChildElementObjectNumbers.Select(id => elements[id].StructureType));
        Assert.Equal(new[] { "P" }, columns[1].ChildElementObjectNumbers.Select(id => elements[id].StructureType));

        // The second primary paragraph paints on page two; its source owner must
        // remain ahead of the sidebar's caption, which already painted on page one.
        Assert.DoesNotContain("ParagraphTwo", document.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("CaptionThree", document.Pages[0].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("ParagraphTwo", document.Pages[1].ExtractText(), StringComparison.Ordinal);
        string fullText = document.ExtractText();
        foreach (string marker in new[] { "ParagraphOneEnd", "ParagraphTwo", "CaptionThree", "ParagraphFour" })
            Assert.Equal(1, fullText.Split(new[] { marker }, StringSplitOptions.None).Length - 1);
    }

    [Fact]
    public void HtmlFlexPdf_IndependentRowsRetainSeparateLogicalParentScopes() {
        string html = "<style>@page{size:400px 400px;margin:10px}body,p{margin:0;font:14px/20px Arial}.row{display:flex}.later{order:-1}</style>"
            + "<section><div class='row'><p>FirstRowOne</p><p class='later'>FirstRowTwo</p></div>"
            + "<div class='row'><p>SecondRowOne</p><p class='later'>SecondRowTwo</p></div></section>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions { AutoFitWidePrintContent = false });
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(PdfCore.PdfInspector.Inspect(pdf).TaggedContent);
        var elements = tagged.StructureElements.ToDictionary(item => item.ObjectNumber);
        PdfCore.PdfStructureElementInfo section = Assert.Single(tagged.StructureElements, item => item.StructureType == "Sect");
        PdfCore.PdfStructureElementInfo[] rows = section.ChildElementObjectNumbers.Select(id => elements[id]).ToArray();
        Assert.Equal(new[] { "Div", "Div" }, rows.Select(item => item.StructureType));
        foreach (PdfCore.PdfStructureElementInfo row in rows) {
            Assert.Equal(2, row.ChildElementObjectNumbers.Count);
            foreach (int id in row.ChildElementObjectNumbers)
                Assert.Equal("P", elements[elements[id].ChildElementObjectNumbers.Single()].StructureType);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlFlexPdf_RepeatedTableFragmentsKeepOneColumnAndTableOwner(bool insideList) {
        string rows = string.Concat(Enumerable.Range(1, 18).Select(index => "<tr><td>Row" + index.ToString("D2") + "</td></tr>"));
        string tableHtml = "<table><thead><tr><th scope='col'>Header</th></tr></thead><tbody>" + rows
            + "</tbody><tfoot><tr><td>Footer</td></tr></tfoot></table>";
        if (insideList) tableHtml = "<ul><li>Item" + tableHtml + "</li></ul>";
        string html = "<style>@page{size:400px 180px;margin:10px}body,p,table,td,th{margin:0;font:14px/20px Arial}table{border-spacing:0}td,th{padding:0}</style>"
            + "<section style='display:flex;align-items:flex-start'><div style='width:240px;flex-shrink:0'>"
            + tableHtml + "</div><div style='width:100px;flex-shrink:0'><p>Sidebar</p></div></section>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions { AutoFitWidePrintContent = false });
        PdfCore.PdfReadDocument document = PdfCore.PdfReadDocument.Open(pdf);
        Assert.True(document.Pages.Count >= 3);
        Assert.Contains("Header", document.Pages[1].ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Footer", document.Pages[1].ExtractText(), StringComparison.Ordinal);
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(PdfCore.PdfInspector.Inspect(pdf).TaggedContent);
        PdfCore.PdfStructureElementInfo table = Assert.Single(tagged.StructureElements, item => item.StructureType == "Table");
        var elements = tagged.StructureElements.ToDictionary(item => item.ObjectNumber);
        Assert.Equal(insideList ? "LBody" : "Div", elements[table.ParentObjectNumber!.Value].StructureType);
        if (insideList) {
            Assert.Single(tagged.StructureElements, item => item.StructureType == "L");
            Assert.Single(tagged.StructureElements, item => item.StructureType == "LI");
            Assert.Single(tagged.StructureElements, item => item.StructureType == "Lbl");
        }
        Assert.Single(tagged.StructureElements, item => item.StructureType == "TH");
        Assert.Equal(19, tagged.StructureElements.Count(item => item.StructureType == "TD"));
        string allText = document.ExtractText();
        foreach (int row in Enumerable.Range(1, 18))
            Assert.Equal(1, allText.Split(new[] { "Row" + row.ToString("D2") }, StringSplitOptions.None).Length - 1);
    }

    [Fact]
    public void HtmlFlexPdf_DisplayContentsAndArtifactsKeepTheirExistingOwners() {
        string html = "<style>@page{size:400px 400px;margin:10px}body,p{margin:0;font:14px/20px Arial}</style>"
            + "<main style='display:flex'><section style='display:contents'><p>First</p><p>Second</p></section>"
            + "<div style='-officeimo-pdf-tag-type:artifact'>Decoration</div></main>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions { AutoFitWidePrintContent = false });
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(PdfCore.PdfInspector.Inspect(pdf).TaggedContent);
        var elements = tagged.StructureElements.ToDictionary(item => item.ObjectNumber);
        PdfCore.PdfStructureElementInfo main = Assert.Single(tagged.StructureElements,
            item => item.StructureType == "Sect" && item.ParentObjectNumber == tagged.StructureElements.Single(root => root.StructureType == "Document").ObjectNumber);
        PdfCore.PdfStructureElementInfo section = Assert.Single(main.ChildElementObjectNumbers.Select(id => elements[id]));
        Assert.Equal("Sect", section.StructureType);
        Assert.Equal(new[] { "Div", "Div" }, section.ChildElementObjectNumbers.Select(id => elements[id].StructureType));
        Assert.DoesNotContain("Decoration", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("section", false, "Sect")]
    [InlineData("section", true, "Sect")]
    [InlineData("ul", false, "L")]
    public void HtmlFlexPdf_FlattenedSemanticOwnerPrecedesReorderedSibling(string tag, bool nested, string expectedRole) {
        string contents = tag == "ul" ? "<li>First</li><li>Second</li>" : "<p>First</p><p>Second</p>";
        if (nested) contents = "<section style='display:contents'>" + contents + "</section>";
        string html = "<main style='display:flex'><" + tag + " style='display:contents'>" + contents
            + "</" + tag + "><div style='order:-1'>Third</div></main>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes();
        PdfCore.PdfTaggedContentInfo tagged = Assert.IsType<PdfCore.PdfTaggedContentInfo>(PdfCore.PdfInspector.Inspect(pdf).TaggedContent);
        var elements = tagged.StructureElements.ToDictionary(item => item.ObjectNumber);
        PdfCore.PdfStructureElementInfo document = Assert.Single(tagged.StructureElements, item => item.StructureType == "Document");
        PdfCore.PdfStructureElementInfo main = Assert.Single(document.ChildElementObjectNumbers.Select(id => elements[id]));
        Assert.Equal(new[] { expectedRole, "Div" }, main.ChildElementObjectNumbers.Select(id => elements[id].StructureType));
    }

    private static string CreateFlexLogicalOrderFixture(string containerStyle, string primaryStyle, string sidebarStyle) =>
        "<!doctype html><html lang='en'><head><style>@page{size:400px 400px;margin:10px}body,p,figcaption,section{margin:0;font:14px/20px Arial}</style></head><body>"
        + "<section style='display:flex;align-items:flex-start;" + containerStyle + "'><div style='width:240px;flex-shrink:0;" + primaryStyle + "'>"
        + "<p>ParagraphOne<br>Line02<br>Line03<br>Line04<br>Line05<br>Line06<br>Line07<br>Line08<br>Line09<br>Line10<br>Line11<br>Line12<br>Line13<br>Line14<br>Line15<br>Line16<br>Line17<br>Line18<br>Line19<br>ParagraphOneEnd</p><p>ParagraphTwo</p></div>"
        + "<div style='width:100px;flex-shrink:0;background:#ffdddd;" + sidebarStyle + "'><figcaption>CaptionThree</figcaption></div></section><p>ParagraphFour</p></body></html>";
}
