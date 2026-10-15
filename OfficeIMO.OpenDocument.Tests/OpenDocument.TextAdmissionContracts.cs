using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentTextAdmissionContractTests {
    private const int DecodedCharacters = 16 * 1024 * 1024;

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void JoinedCollapsedTextReportsTheConfiguredCeilingWhenTheBudgetIsExhausted(int remaining) {
        OdsDocument document = (OdsDocument)LoadTextDocument("ods");
        XElement cell = document.GetXml("content.xml").Descendants(OdfNamespaces.Table + "table-cell").Single();
        cell.ReplaceNodes(new XElement(OdfNamespaces.Text + "p", new string(' ', DecodedCharacters - remaining)),
            new XElement(OdfNamespaces.Text + "p", "b"));

        AssertLimit(() => document.Sheets[0].GetValue(0, 0), "MaximumDecodedCharacters", DecodedCharacters);
    }

    [Theory]
    [InlineData("odt")]
    [InlineData("ods")]
    [InlineData("odp")]
    [InlineData("odg")]
    public void NativeTextGettersClassifyCompactWhitespaceExpansionAsAnAdmissionFailure(string family) {
        OdfDocument document = LoadTextDocument(family);
        document.GetXml("content.xml").Descendants(OdfNamespaces.Text + "p").Single()
            .ReplaceNodes(Spaces(DecodedCharacters + 1));

        AssertLimit(() => ReadText(document), "MaximumDecodedCharacters", DecodedCharacters);
    }

    [Fact]
    public void JoinedCellTextReportsTheConfiguredAggregateCharacterCeiling() {
        OdsDocument document = (OdsDocument)LoadTextDocument("ods");
        XElement cell = document.GetXml("content.xml").Descendants(OdfNamespaces.Table + "table-cell").Single();
        cell.ReplaceNodes(new XElement(OdfNamespaces.Text + "p", "a"),
            new XElement(OdfNamespaces.Text + "p", Spaces(DecodedCharacters - 1)));

        // Earlier content reduces the remaining allowance, not the configured aggregate ceiling.
        AssertLimit(() => document.Sheets[0].GetValue(0, 0), "MaximumDecodedCharacters", DecodedCharacters);
    }

    [Fact]
    public void NativeParagraphTextReportsTheContainerCeiling() {
        OdtDocument document = (OdtDocument)LoadTextDocument("odt");
        XElement paragraph = document.GetXml("content.xml").Descendants(OdfNamespaces.Text + "p").Single();
        paragraph.ReplaceNodes(Nested(OdfNamespaces.Text + "span", 128, new XText("leaf")));
        Assert.Equal("leaf", document.Paragraphs[0].Text);
        paragraph.ReplaceNodes(Nested(OdfNamespaces.Text + "span", 129, new XText("leaf")));

        AssertLimit(() => document.Paragraphs[0].Text, "MaximumContainerDepth", 128);
    }

    [Fact]
    public void NativeDrawListTraversalReportsTheContainerCeiling() {
        OdgDocument document = (OdgDocument)LoadTextDocument("odg");
        XElement textBox = document.GetXml("content.xml").Descendants(OdfNamespaces.Draw + "text-box").Single();
        textBox.ReplaceNodes(NestedLists(64));
        Assert.Equal("leaf", document.Pages[0].Shapes[0].Text);
        textBox.ReplaceNodes(NestedLists(65));

        AssertLimit(() => document.Pages[0].Shapes[0].Text, "MaximumContainerDepth", 128);
    }

    [Theory]
    [InlineData("selected")]
    [InlineData("nested")]
    [InlineData("paragraphs")]
    public void NativeTextAndParagraphTraversalReportTheVisitedNodeCeiling(string route) {
        bool paragraphs = route == "paragraphs";
        OdfDocument document = LoadTextDocument(paragraphs ? "odg" : "odt");
        XElement root = document.GetXml("content.xml").Descendants(OdfNamespaces.Text + "p").Single();
        if (paragraphs) root = root.Parent!;
        root.RemoveNodes();
        if (route == "nested") {
            var span = new XElement(OdfNamespaces.Text + "span"); root.Add(span); root = span;
        }
        int nodes = route == "nested" ? 99_999 : 100_000;
        for (int index = 0; index < nodes; index++)
            root.Add(paragraphs ? (XNode)new XElement(OdfNamespaces.Text + "p") : new XComment("node"));
        Assert.Equal(string.Empty, ReadText(document).Trim());
        root.Add(paragraphs ? (XNode)new XElement(OdfNamespaces.Text + "p") : new XComment("over-limit"));

        AssertLimit(() => ReadText(document), "MaximumVisitedElements", 100_000);
    }

    [Fact]
    public void LoadedContentSafetyReportsTheFragmentContextContainerCeiling() {
        OdtDocument document = (OdtDocument)LoadTextDocument("odt");
        document.GetXml("content.xml").Descendants(OdfNamespaces.Text + "p").Single()
            .ReplaceNodes(Nested(OdfNamespaces.Text + "span", 129, new XText("leaf")));

        AssertLimit(() => document.InspectLoadedContentSafety(), "MaximumContainerDepth", 128);
    }

    [Fact]
    public void LoadedContentSafetyReportsTheFragmentContextNodeCeiling() {
        OdtDocument document = (OdtDocument)LoadTextDocument("odt");
        XElement paragraph = document.GetXml("content.xml").Descendants(OdfNamespaces.Text + "p").Single();
        paragraph.RemoveNodes();
        for (int index = 0; index <= 100_000; index++) paragraph.Add(new XComment("context"));
        paragraph.Add(new XElement(OdfNamespaces.Text + "span", "leaf"));

        AssertLimit(() => document.InspectLoadedContentSafety(), "MaximumVisitedElements", 100_000);
    }

    [Theory]
    [InlineData("odt", false)]
    [InlineData("odt", true)]
    [InlineData("odp", false)]
    [InlineData("odp", true)]
    public void NativeTableCollectionsAdmitOneMillionItemsAndClassifyTheNextItem(string family, bool cells) {
        OdfDocument document = LoadTableDocument(family);
        XElement repeated = document.GetXml("content.xml").Descendants(OdfNamespaces.Table + (cells ? "table-cell" : "table-row")).Single();
        XName attribute = OdfNamespaces.Table + (cells ? "number-columns-repeated" : "number-rows-repeated");
        repeated.SetAttributeValue(attribute, 1_000_000);
        Assert.Equal(1_000_000, ReadTableCount(document, cells));
        repeated.SetAttributeValue(attribute, 1_000_001);

        AssertLimit(() => ReadTableCount(document, cells), "MaximumLogicalItems", 1_000_000);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PresentationTableAdmissionPrecedesRepeatSumOverflow(bool cells) {
        OdpPresentation document = (OdpPresentation)LoadTableDocument("odp");
        XElement first = document.GetXml("content.xml").Descendants(OdfNamespaces.Table + (cells ? "table-cell" : "table-row")).Single();
        var excessive = new XElement(first);
        excessive.SetAttributeValue(OdfNamespaces.Table + (cells ? "number-columns-repeated" : "number-rows-repeated"), long.MaxValue);
        first.AddAfterSelf(excessive);

        AssertLimit(() => ReadTableCount(document, cells), "MaximumLogicalItems", 1_000_000);
    }

    private static void AssertLimit(Func<object?> read, string name, long maximum) {
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(read);
        Assert.Equal("opendocument", error.Owner);
        Assert.Equal(name, error.LimitName);
        Assert.Equal(maximum, error.Maximum);
    }

    private static OdfDocument LoadTextDocument(string family) {
        OdfDocument source;
        if (family == "odt") {
            var document = OdtDocument.Create(); document.AddParagraph("seed"); source = document;
        } else if (family == "ods") {
            var document = OdsDocument.Create(); document.AddSheet("Data").Cell(0, 0).SetString("seed"); source = document;
        } else if (family == "odp") {
            var document = OdpPresentation.Create();
            document.AddSlide().AddTextBox(OdfRect.FromCentimeters(1, 1, 5, 3), "seed"); source = document;
        } else {
            var document = OdgDocument.Create();
            document.AddPage().Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 5, 3), "seed"); source = document;
        }
        using var input = new MemoryStream(source.ToBytes());
        return OdfDocument.Load(input);
    }

    private static string ReadText(OdfDocument document) => document switch {
        OdtDocument text => text.Paragraphs[0].Text,
        OdsDocument spreadsheet => spreadsheet.Sheets[0].GetValue(0, 0).ToString(),
        OdpPresentation presentation => presentation.Slides[0].Shapes[0].TextParagraphs[0].Text,
        OdgDocument drawing => drawing.Pages[0].Shapes[0].Text,
        _ => throw new ArgumentOutOfRangeException(nameof(document))
    };

    private static OdfDocument LoadTableDocument(string family) {
        OdfDocument source;
        if (family == "odt") {
            var document = OdtDocument.Create(); document.AddTable(1, 1); source = document;
        } else {
            var document = OdpPresentation.Create();
            document.AddSlide().AddTable(OdfRect.FromCentimeters(1, 1, 5, 3), 1, 1); source = document;
        }
        using var input = new MemoryStream(source.ToBytes());
        return OdfDocument.Load(input);
    }

    private static int ReadTableCount(OdfDocument document, bool cells) {
        if (document is OdtDocument text) return cells ? text.Tables[0].Rows[0].Cells.Count : text.Tables[0].Rows.Count;
        var table = ((OdpPresentation)document).Slides[0].Shapes.OfType<OdpTable>().Single();
        return cells ? table.Rows[0].Cells.Count : table.Rows.Count;
    }

    private static XElement Spaces(int count) => new(OdfNamespaces.Text + "s", new XAttribute(OdfNamespaces.Text + "c", count));

    private static XNode Nested(XName name, int count, XNode leaf) {
        XNode result = leaf;
        for (int index = 0; index < count; index++) result = new XElement(name, result);
        return result;
    }

    private static XNode NestedLists(int count) {
        XNode result = new XElement(OdfNamespaces.Text + "p", "leaf");
        for (int index = 0; index < count; index++) result = new XElement(OdfNamespaces.Text + "list",
            new XElement(OdfNamespaces.Text + "list-item", result));
        return result;
    }
}
