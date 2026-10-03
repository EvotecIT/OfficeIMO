using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentDecodedTextBudgetTests {
    private const int HalfLimit = 8 * 1024 * 1024;

    [Fact]
    public void JoinedParagraphsCountSeparatorsAgainstTheDecodedTextLimit() {
        XElement first = ParagraphWithSpaces(HalfLimit);
        XElement second = ParagraphWithSpaces(HalfLimit);

        Assert.Throws<InvalidDataException>(() => OdfTextCodec.ReadJoined(new[] { first, second }));
        Assert.Equal("one\ntwo  ", OdfTextCodec.ReadJoined(new[] {
            new XElement(OdfNamespaces.Text + "p", "one"),
            new XElement(OdfNamespaces.Text + "p", "two", new XElement(OdfNamespaces.Text + "s",
                new XAttribute(OdfNamespaces.Text + "c", 2)))
        }));
    }

    [Fact]
    public void OdtAndOdpInlineSnapshotsCountNonvisibleOtherNodes() {
        OdtDocument text = OdtDocument.Create();
        OdtParagraph odtParagraph = text.AddParagraph("visible");
        AddLargeAnnotations(odtParagraph.Element);
        Assert.Equal("visible", odtParagraph.Text);
        Assert.Throws<InvalidDataException>(() => _ = odtParagraph.InlineNodes);

        OdpPresentation presentation = OdpPresentation.Create();
        OdpParagraph odpParagraph = presentation.AddSlide("Slide")
            .AddTextBox(OdfRect.FromCentimeters(1, 1, 10, 3), null, "Body")
            .AddParagraph("visible");
        XElement odpElement = presentation.Package.GetXml("content.xml")
            .Descendants(OdfNamespaces.Text + "p").Single(element => element.Value == "visible");
        AddLargeAnnotations(odpElement);
        Assert.Equal("visible", odpParagraph.Text);
        Assert.Throws<InvalidDataException>(() => _ = odpParagraph.InlineNodes);
    }

    [Fact]
    public void OdsHeaderTextRejectsExpansionAcrossParagraphs() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        OdsCell cell = sheet.Cell(0, 0);
        cell.SetString("Header");
        cell.Element.Elements(OdfNamespaces.Text + "p").Remove();
        cell.Element.Add(ParagraphWithSpaces(HalfLimit), ParagraphWithSpaces(HalfLimit));

        Assert.Throws<InvalidDataException>(() => _ = cell.Text);
        OdsDataPilotTable pivot = document.AddDataPilotTable("Pivot", "Data.A1:Data.A2", "Data.C1:Data.D2");
        Assert.Throws<InvalidDataException>(() => pivot.AddField("Header", "row"));
    }

    [Fact]
    public void UnreadableEmbeddedChartTitleIsOmittedWithoutExpandingIt() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "microsoft-excel-column-chart.ods");
        OdsDocument document = OdsDocument.Load(path);
        XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
        XElement title = document.Package.GetXml("Object 1/content.xml")
            .Descendants(chart + "title").Single();
        title.RemoveNodes();
        title.Add(ParagraphWithSpaces(HalfLimit), ParagraphWithSpaces(HalfLimit));

        Assert.Empty(document.GetSheet("Data")!.Charts);
    }

    private static XElement ParagraphWithSpaces(int count) => new(OdfNamespaces.Text + "p",
        new XElement(OdfNamespaces.Text + "s", new XAttribute(OdfNamespaces.Text + "c", count)));

    private static void AddLargeAnnotations(XElement paragraph) {
        paragraph.Add(new XElement(OdfNamespaces.Office + "annotation", ParagraphWithSpaces(HalfLimit)),
            new XElement(OdfNamespaces.Office + "annotation", ParagraphWithSpaces(HalfLimit)));
    }
}
