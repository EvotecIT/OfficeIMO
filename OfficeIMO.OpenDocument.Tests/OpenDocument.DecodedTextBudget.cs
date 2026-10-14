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

        Assert.Throws<OfficeInspectionResourceLimitException>(() => OdfTextCodec.ReadJoined(new[] { first, second }));
        Assert.Equal("one\ntwo  ", OdfTextCodec.ReadJoined(new[] {
            new XElement(OdfNamespaces.Text + "p", "one"),
            new XElement(OdfNamespaces.Text + "p", "two", new XElement(OdfNamespaces.Text + "s",
                new XAttribute(OdfNamespaces.Text + "c", 2)))
        }));
    }

    [Theory]
    [InlineData("odt", false)]
    [InlineData("odt", true)]
    [InlineData("odp", false)]
    [InlineData("odp", true)]
    public void InlineSnapshotsKeepTheEnclosingCeilingAcrossAnnotationOrder(string family, bool annotationFirst) {
        OdfDocument document;
        Action read;
        Func<string> readText;
        if (family == "odt") {
            var text = OdtDocument.Create(); OdtParagraph paragraph = text.AddParagraph("visible");
            document = text; read = () => _ = paragraph.InlineNodes; readText = () => paragraph.Text;
        } else {
            var presentation = OdpPresentation.Create();
            OdpParagraph paragraph = presentation.AddSlide("Slide").AddTextBox(OdfRect.FromCentimeters(1, 1, 10, 3), null, "Body").AddParagraph("visible");
            document = presentation; read = () => _ = paragraph.InlineNodes; readText = () => paragraph.Text;
        }
        XElement element = document.GetXml("content.xml").Descendants(OdfNamespaces.Text + "p").Single();
        AddLargeAnnotations(element);
        if (annotationFirst) {
            XText visible = element.Nodes().OfType<XText>().Single(); visible.Remove(); element.Add(visible);
        }
        Assert.Equal("visible", readText());
        OfficeInspectionResourceLimitException error = Assert.Throws<OfficeInspectionResourceLimitException>(read);
        Assert.Equal("opendocument", error.Owner); Assert.Equal("MaximumDecodedCharacters", error.LimitName);
        Assert.Equal(2 * HalfLimit, error.Maximum);
    }

    [Fact]
    public void OdsHeaderTextRejectsExpansionAcrossParagraphs() {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        OdsCell cell = sheet.Cell(0, 0);
        cell.SetString("Header");
        cell.Element.Elements(OdfNamespaces.Text + "p").Remove();
        cell.Element.Add(ParagraphWithSpaces(HalfLimit), ParagraphWithSpaces(HalfLimit));

        Assert.Throws<OfficeInspectionResourceLimitException>(() => _ = cell.Text);
        OdsDataPilotTable pivot = document.AddDataPilotTable("Pivot", "Data.A1:Data.A2", "Data.C1:Data.D2");
        Assert.Throws<OfficeInspectionResourceLimitException>(() => pivot.AddField("Header", "row"));
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
