using System;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgDocumentFieldTests {
    private const string NumberCache = "cached-number";
    private const string CountCache = "cached-count";

    [Fact]
    public void DocumentAndStandaloneFieldsRefreshAfterPageAndNumberingEditsWithoutChangingSourceCaches() {
        var document = OdgDocument.Create();
        var first = document.AddPage("First");
        var second = document.AddPage("Second");
        second.MasterPageName = first.MasterPageName;
        AddMarker(first); AddMarker(second);
        var paragraph = first.MasterShapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 14, 8), "Page [", "Header").Paragraphs[0];
        var number = paragraph.AddField(OdfTextFieldKind.PageNumber, NumberCache);
        paragraph.AddText("] of ");
        paragraph.AddField(OdfTextFieldKind.PageCount, CountCache).NumberFormat = "1";

        AssertProjection(document, new[] { "First", "Second" }, new[] { "1", "2" }, "2");
        document.MovePage(1, 0);
        AssertProjection(document, new[] { "Second", "First" }, new[] { "1", "2" }, "2");

        var copy = document.ClonePage(1, "Copy");
        copy.Shapes[0].Text = copy.Name;
        AssertProjection(document, new[] { "Second", "First", "Copy" }, new[] { "1", "2", "3" }, "3");
        var added = document.AddPage("Added");
        added.MasterPageName = first.MasterPageName;
        AddMarker(added);
        AssertProjection(document, new[] { "Second", "First", "Copy", "Added" }, new[] { "1", "2", "3", "4" }, "4");

        document.RemovePage(1);
        var remaining = new[] { "Second", "Copy", "Added" };
        AssertProjection(document, remaining, new[] { "1", "2", "3" }, "3");
        AssertDetachedProjection(document, first);

        string layoutName = (string)first.Master!.Attribute(OdfNamespaces.Style + "page-layout-name")!;
        var layout = document.GetXml("styles.xml").Descendants(OdfNamespaces.Style + "page-layout")
            .Single(element => (string?)element.Attribute(OdfNamespaces.Style + "name") == layoutName)
            .Element(OdfNamespaces.Style + "page-layout-properties")!;
        layout.SetAttributeValue(OdfNamespaces.Style + "num-format", "I");
        document.MarkPartDirty("styles.xml");
        AssertProjection(document, remaining, new[] { "I", "II", "III" }, "3");

        number.NumberFormat = "a";
        AssertProjection(document, remaining, new[] { "a", "b", "c" }, "3");
        number.NumberFormat = null;
        AssertProjection(document, remaining, new[] { "I", "II", "III" }, "3");

        number.NumberFormat = "1";
        number.PageSelection = OdfTextFieldPageSelection.Next;
        number.PageAdjustment = -1;
        AssertProjection(document, remaining, new[] { "1", "2", "" }, "3");
        number.PageSelection = OdfTextFieldPageSelection.Current;
        number.PageAdjustment = 0;
        AssertProjection(document, remaining, new[] { "1", "2", "3" }, "3");
    }

    private static void AssertProjection(OdgDocument document, string[] expectedNames, string[] expectedNumbers, string expectedCount) {
        string[] before = Parts(document);
        var pages = document.Pages;
        Assert.Equal(expectedNames, pages.Select(page => page.Name));
        var projected = document.ToDrawings(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        Assert.Equal(expectedNames.Length, projected.Value.Count);
        Assert.Equal(before, Parts(document));
        for (int index = 0; index < pages.Count; index++) {
            string[] expectedText = { "Page [" + expectedNumbers[index] + "] of " + expectedCount, expectedNames[index] };
            Assert.Equal(expectedText, Text(projected.Value[index]));
            var standalone = pages[index].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            Assert.Equal(expectedText, Text(standalone.Value));
            Assert.Equal(new[] { NumberCache, CountCache }, pages[index].MasterShapes[0].Paragraphs[0].Fields.Select(field => field.DisplayText));
            Assert.Equal(before, Parts(document));
        }
    }

    private static void AssertDetachedProjection(OdgDocument document, OdgPage removed) {
        string[] before = Parts(document);
        string detachedXml = removed.Element.ToString(SaveOptions.DisableFormatting);
        var projected = removed.ToDrawing();
        Assert.Equal(new[] { "Page [" + NumberCache + "] of " + CountCache, removed.Name }, Text(projected.Value));
        Assert.Contains(projected.Report.Mappings, mapping => mapping.Feature.EndsWith(":field-page-context", StringComparison.Ordinal) &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => removed.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal(new[] { NumberCache, CountCache }, removed.MasterShapes[0].Paragraphs[0].Fields.Select(field => field.DisplayText));
        Assert.Equal(before, Parts(document));
        Assert.Equal(detachedXml, removed.Element.ToString(SaveOptions.DisableFormatting));
    }

    private static void AddMarker(OdgPage page) =>
        page.Shapes.AddTextBox(OdfRect.FromCentimeters(1, 10, 14, 2), page.Name, "Marker");

    private static string[] Text(OfficeDrawing drawing) => drawing.Elements.OfType<OfficeDrawingRichText>().Select(text => text.PlainText).ToArray();

    private static string[] Parts(OdgDocument document) => new[] {
        document.GetXml("content.xml").ToString(SaveOptions.DisableFormatting),
        document.GetXml("styles.xml").ToString(SaveOptions.DisableFormatting)
    };
}
