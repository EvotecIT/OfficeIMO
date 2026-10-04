using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsStructureTests {
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void LoadedDocumentsCanInsertMoveRemoveAndReopenWithoutLosingNativeParts(XpsFormat format) {
        var created = XpsDocument.Create(format); created.AddPage(100, 80).AddPath("M0,0H20V20Z");
        created.AddResource("Metadata/opaque.bin", new byte[] { 7, 9 }, "application/octet-stream");
        var doc = XpsDocument.Load(created.Save()); var original = doc.Pages[0];
        var first = doc.Documents[0]; var second = doc.AddDocument();
        var inserted = first.InsertPage(0, 120, 90); second.InsertPage(0, original);
        Assert.Same(original, doc.Pages[2]);
        first.MovePage(1, 0); Assert.Same(original, doc.Pages[0]);
        first.MovePageTo(1, second, 0); Assert.Same(inserted, second.Pages[0]);
        doc.MoveDocument(1, 0); Assert.Same(inserted, doc.Pages[0]);
        doc.RemoveDocumentAt(1); second.RemovePageAt(1);
        var reopened = XpsDocument.Load(doc.Save());
        Assert.Single(reopened.Documents); Assert.Single(reopened.Pages);
        Assert.Equal(120, reopened.Pages[0].Width);
        Assert.Contains(original.PartName, reopened.PartNames);
        Assert.Equal(new byte[] { 7, 9 }, reopened.GetPartBytes("Metadata/opaque.bin"));
        Assert.Equal(doc.Save(), doc.Save());
        // Removed backing objects are still usable, and their edits survive restoration.
        original.AddPath("M2,2H10V10Z", "#FFFF0000");
        doc.InsertDocument(0, first);
        Assert.Same(original, doc.Pages[0]);
        Assert.Equal("#FFFF0000", (string?)XpsDocument.Load(doc.Save()).Pages[0].GetMarkup().Elements().Last().Attribute("Fill"));
    }

    [Fact]
    public void RepeatedReferencesShareStructureAndNavigationIsReindexed() {
        var doc = XpsDocument.Create(); var one = doc.AddPage(100, 100); var two = doc.AddPage(100, 100);
        var pageXml = two.GetMarkup(); XNamespace ns = pageXml.Name.Namespace;
        pageXml.Add(new XElement(ns + "Path", new XAttribute("Name", "target"), new XAttribute("Data", "M0,0H10V10Z"), new XAttribute("Fill", "#FF000000")));
        two.ReplaceMarkup(pageXml);
        one.AddPath("M0,0H10V10Z"); var source = one.GetMarkup(); source.Elements().Single().SetAttributeValue("FixedPage.NavigateUri", "/Documents/1/FixedDocument.fdoc#target"); one.ReplaceMarkup(source);
        byte[] bytes = XpsDocumentTests.Rewrite(doc.Save(), doc.Documents[0].PartName, b => {
            var xml = XElement.Parse(Encoding.UTF8.GetString(b)); xml.SetAttributeValue("custom", "retained");
            xml.Elements().Last().SetAttributeValue("custom-page", "retained");
            xml.Elements().Last().Add(new XElement(ns + "PageContent.LinkTargets", new XElement(ns + "LinkTarget", new XAttribute("Name", "target"))));
            return Encoding.UTF8.GetBytes(xml.ToString());
        });
        doc = XpsDocument.Load(bytes); var document = doc.Documents[0];
        doc.InsertDocument(1, document); Assert.Same(doc.Documents[0], doc.Documents[1]);
        document.MovePage(1, 0); Assert.Same(doc.Pages[0], doc.Pages[2]);
        Assert.Equal("retained", (string?)document.GetMarkup().Attribute("custom"));
        Assert.Equal("retained", (string?)document.GetMarkup().Elements().First().Attribute("custom-page"));
        Assert.Contains("page-1.svg#xps-target", doc.Pages[1].ToSvg().Svg);
        var other = doc.AddDocument(); document.MovePageTo(0, other, 0);
        Assert.Equal("retained", (string?)other.GetMarkup().Elements().Single().Attribute("custom-page"));
        Assert.Single(other.GetMarkup().Descendants(ns + "LinkTarget"));
        Assert.Equal(3, doc.Pages.Count);
        Assert.Contains("page-3.svg#xps-target", doc.Pages[0].ToSvg().Svg);
    }

    [Fact]
    public void FailedRepeatedDocumentEditsAreAtomicAndEmptySequenceCanAppend() {
        var created = XpsDocument.Create(); created.AddPage(100, 100);
        var doc = XpsDocument.Load(created.Save(), new XpsReadOptions { MaximumPages = 2 });
        var first = doc.Documents[0]; doc.InsertDocument(1, first);
        byte[] before = doc.Save(); string[] parts = doc.PartNames.ToArray();
        Assert.Throws<InvalidDataException>(() => first.AddPage());
        Assert.Equal(before, doc.Save()); Assert.Equal(parts, doc.PartNames);
        Assert.Throws<ArgumentOutOfRangeException>(() => doc.MoveDocument(0, 2));
        Assert.Throws<ArgumentException>(() => first.InsertPage(0, XpsDocument.Create().AddPage()));
        doc.RemoveDocumentAt(1); doc.RemoveDocumentAt(0);
        var page = doc.AddPage(200, 300);
        Assert.Single(doc.Documents); Assert.Same(page, doc.Pages.Single());
        Assert.NotEqual(first.PartName, doc.Documents[0].PartName);
        Assert.Single(XpsDocument.Load(doc.Save()).Pages);
    }

    [Fact]
    public void CrossDocumentMoveWithRepeatedDestinationHonorsSequencePageLimitAtomically() {
        var created = XpsDocument.Create(); created.AddPage(100, 100); created.AddDocument();
        var doc = XpsDocument.Load(created.Save(), new XpsReadOptions { MaximumPages = 1 });
        var target = doc.Documents[1]; doc.InsertDocument(2, target);
        byte[] before = doc.Save();
        Assert.Throws<InvalidDataException>(() => doc.Documents[0].MovePageTo(0, target, 0));
        Assert.Equal(before, doc.Save()); Assert.Single(doc.Pages);
    }
    [Fact]
    public void ResourceReplacementIsSharedAndCannotOverwriteStructure() {
        var doc = XpsDocument.Create(); doc.AddPage();
        doc.AddResource("data.bin", new byte[] { 1 }, "application/octet-stream");
        byte[] replacement = { 2, 3 }; doc.ReplaceResource("data.bin", replacement); replacement[0] = 9;
        Assert.Equal(new byte[] { 2, 3 }, XpsDocument.Load(doc.Save()).GetPartBytes("data.bin"));
        Assert.Throws<ArgumentException>(() => doc.ReplaceResource(doc.Pages[0].PartName, new byte[] { 1 }));
        Assert.Throws<ArgumentException>(() => doc.ReplaceResource(doc.Documents[0].PartName, new byte[] { 1 }));
    }
    [Fact]
    public void PageReplacementRespectsLoadedXmlLimitsAndKeepsOriginalOnFailure() {
        var created = XpsDocument.Create(); created.AddPage();
        var doc = XpsDocument.Load(created.Save(), new XpsReadOptions { MaximumXmlDepth = 2 });
        var page = doc.Pages[0]; var xml = page.GetMarkup(); XNamespace ns = xml.Name.Namespace;
        xml.Add(new XElement(ns + "Canvas", new XElement(ns + "Canvas", new XElement(ns + "Canvas"))));
        Assert.Throws<InvalidDataException>(() => page.ReplaceMarkup(xml));
        Assert.Empty(page.GetMarkup().Elements());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void StructuralEditsBudgetRegeneratedMetadataBeforeMutation(bool partLimit) {
        var created = XpsDocument.Create(); created.AddPage(); byte[] initial = created.Save();
        var baseline = XpsDocument.Load(initial);
        int typesLength = baseline.GetPartBytes("[Content_Types].xml").Length;
        long expanded = baseline.PartNames.Sum(n => (long)baseline.GetPartBytes(n).Length);
        var limits = partLimit ? new XpsReadOptions { MaximumPartBytes = typesLength }
            : new XpsReadOptions { MaximumExpandedBytes = expanded + 220 };
        var doc = XpsDocument.Load(initial, limits); byte[] before = doc.Save();
        Assert.Throws<InvalidOperationException>(() => doc.AddDocument());
        Assert.Single(doc.Documents); Assert.Equal(before, doc.Save());
    }

    [Fact]
    public void ResourceAdditionBudgetsGeneratedContentTypesAtomically() {
        var created = XpsDocument.Create(); created.AddPage(); byte[] initial = created.Save();
        var baseline = XpsDocument.Load(initial);
        var doc = XpsDocument.Load(initial, new XpsReadOptions { MaximumPartBytes = baseline.GetPartBytes("[Content_Types].xml").Length });
        Assert.Throws<InvalidOperationException>(() => doc.AddResource("new.bin", new byte[] { 1 }, "application/octet-stream"));
        Assert.DoesNotContain("new.bin", doc.PartNames); Assert.Equal(initial, doc.Save());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ResourceEditsUseSpaceReleasedByCurrentPageMarkup(bool addition) {
        var created = XpsDocument.Create(); created.AddPage().AddPath("M0,0" + new string(' ', 4000) + "L1,1");
        created.AddResource("data.bin", new byte[] { 1 }, "application/octet-stream");
        byte[] initial = created.Save(); var baseline = XpsDocument.Load(initial);
        long expanded = baseline.PartNames.Sum(n => (long)baseline.GetPartBytes(n).Length);
        var doc = XpsDocument.Load(initial, new XpsReadOptions { MaximumExpandedBytes = expanded });
        var markup = doc.Pages[0].GetMarkup(); markup.RemoveNodes(); doc.Pages[0].ReplaceMarkup(markup);
        if (addition) doc.AddResource("other.bin", new byte[100], "application/octet-stream");
        else doc.ReplaceResource("data.bin", new byte[100]);
        Assert.Single(XpsDocument.Load(doc.Save(), new XpsReadOptions { MaximumExpandedBytes = expanded }).Pages);
    }

}
