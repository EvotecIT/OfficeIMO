using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfTaggedMergeTests {
    [Fact]
    public void MergePreservesBothTaggedReportsAndTheirPageOwners() {
        byte[] first = Report("First", "en-US");
        byte[] second = Report("Second", "fr-FR");
        byte[] firstSnapshot = (byte[])first.Clone();
        byte[] secondSnapshot = (byte[])second.Clone();
        byte[] output = PdfDocument.MergeToBytes(first, second);

        using var independent = UglyToad.PdfPig.PdfDocument.Open(output);
        Assert.Equal(2, independent.NumberOfPages);
        Assert.Contains("First report", independent.GetPage(1).Text);
        Assert.Contains("Second report", independent.GetPage(2).Text);
        var read = PdfReadDocument.Open(output);
        var tags = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(output).TaggedContent);
        Assert.Equal(2, tags.StructureElements.Count(element => element.StructureType == "H1"));
        Assert.Single(tags.StructureElements, element => element.StructureType == "Document");
        AssertParentTreeMatchesPages(read, 2);
        Assert.Equal(firstSnapshot, first);
        Assert.Equal(secondSnapshot, second);
    }

    [Fact]
    public void RepeatedSourceGetsDistinctStructureIdsAndParentKeys() {
        byte[] source = PdfDocumentObjectGraphRewriter.Rewrite(PdfTaggedPageOrganizationTests.CreateTaggedSource(),
            null, null, (objects, security) => {
                // This fixture tests repeated semantic ownership, not active-action support.
                ((PdfDictionary)objects[security.RootObjectNumber!.Value].Value).Items.Remove("OpenAction");
                foreach (var page in objects.Values.Select(item => item.Value).OfType<PdfDictionary>()
                    .Where(item => item.Get<PdfName>("Type")?.Name == "Page")) page.Items.Remove("AA");
                return security.InfoObjectNumber;
            });
        byte[] snapshot = (byte[])source.Clone();
        byte[] output = PdfDocument.MergeToBytes(source, source, source);
        var read = PdfReadDocument.Open(output);
        Assert.Equal(6, read.Pages.Count);
        AssertParentTreeMatchesPages(read, 6);
        var tags = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(output).TaggedContent);
        Assert.Single(tags.StructureElements, element => element.StructureType == "Document");
        var ids = tags.StructureElements.Select(element => ((PdfDictionary)read.Objects[element.ObjectNumber].Value)
            .Get<PdfStringObj>("ID")?.Value).Where(id => id != null).ToArray();
        Assert.Equal(ids.Length, ids.Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(snapshot, source);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void MixedInputsKeepExistingTagsWithoutInventingTagsForPlainPages(bool taggedFirst) {
        byte[] tagged = Report("Tagged", "en-US");
        byte[] plain = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Plain attachment")).ToBytes();
        byte[] output = PdfDocument.MergeToBytes(taggedFirst ? new[] { tagged, plain } : new[] { plain, tagged });
        var read = PdfReadDocument.Open(output);
        Assert.Equal(2, read.Pages.Count);
        Assert.NotNull(PdfInspector.Inspect(output).TaggedContent);
        int taggedPage = taggedFirst ? 0 : 1;
        Assert.True(read.Pages[taggedPage].PageDictionary.Items.ContainsKey("StructParents"));
        Assert.False(read.Pages[1 - taggedPage].PageDictionary.Items.ContainsKey("StructParents"));
        using var independent = UglyToad.PdfPig.PdfDocument.Open(output);
        Assert.Contains("Tagged report", independent.GetPage(taggedPage + 1).Text);
        Assert.Contains("Plain attachment", independent.GetPage(2 - taggedPage).Text);
    }

    private static byte[] Report(string title, string language) => PdfDocument.Create(new PdfCreateOptions {
        Language = language
    }).TaggedPdfCatalogMarkers().H1(title + " report")
        .Paragraph(paragraph => paragraph.Text(title + " report details")).ToBytes();

    private static void AssertParentTreeMatchesPages(PdfReadDocument read, int expectedTaggedPages) {
        var catalog = (PdfDictionary)read.Objects[read.Security.RootObjectNumber!.Value].Value;
        var root = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(read.Objects, catalog.Items["StructTreeRoot"]));
        var parentTree = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(read.Objects, root.Items["ParentTree"]));
        var nums = Assert.IsType<PdfArray>(parentTree.Items["Nums"]);
        int[] keys = Enumerable.Range(0, nums.Items.Count / 2)
            .Select(index => (int)((PdfNumber)nums.Items[index * 2]).Value).ToArray();
        Assert.Equal(expectedTaggedPages, keys.Length);
        Assert.Equal(keys.Length, keys.Distinct().Count());
        Assert.All(read.Pages, page => Assert.Contains((int)page.PageDictionary.Get<PdfNumber>("StructParents")!.Value, keys));
    }
}
