using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfTaggedMergeTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void MergeRejectsNestedClassArraysWithoutChangingInputs(bool cyclic) {
        byte[] primary = Report("Primary", "en-US");
        byte[] incoming = WithClassArrays(cyclic ? 1 : 300, cyclic, shared: false);
        byte[] primarySnapshot = (byte[])primary.Clone();
        byte[] incomingSnapshot = (byte[])incoming.Clone();

        // The ordinary reader can inspect these pages: the failure belongs to class remapping.
        Assert.Single(PdfReadDocument.Open(incoming).Pages);
        var error = Assert.Throws<InvalidOperationException>(() => PdfDocument.MergeToBytes(primary, incoming));
        Assert.Contains("Nested structure class arrays", error.Message);
        Assert.Equal(primarySnapshot, primary);
        Assert.Equal(incomingSnapshot, incoming);
    }

    [Fact]
    public void MergeRetainsSharedClassArraysAndRenamesTheirClassOnce() {
        byte[] output = PdfDocument.MergeToBytes(Report("Primary", "en-US"), WithClassArrays(1, cyclic: false, shared: true));
        var read = PdfReadDocument.Open(output);
        var catalog = (PdfDictionary)read.Objects[read.Security.RootObjectNumber!.Value].Value;
        var root = (PdfDictionary)PdfObjectLookup.Resolve(read.Objects, catalog.Items["StructTreeRoot"])!;
        var classes = (PdfDictionary)PdfObjectLookup.Resolve(read.Objects, root.Items["ClassMap"])!;
        string renamed = Assert.Single(classes.Items.Keys);
        Assert.EndsWith(":Shared", renamed);
        var owners = read.Objects.Values.Select(item => item.Value).OfType<PdfDictionary>()
            .Where(dictionary => dictionary.Items.ContainsKey("C")).ToArray();
        Assert.Equal(2, owners.Length);
        var classReferences = owners.Select(owner => Assert.IsType<PdfReference>(owner.Items["C"])).ToArray();
        Assert.Single(classReferences.Select(reference => reference.ObjectNumber).Distinct());
        Assert.All(owners, owner => {
            var array = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(read.Objects, owner.Items["C"]));
            Assert.Equal(renamed, Assert.IsType<PdfName>(array.Items[0]).Name);
            Assert.Equal(1, Assert.IsType<PdfNumber>(array.Items[1]).Value);
        });
        var attributes = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(read.Objects, classes.Items[renamed]));
        var original = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(read.Objects, attributes.Items["ClassArraySnapshot"]));
        Assert.Equal("Shared", Assert.IsType<PdfName>(original.Items[0]).Name);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(output);
        Assert.Equal(2, independent.NumberOfPages);
        Assert.Contains("Incoming report", independent.GetPage(2).Text);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void MergeRejectsExcessivelyDeepIncomingOrPrimaryStructure(bool incoming) {
        byte[] deep = PdfDocumentObjectGraphRewriter.Rewrite(Report("Deep", "en-US"), null, null, (objects, security) => {
            var catalog = (PdfDictionary)objects[security.RootObjectNumber!.Value].Value;
            var root = (PdfDictionary)PdfObjectLookup.Resolve(objects, catalog.Items["StructTreeRoot"])!;
            PdfObject children = root.Items["K"];
            int next = objects.Keys.Min() - 1;
            for (int index = 0; index < 300; index++) {
                var element = new PdfDictionary();
                element.Items["Type"] = new PdfName("StructElem");
                element.Items["S"] = new PdfName("Part");
                element.Items["K"] = children;
                var reference = new PdfReference(next, 0);
                objects.Add(next, new PdfIndirectObject(next, 0, element));
                if (PdfObjectLookup.Resolve(objects, children) is PdfDictionary child) child.Items["P"] = reference;
                children = reference; next--;
            }
            ((PdfDictionary)PdfObjectLookup.Resolve(objects, children)!).Items["P"] = catalog.Items["StructTreeRoot"];
            root.Items["K"] = children;
            return security.InfoObjectNumber;
        });
        Assert.Single(PdfReadDocument.Open(deep).Pages);
        byte[] ordinary = Report("Ordinary", "en-US");
        var error = Assert.Throws<InvalidOperationException>(() =>
            PdfDocument.MergeToBytes(incoming ? new[] { ordinary, deep } : new[] { deep, ordinary }));
        Assert.Contains("nesting limit", error.Message);
    }

    private static byte[] WithClassArrays(int count, bool cyclic, bool shared) =>
        PdfDocumentObjectGraphRewriter.Rewrite(Report("Incoming", "en-US"), null, null, (objects, security) => {
            var catalog = (PdfDictionary)objects[security.RootObjectNumber!.Value].Value;
            var root = (PdfDictionary)PdfObjectLookup.Resolve(objects, catalog.Items["StructTreeRoot"])!;
            var classMap = new PdfDictionary();
            var attributes = new PdfDictionary(); attributes.Items["O"] = new PdfName("Layout");
            classMap.Items["Shared"] = attributes; root.Items["ClassMap"] = classMap;
            int first = objects.Keys.Min() - count;
            for (int index = 0; index < count; index++) {
                var array = new PdfArray();
                if (index + 1 < count) array.Items.Add(new PdfReference(first + index + 1, 0));
                else if (cyclic) array.Items.Add(new PdfReference(first, 0));
                else { array.Items.Add(new PdfName("Shared")); array.Items.Add(new PdfNumber(1)); }
                objects.Add(first + index, new PdfIndirectObject(first + index, 0, array));
            }
            if (shared) attributes.Items["ClassArraySnapshot"] = new PdfReference(first, 0);
            var elements = objects.Values.Select(item => item.Value).OfType<PdfDictionary>()
                .Where(dictionary => dictionary.Get<PdfName>("Type")?.Name == "StructElem" &&
                    dictionary.Get<PdfName>("S")?.Name != "Document").ToArray();
            foreach (var element in shared ? elements : elements.Take(1)) element.Items["C"] = new PdfReference(first, 0);
            return security.InfoObjectNumber;
        });
}
