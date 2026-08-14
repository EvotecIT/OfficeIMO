using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActiveOutlineDestinationDictionaryCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var root = new PdfDictionary();
            PdfReference rootReference = AddObject(objects, root);
            var item = new PdfDictionary();
            item.Items["Title"] = new PdfStringObj("Protected destination");
            item.Items["Parent"] = rootReference;
            item.Items["Dest"] = candidate;
            PdfReference itemReference = AddObject(objects, item);
            root.Items["First"] = itemReference;
            root.Items["Last"] = itemReference;
            catalog.Items["Outlines"] = rootReference;
        });
    }

    [Fact]
    public void ActiveFormTransparencyGroupCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            var formDictionary = new PdfDictionary();
            formDictionary.Items["Type"] = new PdfName("XObject");
            formDictionary.Items["Subtype"] = new PdfName("Form");
            var bounds = new PdfArray();
            bounds.Items.Add(new PdfNumber(0));
            bounds.Items.Add(new PdfNumber(0));
            bounds.Items.Add(new PdfNumber(10));
            bounds.Items.Add(new PdfNumber(10));
            formDictionary.Items["BBox"] = bounds;
            formDictionary.Items["Group"] = candidate;
            PdfReference formReference = AddObject(objects, new PdfStream(formDictionary, Array.Empty<byte>()));
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var xObjects = new PdfDictionary();
            xObjects.Items["Fm1"] = formReference;
            var resources = new PdfDictionary();
            resources.Items["XObject"] = xObjects;
            page.Items["Resources"] = resources;
        });
    }

    [Fact]
    public void BoundedGraphRewriteMatchesTheUnboundedSerializationAtTheExactOutputLimit() {
        byte[] source = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] expected = PdfDocumentObjectGraphRewriter.Rewrite(source, null, null);

        byte[] actual = PdfDocumentObjectGraphRewriter.Rewrite(
            source,
            sourceReadOptions: null,
            outputEncryption: null,
            maximumOutputBytes: expected.LongLength);

        Assert.Equal(expected, actual);
    }
}
