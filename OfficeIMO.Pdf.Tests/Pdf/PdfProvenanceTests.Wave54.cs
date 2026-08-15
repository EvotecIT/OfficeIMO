using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void DefaultParserBudgetBoundsAggregateDecodedStreamCaching() {
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            new PdfDecodedStreamBudget(
                PdfReadLimits.Default,
                PdfReadLimits.DefaultMaxTotalDecodedStreamBytes + 1));

        Assert.Equal(PdfReadLimitKind.TotalDecodedStreamBytes, exception.Kind);
    }

    [Fact]
    public void IndirectEmbeddedFileStreamTypeCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var dictionary = new PdfDictionary();
            dictionary.Items["Type"] = AddObject(objects, new PdfName("EmbeddedFile"));
            dictionary.Items["AF"] = ArrayWith(candidate);
            catalog.Items["PrivateStream"] = AddObject(objects, new PdfStream(dictionary, Array.Empty<byte>()));
        });
    }

    [Fact]
    public void OrdinaryNameTreeLeafGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            var entries = new PdfArray();
            entries.Items.Add(new PdfStringObj("resource", true));
            entries.Items.Add(candidate);
            var extensionTree = new PdfDictionary();
            extensionTree.Items["Names"] = entries;
            names.Items["PrivateExtension"] = AddObject(objects, extensionTree);
        });
    }
}
