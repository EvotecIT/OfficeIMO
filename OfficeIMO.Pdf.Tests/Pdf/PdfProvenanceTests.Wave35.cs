using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NamedJavaScriptActionGraphsCannotOwnProvenanceAssociations(bool candidateIsLeafAction) {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            PdfReference leaf;
            if (candidateIsLeafAction) {
                PdfDictionary candidateDictionary = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
                candidateDictionary.Items["S"] = new PdfName("JavaScript");
                candidateDictionary.Items["JS"] = new PdfStringObj("app.alert('candidate')", true);
                leaf = candidate;
            } else {
                var action = new PdfDictionary();
                action.Items["S"] = new PdfName("JavaScript");
                action.Items["JS"] = new PdfStringObj("app.alert('next')", true);
                action.Items["Next"] = candidate;
                leaf = AddObject(objects, action);
            }
            var entries = new PdfArray();
            entries.Items.Add(new PdfStringObj("Startup", true));
            entries.Items.Add(leaf);
            var tree = new PdfDictionary();
            tree.Items["Names"] = entries;
            names.Items["JavaScript"] = AddObject(objects, tree);
            catalog.Items["AF"] = ArrayWith(candidate);
        });
    }

    [Fact]
    public void CallerSpecificDecodeLimitIsNotReclassifiedAsAggregateBudget() {
        var limits = new PdfReadLimits {
            MaxDecodedStreamBytes = 16,
            MaxTotalDecodedStreamBytes = 10
        };
        var budget = new PdfDecodedStreamBudget(limits, initialUsedBytes: 2);
        var stream = new PdfStream(new PdfDictionary(), new byte[5]);

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            budget.Decode(stream, new Dictionary<int, PdfIndirectObject>(), maximumRequestedBytes: 4));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(4, exception.Limit);
    }

    [Fact]
    public void RequiredDecodeRevalidatesAPermissiveCacheEntry() {
        var dictionary = new PdfDictionary();
        dictionary.Items["Filter"] = new PdfName("DCTDecode");
        var stream = new PdfStream(dictionary, new byte[] { 1, 2, 3 });
        var budget = new PdfDecodedStreamBudget(new PdfReadLimits { MaxDecodedStreamBytes = 16 });
        var objects = new Dictionary<int, PdfIndirectObject>();

        Assert.Equal(stream.Data, budget.Decode(stream, objects, maximumRequestedBytes: 16));
        Assert.Throws<InvalidDataException>(() => budget.DecodeRequired(stream, objects, maximumRequestedBytes: 16));
        Assert.Equal(stream.Data.Length, budget.UsedBytes);
    }
}
