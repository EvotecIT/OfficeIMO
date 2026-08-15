using System.IO.Compression;
using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ProvenanceRemovalDeletesWholeNumberTreePairs() {
        byte[] original = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] withNumberTree = PdfDocumentObjectGraphRewriter.Rewrite(original, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            var numbers = new PdfArray();
            numbers.Items.Add(new PdfNumber(0));
            numbers.Items.Add(candidate);
            var privateTree = new PdfDictionary();
            privateTree.Items["Nums"] = numbers;
            catalog.Items["PrivateTree"] = privateTree;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(withNumberTree);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());
        PdfDictionary catalog = Assert.IsType<PdfDictionary>(PdfSyntax.FindCatalog(parsed.Map, parsed.TrailerRaw));
        PdfDictionary privateTree = Assert.IsType<PdfDictionary>(catalog.Items["PrivateTree"]);

        Assert.True(result.WasChanged);
        Assert.False(privateTree.Items.ContainsKey("Nums"));
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void IndirectStructureElementTypeCannotMasqueradeAsAFileSpecification() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            PdfDictionary candidateDictionary = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            candidateDictionary.Items["Type"] = AddObject(objects, new PdfName("StructElem"));
            var structureTree = new PdfDictionary();
            structureTree.Items["Type"] = new PdfName("StructTreeRoot");
            structureTree.Items["K"] = candidate;
            catalog.Items["StructTreeRoot"] = AddObject(objects, structureTree);
        });
    }

    [Fact]
    public void RequiredDecodeCachesPermanentFilterFailures() {
        var dictionary = new PdfDictionary();
        dictionary.Items["Filter"] = new PdfName("DCTDecode");
        var stream = new PdfStream(dictionary, new byte[] { 1, 2, 3 });
        var budget = new PdfDecodedStreamBudget(new PdfReadLimits { MaxDecodedStreamBytes = 64 });
        var objects = new Dictionary<int, PdfIndirectObject>();

        Assert.Throws<InvalidDataException>(() => budget.DecodeRequired(stream, objects, maximumRequestedBytes: 64));
        dictionary.Items.Remove("Filter");
        Assert.Throws<InvalidDataException>(() => budget.DecodeRequired(stream, objects, maximumRequestedBytes: 64));
        Assert.Equal(0, budget.UsedBytes);
    }

    [Fact]
    public void RequiredDecodeRetriesALimitFailureWhenTheCallerRaisesTheLimit() {
        byte[] expanded = Enumerable.Repeat((byte)'x', 64).ToArray();
        byte[] compressed;
        using (var output = new MemoryStream()) {
            using (var compressor = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
                compressor.Write(expanded, 0, expanded.Length);
            }
            compressed = output.ToArray();
        }
        var dictionary = new PdfDictionary();
        dictionary.Items["Filter"] = new PdfName("FlateDecode");
        var stream = new PdfStream(dictionary, compressed);
        var budget = new PdfDecodedStreamBudget(new PdfReadLimits {
            MaxDecodedStreamBytes = 128,
            MaxTotalDecodedStreamBytes = 128
        });
        var objects = new Dictionary<int, PdfIndirectObject>();

        Assert.Throws<PdfReadLimitException>(() => budget.DecodeRequired(stream, objects, maximumRequestedBytes: 16));
        byte[] decoded = budget.DecodeRequired(stream, objects, maximumRequestedBytes: 128);

        Assert.Equal(expanded, decoded);
        Assert.Equal(expanded.Length, budget.UsedBytes);
    }
}
