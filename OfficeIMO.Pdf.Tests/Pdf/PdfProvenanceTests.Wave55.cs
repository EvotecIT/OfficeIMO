using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ProvenanceExpandedByteLimitControlsTheAggregateParserBudget() {
        const long requested = 768L * 1024L * 1024L;

        PdfReadLimits effective = PdfReadLimits.Default.WithMaximumContainerEntries(100, requested);

        Assert.Equal(requested, effective.MaxTotalDecodedStreamBytes);
    }

    [Fact]
    public void ActiveArticleThreadGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            PdfIndirectObject pageObject = Assert.Single(
                objects.Values,
                item => item.Value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page");
            var bead = new PdfDictionary();
            bead.Items["P"] = new PdfReference(pageObject.ObjectNumber, pageObject.Generation);
            bead.Items["AF"] = ArrayWith(candidate);
            PdfReference beadReference = AddObject(objects, bead);
            bead.Items["N"] = beadReference;
            bead.Items["V"] = beadReference;
            var thread = new PdfDictionary();
            thread.Items["F"] = beadReference;
            var threads = new PdfArray();
            threads.Items.Add(AddObject(objects, thread));
            catalog.Items["Threads"] = threads;
        });
    }
}
