using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActiveSubtypeLessStructureTreeRootCannotMasqueradeAsFileSpecification() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] structuralCarrier = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            fileSpecification.Items.Remove("Type");
            catalog.Items["StructTreeRoot"] = candidate;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(structuralCarrier);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(structuralCarrier);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(structuralCarrier, result.ToArray());
    }

    [Fact]
    public void AttachmentTraversalDoesNotChargePrevalidatedNameAndAssociatedFileItemsAgain() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        var parsed = PdfSyntax.ParseObjects(pdf);
        int structuralEntries = CountStructuralEntries(parsed.Map);
        var options = new OfficeProvenanceOptions { MaxContainerEntries = structuralEntries };

        OfficeProvenanceReport report = PdfProvenance.Inspect(pdf, options);

        Assert.True(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    private static int CountStructuralEntries(Dictionary<int, PdfIndirectObject> objects) {
        var visited = new HashSet<PdfObject>();
        var pending = new Stack<PdfObject>(objects.Values.Select(static item => item.Value));
        int count = 0;
        while (pending.Count > 0) {
            PdfObject value = pending.Pop();
            if (!visited.Add(value)) continue;
            PdfDictionary? dictionary = value is PdfStream stream ? stream.Dictionary : value as PdfDictionary;
            if (dictionary != null) {
                count = checked(count + 1 + dictionary.Items.Count);
                foreach (PdfObject child in dictionary.Items.Values) {
                    if (child is not PdfReference) pending.Push(child);
                }
            } else if (value is PdfArray array) {
                count = checked(count + 1 + array.Items.Count);
                foreach (PdfObject child in array.Items) {
                    if (child is not PdfReference) pending.Push(child);
                }
            }
        }
        return count;
    }
}
