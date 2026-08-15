using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void StructTreeParentTreeDictionariesCannotOwnProvenanceAssociations() {
        byte[] pdf = RewriteCandidateAssociation((objects, catalog, candidate) => {
            var owner = new PdfDictionary();
            owner.Items["AF"] = ArrayWith(candidate);
            PdfReference ownerReference = AddObject(objects, owner);
            var parentTree = new PdfDictionary();
            parentTree.Items["Nums"] = new PdfArray();
            ((PdfArray)parentTree.Items["Nums"]).Items.Add(new PdfNumber(0));
            ((PdfArray)parentTree.Items["Nums"]).Items.Add(ownerReference);
            var structureTree = new PdfDictionary();
            structureTree.Items["Type"] = new PdfName("StructTreeRoot");
            structureTree.Items["ParentTree"] = AddObject(objects, parentTree);
            catalog.Items["StructTreeRoot"] = structureTree;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(pdf);

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void NamedDestinationValueDictionariesCannotOwnProvenanceAssociations() {
        byte[] pdf = RewriteCandidateAssociation((objects, catalog, candidate) => {
            var destination = new PdfDictionary();
            destination.Items["AF"] = ArrayWith(candidate);
            var leaf = new PdfDictionary();
            var names = new PdfArray();
            names.Items.Add(new PdfStringObj("destination"));
            names.Items.Add(AddObject(objects, destination));
            leaf.Items["Names"] = names;
            PdfDictionary catalogNames = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            catalogNames.Items["Dests"] = AddObject(objects, leaf);
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(pdf);

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void FileSpecificationDescendantsCannotOwnProvenanceAssociations() {
        byte[] pdf = RewriteCandidateAssociation((objects, _, candidate) => {
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            var descendant = new PdfDictionary();
            descendant.Items["AF"] = ArrayWith(candidate);
            fileSpecification.Items["PrivateData"] = descendant;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(pdf);

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void ActiveExternalFileActionsPreventTheirFileSpecificationRemoval() {
        byte[] original = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] pdf = PdfDocumentObjectGraphRewriter.Rewrite(original, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            var action = new PdfDictionary();
            action.Items["S"] = new PdfName("GoToR");
            action.Items["F"] = candidate;
            action.Items["D"] = new PdfStringObj("remote");
            var outlineItem = new PdfDictionary();
            outlineItem.Items["Title"] = new PdfStringObj("Remote");
            outlineItem.Items["A"] = AddObject(objects, action);
            PdfReference outlineItemReference = AddObject(objects, outlineItem);
            var outlines = new PdfDictionary();
            outlines.Items["Type"] = new PdfName("Outlines");
            outlines.Items["First"] = outlineItemReference;
            outlines.Items["Last"] = outlineItemReference;
            catalog.Items["Outlines"] = AddObject(objects, outlines);
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(pdf);

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void ParserSharesTheExpandedByteBudgetAcrossObjectStreams() {
        byte[] pdf = CreatePdfWithTwoObjectStreams(70);
        var options = new OfficeProvenanceOptions { MaxExpandedContainerBytes = 100 };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => PdfProvenance.Inspect(pdf, options));

        Assert.Equal(PdfReadLimitKind.TotalDecodedStreamBytes, exception.Kind);
    }

    private static byte[] RewriteCandidateAssociation(Action<Dictionary<int, PdfIndirectObject>, PdfDictionary, PdfReference> mutate) {
        byte[] original = CreatePdfWithCandidateAndRetainedAttachment();
        return PdfDocumentObjectGraphRewriter.Rewrite(original, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            catalog.Items.Remove("AF");
            mutate(objects, catalog, candidate);
            return security.InfoObjectNumber;
        });
    }

    private static PdfArray ArrayWith(PdfObject value) {
        var array = new PdfArray();
        array.Items.Add(value);
        return array;
    }

    private static PdfReference AddObject(Dictionary<int, PdfIndirectObject> objects, PdfObject value) {
        int number = objects.Keys.Max() + 1;
        objects[number] = new PdfIndirectObject(number, 0, value);
        return new PdfReference(number, 0);
    }

    private static byte[] CreatePdfWithTwoObjectStreams(int paddingCharacters) {
        string first = "20 0 << /Pad (" + new string('a', paddingCharacters) + ") >>";
        string second = "21 0 << /Pad (" + new string('b', paddingCharacters) + ") >>";
        var text = new StringBuilder();
        text.Append("%PDF-1.7\n")
            .Append("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n")
            .Append("2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n")
            .Append("3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 10 10] >>\nendobj\n");
        AppendObjectStream(text, 5, first);
        AppendObjectStream(text, 6, second);
        text.Append("trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return Encoding.ASCII.GetBytes(text.ToString());
    }

    private static void AppendObjectStream(StringBuilder text, int objectNumber, string payload) {
        int first = payload.IndexOf("<<", StringComparison.Ordinal);
        text.Append(objectNumber).Append(" 0 obj\n<< /Type /ObjStm /N 1 /First ")
            .Append(first).Append(" /Length ").Append(Encoding.ASCII.GetByteCount(payload))
            .Append(" >>\nstream\n").Append(payload).Append("\nendstream\nendobj\n");
    }
}
