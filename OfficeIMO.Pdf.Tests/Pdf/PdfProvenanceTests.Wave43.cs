using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActiveAnnotationDestinationDictionaryCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var annotation = new PdfDictionary();
            annotation.Items["Type"] = new PdfName("Annot");
            annotation.Items["Subtype"] = new PdfName("Link");
            annotation.Items["Dest"] = candidate;
            page.Items["Annots"] = ArrayWith(AddObject(objects, annotation));
        });
    }

    [Theory]
    [InlineData("Lock")]
    [InlineData("SV")]
    public void ActiveSignatureFieldConstraintDictionaryCannotOwnProvenanceAssociations(string constraintKey) {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var field = new PdfDictionary();
            field.Items["FT"] = new PdfName("Sig");
            field.Items[constraintKey] = candidate;
            var acroForm = new PdfDictionary();
            acroForm.Items["Fields"] = ArrayWith(AddObject(objects, field));
            catalog.Items["AcroForm"] = AddObject(objects, acroForm);
        });
    }

    [Fact]
    public void RemovalPreservesMalformedUnrelatedNameTreeChildren() {
        byte[] original = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] malformed = PdfDocumentObjectGraphRewriter.Rewrite(original, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, names.Items["EmbeddedFiles"]));
            PdfArray entries = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, embeddedFiles.Items["Names"]));
            var targetEntries = new PdfArray();
            var retainedEntries = new PdfArray();
            for (int index = 0; index + 1 < entries.Items.Count; index += 2) {
                PdfStringObj name = Assert.IsType<PdfStringObj>(entries.Items[index]);
                PdfArray destination = name.Value.EndsWith(".c2pa", StringComparison.Ordinal) ? targetEntries : retainedEntries;
                destination.Items.Add(entries.Items[index]);
                destination.Items.Add(entries.Items[index + 1]);
            }
            PdfReference targetChild = AddObject(objects, DictionaryWithNames(targetEntries));
            PdfReference retainedChild = AddObject(objects, DictionaryWithNames(retainedEntries));
            PdfReference malformedChild = AddObject(objects, PdfNull.Instance);
            var kids = new PdfArray();
            kids.Items.Add(targetChild);
            kids.Items.Add(malformedChild);
            kids.Items.Add(retainedChild);
            embeddedFiles.Items.Remove("Names");
            embeddedFiles.Items["Kids"] = kids;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(malformed);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());
        PdfDictionary outputCatalog = Assert.IsType<PdfDictionary>(PdfSyntax.FindCatalog(parsed.Map, parsed.TrailerRaw));
        PdfDictionary outputNames = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(parsed.Map, outputCatalog.Items["Names"]));
        PdfDictionary outputEmbeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(parsed.Map, outputNames.Items["EmbeddedFiles"]));
        PdfArray outputKids = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(parsed.Map, outputEmbeddedFiles.Items["Kids"]));

        Assert.Contains(outputKids.Items, item => PdfObjectLookup.Resolve(parsed.Map, item) is PdfNull);
        Assert.Equal("keep.txt", Assert.Single(PdfAttachmentExtractor.ExtractAttachments(result.ToArray())).FileName);
    }

    private static PdfDictionary DictionaryWithNames(PdfArray names) {
        var dictionary = new PdfDictionary();
        dictionary.Items["Names"] = names;
        return dictionary;
    }
}
