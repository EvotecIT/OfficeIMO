using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void SharedEmbeddedFileDictionaryChargesItsVariantsOnce() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] repeatedVariants = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associatedFiles = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associatedFiles, "content-credential.c2pa");
            PdfDictionary fileSpec = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, fileSpec.Items["EF"]));
            PdfObject stream = embeddedFiles.Items["UF"];
            for (int index = 0; index < 256; index++) embeddedFiles.Items["Variant" + index] = stream;
            return security.InfoObjectNumber;
        });

        Assert.Throws<InvalidDataException>(() => PdfProvenance.Inspect(
            repeatedVariants,
            new OfficeProvenanceOptions { MaxContainerEntries = 128 }));
    }

    [Fact]
    public void SharedEmbeddedFileParametersCanAlsoBeAnIndependentInformationResource() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] sharedParameters = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associatedFiles = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associatedFiles, "content-credential.c2pa");
            PdfDictionary fileSpec = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, fileSpec.Items["EF"]));
            PdfReference streamReference = Assert.IsType<PdfReference>(embeddedFiles.Items["UF"]);
            PdfStream stream = Assert.IsType<PdfStream>(objects[streamReference.ObjectNumber].Value);
            var parameters = new PdfDictionary();
            var association = new PdfArray();
            association.Items.Add(candidate);
            parameters.Items["AF"] = association;
            int parametersObjectNumber = objects.Keys.Max() + 1;
            objects[parametersObjectNumber] = new PdfIndirectObject(parametersObjectNumber, 0, parameters);
            stream.Dictionary.Items["Params"] = new PdfReference(parametersObjectNumber, 0);
            catalog.Items["Wave24Resource"] = new PdfReference(parametersObjectNumber, 0);
            catalog.Items.Remove("AF");
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(sharedParameters);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(sharedParameters);

        Assert.True(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void TypedPageWithStrayKidsStillContributesAndLosesItsFileAttachment() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        int annotationObjectNumber = 0;
        byte[] pageWithKids = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associatedFiles = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associatedFiles, "content-credential.c2pa");
            PdfDictionary names = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, catalog.Items["Names"]));
            names.Items.Remove("EmbeddedFiles");
            PdfDictionary page = Assert.IsType<PdfDictionary>(objects.Values.Select(item => item.Value)
                .First(value => value is PdfDictionary dictionary && dictionary.Get<PdfName>("Type")?.Name == "Page"));
            page.Items["Kids"] = new PdfArray();
            annotationObjectNumber = objects.Keys.Max() + 1;
            var attachment = new PdfDictionary();
            attachment.Items["Type"] = new PdfName("Annot");
            attachment.Items["Subtype"] = new PdfName("FileAttachment");
            attachment.Items["FS"] = candidate;
            objects[annotationObjectNumber] = new PdfIndirectObject(annotationObjectNumber, 0, attachment);
            var annotations = new PdfArray();
            annotations.Items.Add(new PdfReference(annotationObjectNumber, 0));
            page.Items["Annots"] = annotations;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(pageWithKids);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());

        Assert.True(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.True(result.WasChanged);
        Assert.DoesNotContain(annotationObjectNumber, parsed.Map.Keys);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void RemovalDeletesReferencedManifestStreamAsWellAsItsFileSpec() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] alias = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associatedFiles = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associatedFiles, "content-credential.c2pa");
            PdfDictionary fileSpec = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, fileSpec.Items["EF"]));
            PdfReference streamReference = Assert.IsType<PdfReference>(embeddedFiles.Items["UF"]);
            catalog.Items["Wave24StreamAlias"] = streamReference;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(alias);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());
        PdfDictionary retainedCatalog = Assert.IsType<PdfDictionary>(PdfSyntax.FindCatalog(parsed.Map, parsed.TrailerRaw));

        Assert.True(result.WasChanged);
        Assert.DoesNotContain(parsed.Map.Values, item => item.Value is PdfStream stream &&
            stream.Dictionary.Get<PdfName>("Subtype")?.Name.Contains("c2pa", StringComparison.OrdinalIgnoreCase) == true);
        Assert.False(retainedCatalog.Items.ContainsKey("Wave24StreamAlias"));
        Assert.Empty(result.After.Evidence);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RemovalRejectsManifestStreamSharedWithRetainedAttachment(bool malformedRetainedFileSpec) {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] sharedStream = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associatedFiles = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associatedFiles, "content-credential.c2pa");
            PdfReference retained = FindFileSpecReference(objects, associatedFiles, "keep.txt");
            PdfDictionary candidateSpec = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            PdfDictionary retainedSpec = Assert.IsType<PdfDictionary>(objects[retained.ObjectNumber].Value);
            if (malformedRetainedFileSpec) {
                retainedSpec.Items.Remove("Type");
                retainedSpec.Items.Remove("F");
                retainedSpec.Items.Remove("UF");
            }
            PdfDictionary candidateEmbeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, candidateSpec.Items["EF"]));
            PdfDictionary retainedEmbeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, retainedSpec.Items["EF"]));
            PdfObject manifestStream = candidateEmbeddedFiles.Items["UF"];
            retainedEmbeddedFiles.Items["UF"] = manifestStream;
            retainedEmbeddedFiles.Items["F"] = manifestStream;
            return security.InfoObjectNumber;
        });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => PdfProvenance.Remove(sharedStream));

        Assert.Contains("shares the selected provenance stream", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RemovalScrubsStaleGenerationReferencesToDeletedObjects() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] staleReference = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associatedFiles = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associatedFiles, "content-credential.c2pa");
            catalog.Items["Wave24Stale"] = candidate;
            return security.InfoObjectNumber;
        });
        const string marker = "/Wave24Stale ";
        int offset = Encoding.ASCII.GetString(staleReference).IndexOf(marker, StringComparison.Ordinal);
        Assert.True(offset >= 0);
        offset += marker.Length;
        while (staleReference[offset] != (byte)' ') offset++;
        Assert.Equal((byte)'0', staleReference[++offset]);
        staleReference[offset] = (byte)'1';

        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(staleReference);
        var parsed = PdfSyntax.ParseObjects(result.ToArray());
        PdfDictionary catalog = Assert.IsType<PdfDictionary>(PdfSyntax.FindCatalog(parsed.Map, parsed.TrailerRaw));

        Assert.True(result.WasChanged);
        Assert.False(catalog.Items.ContainsKey("Wave24Stale"));
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void ActiveLegacyDestinationDictionaryCannotMasqueradeAsFileSpecification() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        byte[] structuralCarrier = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associations = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference candidate = FindFileSpecReference(objects, associations, "content-credential.c2pa");
            PdfDictionary fileSpecification = Assert.IsType<PdfDictionary>(objects[candidate.ObjectNumber].Value);
            fileSpecification.Items.Remove("Type");
            var destinations = new PdfDictionary();
            destinations.Items["credential"] = candidate;
            catalog.Items["Dests"] = destinations;
            return security.InfoObjectNumber;
        });

        OfficeProvenanceReport report = PdfProvenance.Inspect(structuralCarrier);
        OfficeProvenanceRemovalResult result = PdfProvenance.Remove(structuralCarrier);

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(structuralCarrier, result.ToArray());
    }
}
