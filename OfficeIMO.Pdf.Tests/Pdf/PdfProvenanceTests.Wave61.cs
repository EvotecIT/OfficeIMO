using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActiveShadingFunctionAndColorSpaceGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var shading = new PdfDictionary();
            shading.Items["ShadingType"] = new PdfNumber(2);
            shading.Items["Function"] = candidate;
            shading.Items["ColorSpace"] = candidate;
            catalog.Items["PrivateShading"] = AddObject(objects, new PdfStream(shading, Array.Empty<byte>()));
        });
    }

    [Fact]
    public void ActiveRichMediaAnnotationGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(
                objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var annotation = new PdfDictionary();
            annotation.Items["Type"] = new PdfName("Annot");
            annotation.Items["Subtype"] = new PdfName("RichMedia");
            annotation.Items["RichMediaContent"] = candidate;
            annotation.Items["RichMediaSettings"] = candidate;
            var annotations = new PdfArray();
            annotations.Items.Add(AddObject(objects, annotation));
            page.Items["Annots"] = annotations;
        });
    }

    [Fact]
    public void ParsedStreamCacheIsReusedDuringProvenanceAttachmentExtraction() {
        byte[] pdf = CreatePdfWithCandidateAndRetainedAttachment();
        int manifestLength = CreateManifestStore().Length;
        byte[] shared = PdfDocumentObjectGraphRewriter.Rewrite(pdf, null, null, (objects, security) => {
            PdfDictionary catalog = Assert.IsType<PdfDictionary>(objects[security.RootObjectNumber!.Value].Value);
            PdfArray associated = Assert.IsType<PdfArray>(PdfObjectLookup.Resolve(objects, catalog.Items["AF"]));
            PdfReference fileSpecReference = FindFileSpecReference(objects, associated, "content-credential.c2pa");
            PdfDictionary fileSpec = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, fileSpecReference));
            PdfDictionary embeddedFiles = Assert.IsType<PdfDictionary>(PdfObjectLookup.Resolve(objects, fileSpec.Items["EF"]));
            catalog.Items["Metadata"] = embeddedFiles.Items["F"];
            return security.InfoObjectNumber;
        });
        var options = new OfficeProvenanceOptions {
            MaxManifestBytes = manifestLength * 2,
            MaxExpandedContainerBytes = manifestLength * 2L
        };
        var readOptions = new PdfLoadOptions {
            Limits = new PdfReadLimits {
                MaxDecodedStreamBytes = manifestLength * 2,
                MaxTotalDecodedStreamBytes = manifestLength + 64L,
                MaxTotalAttachmentBytes = manifestLength * 2L
            }
        };

        OfficeProvenanceReport report = PdfProvenance.Inspect(shared, options, readOptions);

        Assert.NotEmpty(report.Evidence);
    }
}
