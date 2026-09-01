using OfficeIMO.Pdf;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActiveFormPieceInfoGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var form = new PdfDictionary();
            form.Items["Type"] = new PdfName("XObject");
            form.Items["Subtype"] = new PdfName("Form");
            form.Items["PieceInfo"] = candidate;
            catalog.Items["PrivateForm"] = AddObject(objects, new PdfStream(form, Array.Empty<byte>()));
        });
    }

    [Fact]
    public void ExternalStreamFileSpecificationsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var stream = new PdfDictionary();
            stream.Items["F"] = candidate;
            catalog.Items["PrivateExternalStream"] = AddObject(objects, new PdfStream(stream, Array.Empty<byte>()));
        });
    }

    [Fact]
    public void SynthesizedReadOptionsHonorRaisedProvenanceAssetLimit() {
        const long raisedLimit = 600L * 1024L * 1024L;
        var limits = new OfficeProvenanceOptions { MaxAssetBytes = raisedLimit };

        PdfLoadOptions readOptions = PdfProvenance.CreateReadOptionsForInspection(limits, readOptions: null);

        Assert.Equal(raisedLimit, readOptions.Limits.MaxInputBytes);
    }

    [Fact]
    public void ExplicitLowerPdfInputLimitRemainsAuthoritative() {
        const long raisedLimit = 600L * 1024L * 1024L;
        const long explicitLimit = 384L * 1024L * 1024L;
        var limits = new OfficeProvenanceOptions { MaxAssetBytes = raisedLimit };
        var requested = new PdfLoadOptions { Limits = new PdfReadLimits { MaxInputBytes = explicitLimit } };

        PdfLoadOptions readOptions = PdfProvenance.CreateReadOptionsForInspection(limits, requested);

        Assert.Equal(explicitLimit, readOptions.Limits.MaxInputBytes);
    }
}
