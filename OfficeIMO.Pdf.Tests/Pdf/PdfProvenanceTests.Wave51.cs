using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActiveSignatureValueGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var byteRange = new PdfArray();
            byteRange.Items.Add(new PdfNumber(0));
            byteRange.Items.Add(new PdfNumber(1));
            byteRange.Items.Add(new PdfNumber(2));
            byteRange.Items.Add(new PdfNumber(3));
            var signatureValue = new PdfDictionary();
            signatureValue.Items["ByteRange"] = byteRange;
            signatureValue.Items["AF"] = ArrayWith(candidate);
            var field = new PdfDictionary();
            field.Items["FT"] = new PdfName("Sig");
            field.Items["V"] = AddObject(objects, signatureValue);
            var acroForm = new PdfDictionary();
            acroForm.Items["Fields"] = ArrayWith(AddObject(objects, field));
            catalog.Items["AcroForm"] = AddObject(objects, acroForm);
        });
    }
}
