using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActiveFormOpiGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var form = new PdfDictionary();
            form.Items["Subtype"] = new PdfName("Form");
            form.Items["OPI"] = candidate;
            catalog.Items["PrivateForm"] = AddObject(objects, new PdfStream(form, Array.Empty<byte>()));
        });
    }
}
