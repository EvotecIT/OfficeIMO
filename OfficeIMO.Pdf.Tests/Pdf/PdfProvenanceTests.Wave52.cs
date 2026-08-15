using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void CatalogLegalDictionariesCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var legal = new PdfDictionary();
            legal.Items["AF"] = ArrayWith(candidate);
            catalog.Items["Legal"] = AddObject(objects, legal);
        });
    }
}
