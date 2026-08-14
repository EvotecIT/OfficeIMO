using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Theory]
    [InlineData("MK")]
    [InlineData("BS")]
    public void AcroFormOnlyWidgetStyleGraphsCannotOwnProvenanceAssociations(string styleKey) {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var style = new PdfDictionary();
            style.Items["Value"] = candidate;
            var widget = new PdfDictionary();
            widget.Items["Type"] = new PdfName("Annot");
            widget.Items["Subtype"] = new PdfName("Widget");
            widget.Items[styleKey] = AddObject(objects, style);
            var acroForm = new PdfDictionary();
            acroForm.Items["Fields"] = ArrayWith(AddObject(objects, widget));
            catalog.Items["AcroForm"] = AddObject(objects, acroForm);
        });
    }

    [Fact]
    public void ActiveStructureTreeIdMapCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var names = new PdfArray();
            names.Items.Add(new PdfStringObj("node"));
            names.Items.Add(candidate);
            var idTree = new PdfDictionary();
            idTree.Items["Names"] = names;
            var structureTree = new PdfDictionary();
            structureTree.Items["Type"] = new PdfName("StructTreeRoot");
            structureTree.Items["IDTree"] = AddObject(objects, idTree);
            catalog.Items["StructTreeRoot"] = AddObject(objects, structureTree);
        });
    }

}
