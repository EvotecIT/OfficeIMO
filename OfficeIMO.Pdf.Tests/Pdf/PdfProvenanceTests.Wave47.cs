using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActiveStructureTreeClassMapCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var classMap = new PdfDictionary();
            classMap.Items["Artifact"] = candidate;
            var structureTree = new PdfDictionary();
            structureTree.Items["Type"] = new PdfName("StructTreeRoot");
            structureTree.Items["ClassMap"] = AddObject(objects, classMap);
            catalog.Items["StructTreeRoot"] = AddObject(objects, structureTree);
        });
    }

    [Fact]
    public void AcroFormOnlyWidgetAppearanceCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, catalog, candidate) => {
            var appearance = new PdfDictionary();
            appearance.Items["N"] = candidate;
            var widget = new PdfDictionary();
            widget.Items["Type"] = new PdfName("Annot");
            widget.Items["Subtype"] = new PdfName("Widget");
            widget.Items["AP"] = AddObject(objects, appearance);
            var acroForm = new PdfDictionary();
            acroForm.Items["Fields"] = ArrayWith(AddObject(objects, widget));
            catalog.Items["AcroForm"] = AddObject(objects, acroForm);
        });
    }
}
