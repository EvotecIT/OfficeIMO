using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfProvenanceTests {
    [Fact]
    public void ActiveThreeDimensionalAnnotationViewGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var annotation = new PdfDictionary();
            annotation.Items["Type"] = new PdfName("Annot");
            annotation.Items["Subtype"] = new PdfName("3D");
            annotation.Items["3DV"] = candidate;
            page.Items["Annots"] = ArrayWith(AddObject(objects, annotation));
        });
    }

    [Fact]
    public void ActiveIccProfileAlternateGraphsCannotOwnProvenanceAssociations() {
        AssertStructuralOwnerRejected((objects, _, candidate) => {
            PdfDictionary page = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            var profileDictionary = new PdfDictionary();
            profileDictionary.Items["N"] = new PdfNumber(3);
            profileDictionary.Items["Alternate"] = candidate;
            PdfReference profile = AddObject(objects, new PdfStream(profileDictionary, Array.Empty<byte>()));
            var iccBased = new PdfArray();
            iccBased.Items.Add(new PdfName("ICCBased"));
            iccBased.Items.Add(profile);
            var colorSpaces = new PdfDictionary();
            colorSpaces.Items["CS1"] = iccBased;
            var resources = new PdfDictionary();
            resources.Items["ColorSpace"] = colorSpaces;
            page.Items["Resources"] = resources;
        });
    }
}
