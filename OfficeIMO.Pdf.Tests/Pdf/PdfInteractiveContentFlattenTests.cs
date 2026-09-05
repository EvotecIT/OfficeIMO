using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFormFillerTests {
    [Fact]
    public void FlattenInteractiveContent_FlattensFormsAndVisualAnnotationsInOneOperation() {
        byte[] form = PdfFormFiller.FillFields(BuildTextWidgetFormPdf(), new Dictionary<string, string> {
            ["Name"] = "Flatten both"
        });
        byte[] annotated = PdfDocument.Load(form).Annotations.Add(new PdfAnnotationCreateOptions {
            Subtype = "FreeText",
            Contents = "Flatten annotation",
            Rectangle = new[] { 150D, 80D, 260D, 110D },
            InteriorColor = new[] { 0.9D, 0.95D, 1D },
            Opacity = 0.75D,
            BorderWidth = 2D,
            BorderStyle = PdfAnnotationBorderStyle.Dashed,
            BorderDashPattern = new[] { 4D, 2D }
        }).Bytes;

        PdfInteractiveContentFlattenResult result = PdfDocument.Load(annotated).FlattenInteractiveContent();
        PdfDocumentInfo info = result.ToDocument().Inspect();

        Assert.True(result.Applied);
        Assert.Equal(1, result.FlattenedFormFieldCount);
        Assert.Equal(1, result.FlattenedAnnotationCount);
        Assert.False(info.HasForms);
        Assert.Empty(info.GetAnnotationsBySubtype("FreeText"));
        Assert.Contains("Flatten both", PdfTextExtractor.ExtractAllText(result.Bytes), StringComparison.Ordinal);
    }
}
