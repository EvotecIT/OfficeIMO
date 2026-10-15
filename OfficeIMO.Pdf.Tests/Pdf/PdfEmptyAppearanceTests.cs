using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

/// <summary>
/// Appearances with an empty bounding box paint nothing. macOS PDFKit (and so Preview) writes them for the off state of
/// every form widget, so a page that carries them still renders, reports diagnostics and flattens.
/// </summary>
public sealed class PdfEmptyAppearanceTests {
    [Theory]
    [InlineData(90)]
    [InlineData(180)]
    [InlineData(270)]
    public void RotationNormalizationPreservesEmptyAppearancesAndWidgetStates(int rotation) {
        PdfDocument source = PdfDocument.Load(BuildPdf(widgetOffAppearanceBox: "[0 0 0 0]",
            squareAppearanceBox: "[0 0 0 0]")).Pages.Rotate(rotation);
        PdfPageRenderResult before = Assert.Single(PdfPageImageRenderer.RenderPages(source.ToBytes(),
            options: new PdfPageRenderOptions { ContinueOnError = false }));

        PdfDocument normalized = source.Pages.NormalizeRotation();
        PdfPageRenderResult after = Assert.Single(PdfPageImageRenderer.RenderPages(normalized.ToBytes(),
            options: new PdfPageRenderOptions { ContinueOnError = false }));

        Assert.Equal(0, normalized.Inspect().Pages[0].RotationDegrees);
        Assert.Single(normalized.Annotations.GetForEditing());
        Assert.Single(normalized.Inspect().FormFields);
        Assert.True(before.Succeeded);
        Assert.True(after.Succeeded);
        Assert.NotNull(before.Bytes);
        Assert.Equal(before.Bytes, after.Bytes);
    }

    [Fact]
    public void RenderingSkipsAnEmptyWidgetAppearanceAndDrawsTheRestOfThePage() {
        byte[] pdf = BuildPdf(widgetOffAppearanceBox: "[0 0 0 0]", squareAppearanceBox: "[0 0 20 20]");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf,
            options: new PdfPageRenderOptions { ContinueOnError = false }));

        Assert.True(result.Succeeded);
        Assert.NotNull(result.Bytes);
    }

    [Fact]
    public void FieldFlatteningDrawsNothingForAnEmptyWidgetAppearance() {
        byte[] flattened = PdfFormFiller.FlattenFields(BuildPdf(widgetOffAppearanceBox: "[0 0 0 0]", squareAppearanceBox: "[0 0 20 20]"));

        Assert.DoesNotContain(PdfDocument.Load(flattened).Annotations.GetForEditing(), annotation => annotation.Subtype == "Widget");
        Assert.True(Assert.Single(PdfPageImageRenderer.RenderPages(flattened)).Succeeded);
    }

    [Fact]
    public void AnnotationFlatteningRemovesAnAnnotationWhoseAppearanceIsEmpty() {
        // The annotation flattener leaves documents with form fields alone, so this page has only the square.
        PdfDocument document = PdfDocument.Load(BuildPdf(widgetOffAppearanceBox: null, squareAppearanceBox: "[0 0 0 0]"));

        PdfAnnotationEditResult result = document.Annotations.Flatten();

        Assert.DoesNotContain(PdfDocument.Load(result.Bytes).Annotations.GetForEditing(), annotation => annotation.Subtype == "Square");
    }

    // One page with a square annotation and, when widgetOffAppearanceBox is given, a checkbox widget in its off state.
    private static byte[] BuildPdf(string? widgetOffAppearanceBox, string squareAppearanceBox) {
        bool form = widgetOffAppearanceBox != null;
        string[] objects = {
            form ? "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [4 0 R] >> >>" : "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << >> /Contents 8 0 R /Annots [" + (form ? "4 0 R " : "") + "7 0 R] >>",
            form ? "<< /Type /Annot /Subtype /Widget /FT /Btn /T (agree) /V /Off /AS /Off /Rect [20 20 40 40] /P 3 0 R /AP << /N << /Yes 5 0 R /Off 6 0 R >> >> >>" : "<< >>",
            Form("[0 0 20 20]", "0 0 1 rg 2 2 16 16 re f"),
            Form(widgetOffAppearanceBox ?? "[0 0 20 20]", ""),
            "<< /Type /Annot /Subtype /Square /Rect [100 100 140 140] /C [1 0 0] /AP << /N 9 0 R >> >>",
            Stream("<< /Length {0} >>", "0 g 10 160 50 20 re f"),
            Form(squareAppearanceBox, "1 0 0 rg 0 0 20 20 re f"),
        };
        return Assemble(objects);
    }

    private static string Form(string box, string content) =>
        Stream("<< /Type /XObject /Subtype /Form /BBox " + box + " /Resources << >> /Length {0} >>", content);

    private static string Stream(string dictionary, string content) =>
        string.Format(System.Globalization.CultureInfo.InvariantCulture, dictionary, content.Length) + "\nstream\n" + content + "\nendstream";

    private static byte[] Assemble(IReadOnlyList<string> objects) {
        var pdf = new StringBuilder("%PDF-1.7\n");
        var offsets = new List<int>();
        for (int index = 0; index < objects.Count; index++) {
            offsets.Add(pdf.Length);
            pdf.Append(index + 1).Append(" 0 obj\n").Append(objects[index]).Append("\nendobj\n");
        }
        int xref = pdf.Length;
        pdf.Append("xref\n0 ").Append(objects.Count + 1).Append("\n0000000000 65535 f \n");
        foreach (int offset in offsets) pdf.Append(offset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture)).Append(" 00000 n \n");
        pdf.Append("trailer << /Root 1 0 R /Size ").Append(objects.Count + 1).Append(" >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }
}
