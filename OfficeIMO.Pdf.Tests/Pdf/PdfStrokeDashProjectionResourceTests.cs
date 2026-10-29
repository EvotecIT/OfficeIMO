using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfExternalDocumentCompatibilityTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void DashProjectionBudgetSpansRepeatedPaintAndFormInvocations(bool nested, bool scaled) {
        var page = CreateDashProjectionPage(nested, scaled, stroke: true);

        var error = Assert.Throws<PdfReadLimitException>(() => page.ToDrawing());

        Assert.Equal(PdfReadLimitKind.ContentOperands, error.Kind);
        Assert.Equal(64, error.Limit);
        Assert.True(error.Actual > error.Limit);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FillOnlyPaintDoesNotProjectUnusedDashes(bool nested) {
        var drawing = CreateDashProjectionPage(nested, scaled: true, stroke: false).ToDrawing();

        var shapes = nested ? drawing.Elements.OfType<OfficeDrawingGroup>().SelectMany(group => group.Drawing.Shapes).ToArray() : drawing.Shapes.ToArray();
        Assert.Equal(6, shapes.Length);
        Assert.All(shapes, shape => Assert.Empty(shape.Shape.StrokeDashArray));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OrdinaryDashPatternsRetainTheirScaledLengths(bool nested) {
        var drawing = CreateDashProjectionPage(nested, scaled: true, stroke: true, entries: 2).ToDrawing();
        var shapes = nested ? drawing.Elements.OfType<OfficeDrawingGroup>().SelectMany(group => group.Drawing.Shapes).ToArray() : drawing.Shapes.ToArray();

        Assert.Equal(6, shapes.Length);
        Assert.All(shapes, shape => Assert.Equal(new[] { 2D, 2D }, shape.Shape.StrokeDashArray));
    }

    private static PdfReadPage CreateDashProjectionPage(bool nested, bool scaled, bool stroke, int entries = 16) {
        string state = (scaled ? "2 0 0 2 0 0 cm " : "") + "[" + string.Join(" ", Enumerable.Repeat("1", entries)) + "] 0 d ";
        string paint = "0 0 2 2 re " + (stroke ? "S " : "f ");
        string content = state + string.Concat(Enumerable.Repeat(nested ? "/G Do " : paint, 6));
        return PdfReadDocument.Open(BuildPdf(new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 20 20] /Resources << /XObject << /G 5 0 R >> >> /Contents 4 0 R >>\nendobj",
            BuildStreamObject(4, Encoding.ASCII.GetBytes(content)),
            BuildStreamObject(5, Encoding.ASCII.GetBytes(paint), "/Type /XObject /Subtype /Form /BBox [0 0 20 20] /Resources << >>")
        }, rootObjectNumber: 1), new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxContentOperands = 64 }
        }).Pages[0];
    }
}
