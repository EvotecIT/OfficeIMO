using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData("2 Tc (spaced) Tj", false)]
    [InlineData("4 Tw (spaced text) Tj", false)]
    [InlineData("2 Tc (spaced) Tj", true)]
    [InlineData("4 Tw (spaced text) Tj", true)]
    public void RenderPage_DoesNotScaleAggregateAdvanceForExplicitTextSpacing(string textOperation, bool clipped) {
        string clipOperation = clipped ? "20 90 50 30 re W n " : string.Empty;
        string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf(
            clipOperation + "BT /F1 20 Tf 20 100 Td " + textOperation + " ET",
            "<< /Font << /F1 5 0 R >> >>",
            font);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeDrawingText text = clipped
            ? Assert.Single(Assert.IsType<OfficeDrawingGroup>(Assert.Single(drawing.Elements)).Drawing.Elements.OfType<OfficeDrawingText>())
            : Assert.Single(drawing.Elements.OfType<OfficeDrawingText>());

        Assert.Null(text.TextAdvanceWidth);
        Assert.Equal(OfficeTextOverflowBehavior.Ellipsis, text.OverflowBehavior);
    }
}
