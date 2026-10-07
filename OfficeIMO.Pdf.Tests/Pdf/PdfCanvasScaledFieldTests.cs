using System.Linq;
using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfCanvasScaledFieldTests {
    [Theory]
    [InlineData(0.5)]
    [InlineData(2.0)]
    public void OpaqueUniformScalePreservesFieldValueAndPhysicalWidgetBounds(double scale) {
        byte[] pdf = PdfDocument.Create().Canvas(canvas => canvas.Effect(
            OfficeTransform.Scale(scale, scale), 1D,
            content => content.TextField("name", "Ada", 10D, 20D, 80D, 20D))).ToBytes();
        PdfFormField field = Assert.Single(PdfInspector.Inspect(pdf).FormFields);
        Assert.Equal("Ada", field.Value);
        var widget = Assert.Single(field.Widgets);
        Assert.Equal(10D * scale, widget.X1, 4);
        Assert.Equal(80D * scale, widget.X2 - widget.X1, 4);
        Assert.Equal(20D * scale, widget.Y2 - widget.Y1, 4);
    }
    [Theory]
    [InlineData(0.5)]
    [InlineData(2.0)]
    public void NestedUniformScalePreservesAuthoredAppearanceAndScaledEditingStyle(double scale) {
        var style = new PdfFormFieldStyle { BorderWidth = 2D, CornerRadius = 3D };
        byte[] pdf = PdfDocument.Create().Canvas(canvas => canvas.Effect(
            OfficeTransform.Scale(scale, scale), 1D, outer => outer.Effect(
                OfficeTransform.Scale(scale, scale), 1D,
                inner => inner.TextField("name", "Ada", 10D, 20D, 80D, 20D, 12D, style)))).ToBytes();
        PdfFormField field = Assert.Single(PdfInspector.Inspect(pdf).FormFields);
        string totalFontSize = (12D * scale * scale).ToString("0.######", CultureInfo.InvariantCulture);
        Assert.Contains(" " + totalFontSize + " Tf", field.DefaultAppearance);
        string source = Encoding.ASCII.GetString(pdf);
        Assert.Contains("/BBox [0 0 80 20]", source);
        Assert.Equal(2D, style.BorderWidth);
        Assert.Equal(3D, style.CornerRadius);
        var widget = Assert.Single(field.Widgets);
        Assert.Equal(80D * scale * scale, widget.X2 - widget.X1, 4);
    }
}
