using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfGraphicsGradientResourceTests {
    [Theory]
    [InlineData(false, PdfObjectSerializationMode.Buffered)]
    [InlineData(true, PdfObjectSerializationMode.Buffered)]
    [InlineData(false, PdfObjectSerializationMode.ForwardOnly)]
    [InlineData(true, PdfObjectSerializationMode.ForwardOnly)]
    public void DifferentCoordinateShadingsShareColorAndAlphaFunctions(bool radial, PdfObjectSerializationMode mode) {
        var document = PdfDocument.Create(new PdfOptions { PageWidth = 220D, PageHeight = 160D,
            MarginLeft = 0D, MarginRight = 0D, MarginTop = 0D, MarginBottom = 0D,
            FileVersion = PdfFileVersion.Pdf17, ObjectSerializationMode = mode });
        var start = OfficeColor.FromRgba(255, 0, 0, 64);
        var end = OfficeColor.FromRgba(0, 0, 255, 192);
        for (int i = 0; i < 2; i++) {
            var shape = OfficeShape.Rectangle(80D + i * 20D, 40D); shape.StrokeWidth = 0D;
            if (radial) shape.FillRadialGradient = new OfficeRadialGradient(0.5D, 0.5D, 0D,
                0.5D, 0.5D, i == 0 ? 0.5D : 0.4D,
                new OfficeGradientStop(0D, start), new OfficeGradientStop(1D, end));
            else shape.FillGradient = OfficeLinearGradient.Horizontal(start, end);
            document.Shape(shape);
        }
        var (objects, _) = PdfSyntax.ParseObjects(document.ToBytes());
        PdfDictionary[] shadings = objects.Values.Select(value => value.Value).OfType<PdfDictionary>()
            .Where(dictionary => dictionary.Items.ContainsKey("ShadingType")).ToArray();
        foreach (string space in new[] { "DeviceRGB", "DeviceGray" }) {
            PdfDictionary[] matching = shadings.Where(dictionary => dictionary.Get<PdfName>("ColorSpace")?.Name == space).ToArray();
            Assert.Equal(2, matching.Length);
            int first = Assert.IsType<PdfReference>(matching[0].Items["Function"]).ObjectNumber;
            int second = Assert.IsType<PdfReference>(matching[1].Items["Function"]).ObjectNumber;
            Assert.Equal(first, second);
            Assert.IsType<PdfDictionary>(objects[first].Value);
        }
    }
}
