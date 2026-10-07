using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfFunctionShadingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TwoInputFieldHonorsMatrixDomainBoundsAndSharedEvaluationBudget(bool componentArray) {
        var functions = componentArray
            ? new[] { Function("{ pop }", 1), Function("{ exch pop }", 1), Function("{ add 2 div }", 1) }
            : new[] { Function("{ 2 copy add 2 div }", 3) };
        int charged = 0;
        var field = new PdfFunctionShading(functions, PdfPageColorSpaceKind.DeviceRgb,
            new[] { 0D, 1D, 0D, 1D }, new OfficeTransform(2, 0, 0, 4, 10, 20),
            new[] { 10D, 20D, 11.5D, 24D }, OfficeIccRenderingIntent.RelativeColorimetric, null,
            cost => { charged++; return charged <= 1; }, default);
        var input = new double[2]; var output = new double[3];
        Assert.True(field.TrySample(10.5, 23, input, output, out var color));
        Assert.InRange(color.R, 63, 64); Assert.InRange(color.G, 191, 192); Assert.InRange(color.B, 127, 128);
        Assert.True(field.TrySample(9, 23, input, output, out color));
        Assert.Equal(OfficeColor.Transparent, color); // Outside the transformed domain.
        Assert.True(field.TrySample(11.75, 23, input, output, out color));
        Assert.Equal(OfficeColor.Transparent, color); // Inside domain but outside BBox.
        Assert.Equal(1, charged);
        Assert.False(field.TrySample(10.5, 23, input, output, out _));
        Assert.Equal(2, charged);
    }

    [Fact]
    public void CanceledFieldDoesNotEvaluateOrReturnAnUnpaintedSuccess() {
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        var field = new PdfFunctionShading(new[] { Function("{ 2 copy add 2 div }", 3) },
            PdfPageColorSpaceKind.DeviceRgb, new[] { 0D, 1D, 0D, 1D }, OfficeTransform.Identity, null,
            OfficeIccRenderingIntent.RelativeColorimetric, null, _ => throw new Exception("Unexpected work"), cancellation.Token);
        Assert.Throws<OperationCanceledException>(() => field.TrySample(2, 2, new double[2], new double[3], out _));
    }

    private static PdfColorFunction Function(string source, int count) {
        var dictionary = new PdfDictionary();
        dictionary.Items["FunctionType"] = new PdfNumber(4);
        var domain = new PdfArray();
        domain.Items.AddRange(new PdfObject[] { new PdfNumber(0), new PdfNumber(1), new PdfNumber(0), new PdfNumber(1) });
        dictionary.Items["Domain"] = domain;
        var range = new PdfArray();
        range.Items.AddRange(Enumerable.Range(0, count).SelectMany(_ => new PdfObject[] { new PdfNumber(0), new PdfNumber(1) }));
        dictionary.Items["Range"] = range;
        Assert.True(PdfColorSpaceFunctionResolver.TryCreateFunction(new PdfStream(dictionary, Encoding.ASCII.GetBytes(source)),
            2, count, new Dictionary<int, PdfIndirectObject>(), 1024 * 1024, out var result));
        return result;
    }
}
