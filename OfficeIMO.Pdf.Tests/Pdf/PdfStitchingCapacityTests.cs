using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfColorFunctionTests {
    [Theory]
    [InlineData(64, true)]
    [InlineData(256, true)]
    [InlineData(257, false)]
    public void Type3SupportsExpandedGradientsWithinBoundedFunctionCount(int count, bool supported) {
        var children = new PdfArray();
        for (int i = 0; i < count; i++) children.Items.Add(Type2(new[] { 0D }, new[] { 1D }));
        var functionObject = Dictionary(
            ("FunctionType", Number(3)), ("Domain", Numbers(0D, 1D)),
            ("Functions", children),
            ("Bounds", Numbers(Enumerable.Range(1, count - 1).Select(i => (double)i / count).ToArray())),
            ("Encode", Numbers(Enumerable.Range(0, count * 2).Select(i => (double)(i % 2)).ToArray())));
        bool resolved = PdfColorSpaceFunctionResolver.TryCreateFunction(functionObject, 1, 1,
            new Dictionary<int, PdfIndirectObject>(), 256 * 1024, out PdfColorFunction function);
        Assert.Equal(supported, resolved);
        if (!supported) return;
        var result = new double[1];
        Assert.True(function.TryEvaluate(new[] { 12.25 / count }, result));
        Assert.Equal(.25, result[0], 8);
    }
}
