using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingMathFontCancellationTests {
    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)]
    public void CancelledOperationsDoNotInvokeFontResolutionOrPaint(int operation) {
        var source = new OfficeFontFaceCollection();
        source.Add("Fixture Math", ManagedTextShapingTestAssets.CreateMathFont());
        var program = new RecordingProgram(Assert.Single(source.Faces).Program);
        var options = new OfficeMathRenderOptions { Font = new OfficeFontInfo("Fixture Math", 20D) };
        options.Fonts.FontProgramProvider = new FixedProvider(program);
        options.Fonts.Add("Fixture Math", ManagedTextShapingTestAssets.CreateMathFont());
        var expression = OfficeMath.Identifier("x");
        var drawing = new OfficeDrawing(100D, 100D);
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => {
            if (operation == 0) OfficeMathRenderer.Measure(expression, options, cancelled.Token);
            else if (operation == 1) OfficeMathRenderer.Render(expression, options, cancelled.Token);
            else OfficeMathRenderer.AddToDrawing(drawing, expression, 0D, 0D, options, cancelled.Token);
        });
        Assert.Equal(0, program.CoverageQueries);
        Assert.Empty(drawing.Elements);
        OfficeMathRenderer.Measure(expression, options);
        Assert.True(program.CoverageQueries > 0); // The provider boundary is reachable in the same workflow.
    }

    private sealed class FixedProvider(IOfficeFontProgram program) : IOfficeFontProgramProvider {
        public OfficeFontProgramLoadResult TryLoad(OfficeFontProgramLoadRequest request) => new(program, request.Data.Length);
    }

    private sealed class RecordingProgram(IOfficeFontProgram inner) : IOfficeFontProgram {
        internal int CoverageQueries { get; private set; }
        public bool HasGlyphs(string text) { CoverageQueries++; return inner.HasGlyphs(text); }
        public string Fingerprint => inner.Fingerprint;
        public string? DisplayName => inner.DisplayName;
        public int? CollectionIndex => inner.CollectionIndex;
        public int UnitsPerEm => inner.UnitsPerEm;
        public bool IsOpenTypeCff => inner.IsOpenTypeCff;
        public bool ProvidesComplexTextLayout => inner.ProvidesComplexTextLayout;
        public double LineSpacingRatio => inner.LineSpacingRatio;
        public byte[] GetFontDataForShaping() => inner.GetFontDataForShaping();
        public double Measure(string text, double size) => inner.Measure(text, size);
        public IReadOnlyList<double> MeasureTextElements(IReadOnlyList<string> elements, double size) => inner.MeasureTextElements(elements, size);
        public double LineHeight(double size) => inner.LineHeight(size);
        public List<List<OfficePoint>> GetTextContours(string text, double x, double y, double size) => inner.GetTextContours(text, x, y, size);
        public bool TryGetGlyphMetrics(int scalar, out int glyphId, out int advance) => inner.TryGetGlyphMetrics(scalar, out glyphId, out advance);
        public double MeasureShapedText(string text, OfficeTextShapingResult result, double size) => inner.MeasureShapedText(text, result, size);
        public List<List<OfficePoint>> GetShapedTextContours(string text, OfficeTextShapingResult result, double x, double y, double size) =>
            inner.GetShapedTextContours(text, result, x, y, size);
    }
}
