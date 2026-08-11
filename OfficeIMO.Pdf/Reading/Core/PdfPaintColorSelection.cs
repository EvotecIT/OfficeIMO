using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Retains authored color components so the current graphics-state rendering intent can be
/// applied when the color is painted, including when <c>ri</c> or ExtGState <c>RI</c> changes
/// after <c>sc</c>/<c>scn</c>.
/// </summary>
internal sealed class PdfPaintColorSelection {
    private readonly double[] _components;

    private PdfPaintColorSelection(PdfPageColorSpace colorSpace, double[] components) {
        ColorSpace = colorSpace;
        _components = components;
    }

    internal PdfPageColorSpace ColorSpace { get; }

    internal static bool TryCreate(
        IReadOnlyList<object> operands,
        PdfPageColorSpace colorSpace,
        OfficeIccRenderingIntent renderingIntent,
        out PdfPaintColorSelection? selection,
        out OfficeColor color) {
        selection = null;
        color = OfficeColor.Black;
        int componentCount = colorSpace.ComponentCount;
        int endIndex = operands.Count;
        while (endIndex > 0 && operands[endIndex - 1] is not double) endIndex--;
        if (componentCount < 1 || endIndex < componentCount) return false;

        int startIndex = endIndex - componentCount;
        var components = new double[componentCount];
        for (int index = 0; index < componentCount; index++) {
            components[index] = operands[startIndex + index] is double value ? value : 0D;
        }

        selection = new PdfPaintColorSelection(colorSpace, components);
        return selection.TryConvert(renderingIntent, out color);
    }

    internal bool TryConvert(OfficeIccRenderingIntent renderingIntent, out OfficeColor color) =>
        ColorSpace.TryConvertColor(_components, renderingIntent, out color);
}
