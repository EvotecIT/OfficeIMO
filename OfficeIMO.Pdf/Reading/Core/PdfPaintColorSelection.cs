using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Retains authored color components so the current graphics-state rendering intent can be
/// applied when the color is painted, including when <c>ri</c> or ExtGState <c>RI</c> changes
/// after <c>sc</c>/<c>scn</c>.
/// </summary>
internal sealed class PdfPaintColorSelection {
    private readonly double[] _components;
    private readonly PdfOutputIntentColorTransform? _outputIntentColorTransform;

    private PdfPaintColorSelection(
        PdfPageColorSpace colorSpace,
        double[] components,
        PdfOutputIntentColorTransform? outputIntentColorTransform) {
        ColorSpace = colorSpace;
        _components = components;
        _outputIntentColorTransform = outputIntentColorTransform;
    }

    internal PdfPageColorSpace ColorSpace { get; }

    internal static bool TryCreate(
        IReadOnlyList<object> operands,
        PdfPageColorSpace colorSpace,
        OfficeIccRenderingIntent renderingIntent,
        out PdfPaintColorSelection? selection,
        out OfficeColor color,
        PdfOutputIntentColorTransform? outputIntentColorTransform = null) {
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

        selection = new PdfPaintColorSelection(colorSpace, components, outputIntentColorTransform);
        return selection.TryConvert(renderingIntent, out color);
    }

    internal static bool TryCreateDefaultBlack(
        OfficeIccRenderingIntent renderingIntent,
        PdfOutputIntentColorTransform outputIntentColorTransform,
        out PdfPaintColorSelection? selection,
        out OfficeColor color) =>
        TryCreate(
            new object[] { 0D },
            PdfPageColorSpaceKind.DeviceGray,
            renderingIntent,
            out selection,
            out color,
            outputIntentColorTransform);

    internal bool TryConvert(OfficeIccRenderingIntent renderingIntent, out OfficeColor color) {
        if (!ColorSpace.TryConvertColor(_components, renderingIntent, out color)) return false;
        if (_outputIntentColorTransform != null) {
            color = _outputIntentColorTransform.Apply(ColorSpace, _components, color, renderingIntent);
        }
        return true;
    }
}
