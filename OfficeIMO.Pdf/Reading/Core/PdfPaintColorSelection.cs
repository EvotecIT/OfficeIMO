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

    /// <summary>
    /// Reads the trailing numeric components from typed operands without boxing them. A plain device color equal to
    /// <paramref name="current"/> keeps that selection instead of allocating an identical one.
    /// </summary>
    internal static bool TryCreate(
        PdfContentOperands operands,
        PdfPageColorSpace colorSpace,
        OfficeIccRenderingIntent renderingIntent,
        PdfPaintColorSelection? current,
        out PdfPaintColorSelection? selection,
        out OfficeColor color,
        PdfOutputIntentColorTransform? outputIntentColorTransform = null) {
        selection = null;
        color = OfficeColor.Black;
        int componentCount = colorSpace.ComponentCount;
        int endIndex = operands.Count;
        while (endIndex > 0 && !operands.TryGetNumber(endIndex - 1, out _)) endIndex--;
        if (componentCount < 1 || endIndex < componentCount) return false;

        int startIndex = endIndex - componentCount;
        if (current != null && colorSpace.IsPlainDeviceColorSpace && current.ColorSpace.IsPlainDeviceColorSpace &&
            current.ColorSpace.Kind == colorSpace.Kind &&
            ReferenceEquals(current._outputIntentColorTransform, outputIntentColorTransform) &&
            current.HasComponents(operands, startIndex)) {
            selection = current;
            return current.TryConvert(renderingIntent, out color);
        }

        var components = new double[componentCount];
        for (int index = 0; index < componentCount; index++) {
            components[index] = operands.NumberOrZero(startIndex + index);
        }

        selection = new PdfPaintColorSelection(colorSpace, components, outputIntentColorTransform);
        return selection.TryConvert(renderingIntent, out color);
    }

    private bool HasComponents(PdfContentOperands operands, int startIndex) {
        for (int index = 0; index < _components.Length; index++) {
            if (BitConverter.DoubleToInt64Bits(_components[index]) !=
                BitConverter.DoubleToInt64Bits(operands.NumberOrZero(startIndex + index))) return false;
        }
        return true;
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
        if (_outputIntentColorTransform != null &&
            _outputIntentColorTransform.TryApplyDirect(ColorSpace, _components, renderingIntent, out color)) return true;
        if (!ColorSpace.TryConvertColor(_components, renderingIntent, out color)) return false;
        if (_outputIntentColorTransform != null) {
            color = _outputIntentColorTransform.Apply(color, renderingIntent);
        }
        return true;
    }
}
