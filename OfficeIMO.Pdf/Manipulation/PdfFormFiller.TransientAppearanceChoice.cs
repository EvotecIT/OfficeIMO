using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    private static bool TryBuildTransientChoiceAppearance(Dictionary<int, PdfIndirectObject> objects,
        string? defaultAppearance, PdfDictionary? defaultResources, double width, double height,
        PdfArray options, PdfObject? value, PdfArray? indices, int topIndex, int flags, PdfFormFieldStyle style,
        out string content, out PdfDictionary? resources, int maximumDecodedBytes, CancellationToken cancellationToken) {
        content = string.Empty;
        resources = null;
        ChoiceAppearanceData data = ReadChoiceAppearanceData(objects, options, value, indices, topIndex, flags, cancellationToken);
        IReadOnlyList<string> labels = data.Labels;
        IReadOnlyList<int> selected = data.Selected;
        double fontSize = ReadWidgetAppearanceFontSize(defaultAppearance, height);
        string visibleText = data.VisibleText(height, fontSize);
        if (!TryPrepareTransientTextFont(objects, defaultAppearance, defaultResources, visibleText,
                out string fontName, out TextAppearanceFontPlan? plan, out resources, maximumDecodedBytes, cancellationToken)) return false;
        content = PdfAcroFormDictionaryBuilder.BuildChoiceListAppearanceContent(width, height, labels.Count, data.TopIndex,
            selected, fontSize, style, (builder, index, baseline, textStyle) => {
                cancellationToken.ThrowIfCancellationRequested();
                builder.Append(PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(width, height, labels[index],
                    fontSize, textStyle, fontResourceName: fontName, encodeTextSegmentHex: plan?.EncodeTextSegmentHex,
                    measureTextSegmentWidth: plan?.MeasureTextSegmentWidth, singleLineBaseline: baseline));
            });
        return true;
    }
    internal sealed class ChoiceAppearanceData {
        internal IReadOnlyList<string> Labels { get; }
        internal IReadOnlyList<int> Selected { get; }
        internal int TopIndex { get; }
        internal ChoiceAppearanceData(IReadOnlyList<string> labels, IReadOnlyList<int> selected, int topIndex) {
            Labels = labels; Selected = selected; TopIndex = topIndex;
        }
        internal string VisibleText(double height, double fontSize) {
            (int first, int visible) = PdfAcroFormDictionaryBuilder.GetChoiceListVisibleRange(height, fontSize, Labels.Count, TopIndex);
            return string.Concat(Labels.Skip(first).Take(visible));
        }
    }

    internal static ChoiceAppearanceData ReadChoiceAppearanceData(Dictionary<int, PdfIndirectObject> objects,
        PdfArray options, PdfObject? value, PdfArray? indices, int topIndex, int flags, CancellationToken cancellationToken) {
        var selectedValues = new HashSet<string>(StringComparer.Ordinal);
        if (value is PdfArray values) {
            foreach (PdfObject item in values.Items) {
                cancellationToken.ThrowIfCancellationRequested();
                if (TryReadOptionText(objects, item, out string? text) && text != null) selectedValues.Add(text);
            }
        } else if (value != null && TryReadOptionText(objects, value, out string? single) && single != null) selectedValues.Add(single);
        var selected = new List<int>();
        if (indices != null) foreach (PdfObject item in indices.Items) {
            cancellationToken.ThrowIfCancellationRequested();
            if (ResolveObject(objects, item) is PdfNumber number && number.Value >= 0 && number.Value < options.Items.Count &&
                number.Value == Math.Truncate(number.Value)) selected.Add((int)number.Value);
        }
        bool useIndices = indices != null;
        var labels = new List<string>(options.Items.Count);
        for (int index = 0; index < options.Items.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject? option = ResolveObject(objects, options.Items[index]);
            string export = string.Empty, display = string.Empty;
            if (option is PdfArray pair && pair.Items.Count >= 2) {
                TryReadOptionText(objects, pair.Items[0], out string? exportText);
                TryReadOptionText(objects, pair.Items[1], out string? displayText);
                export = exportText ?? string.Empty; display = displayText ?? string.Empty;
            } else if (option != null && TryReadOptionText(objects, option, out string? label)) export = display = label ?? string.Empty;
            labels.Add(display);
            if (!useIndices && selectedValues.Contains(export) && ((flags & MultiSelectChoiceFlag) != 0 || selected.Count == 0)) selected.Add(index);
        }
        if ((flags & MultiSelectChoiceFlag) == 0 && selected.Count > 1) selected.RemoveRange(1, selected.Count - 1);
        return new ChoiceAppearanceData(labels, selected, topIndex);
    }

}
