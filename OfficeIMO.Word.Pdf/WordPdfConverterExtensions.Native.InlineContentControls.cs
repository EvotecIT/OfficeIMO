using System.Globalization;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;


public static partial class WordPdfConverterExtensions {
    private static bool CanProjectNativeInlineContentControl(OfficeXmlNode control) =>
        (IsNativeCheckBoxControl(control) || IsNativeSupportedFormFieldContentControl(control)) &&
        control.WChild("sdtContent").ModelDescendants("r").Any() &&
        !control.HasDescendant(WordXml.W, "fldChar") &&
        control.Ancestor(WordXml.W, "p").HasDescendant(WordXml.M, "oMath") != true &&
        (IsNativeCheckBoxControl(control)
            ? !string.IsNullOrEmpty(GetNativeSdtText(control))
            : !string.IsNullOrEmpty(GetNativeContentControlDisplayValue(control)));

    private static string GetNativeContentControlDisplayValue(OfficeXmlNode control) {
        OfficeXmlNode sdtRun = control;
        string? cached = GetNativeSdtText(sdtRun);
        if (!string.IsNullOrEmpty(cached)) return cached!;
        if (IsNativeDatePickerControl(sdtRun)) return GetNativeDatePickerValue(sdtRun);
        return GetNativeChoiceFieldValue(sdtRun, GetNativeChoiceFieldOptions(sdtRun)) ?? string.Empty;
    }


    /// <summary>The inline content controls enclosing <paramref name="element"/>, nearest first.</summary>
    private static IEnumerable<OfficeXmlNode> GetNativeAncestorSdtRuns(OfficeXmlNode element) =>
        element.Ancestors().Where(WordXml.IsRunContentControl);

    private static IEnumerable<string> GetNativeCheckBoxAppearanceFontFamilies(WordParagraph paragraph) {
        foreach (OfficeXmlNode control in GetNativeCheckBoxControls(paragraph)
            .Where(CanProjectNativeInlineContentControl)) {
            OfficeXmlNode properties = control.WChild("sdtPr").Child(WordXml.W14, "checkbox");
            string? checkedFont = properties.Child(WordXml.W14, "checkedState").Attr(WordXml.W14, "font");
            string? uncheckedFont = properties.Child(WordXml.W14, "uncheckedState").Attr(WordXml.W14, "font");
            if (!string.IsNullOrWhiteSpace(checkedFont)) yield return checkedFont!;
            if (!string.IsNullOrWhiteSpace(uncheckedFont)) yield return uncheckedFont!;
        }
    }

    // The source run remains attached to the SDT so ordering, visibility and
    // formatting still use the original document. Only absent cached text needs
    // a view containing the control's current logical value.
    private static WordParagraph CreateNativeInlineControlFallbackRun(WordParagraph paragraph, OfficeXmlNode control,
        OfficeXmlNode hyperlink) {
        OfficeXmlNode source = control.WChild("sdtContent").ModelDescendants("r").First();
        return paragraph._document.RunView(paragraph.Node(), source, hyperlink,
            CreateNativeResolvedSdtRun(paragraph.Node(), control, GetNativeContentControlDisplayValue(control)));
    }

    private sealed class NativeInlineContentControlProjection {
        private readonly Dictionary<OfficeXmlNode, PdfCore.PdfInlineFormField> _fields = new();
        private readonly HashSet<OfficeXmlNode> _paintedCheckBoxes = new();
        private readonly Dictionary<OfficeXmlNode, bool> _staticBackgroundControls = new();
        private readonly bool _tableCell;

        internal NativeInlineContentControlProjection(bool tableCell) => _tableCell = tableCell;

        internal bool TryCreateRuns(WordParagraph run, NativeTableStyleDefaults tableStyleDefaults,
            NativeDocumentDefaults nativeDefaults, NativeFontMap? fontMap,
            IReadOnlyList<WordTabStop> tabStops, ref int tabIndex, out IReadOnlyList<PdfCore.PdfTextRun> result) {
            result = Array.Empty<PdfCore.PdfTextRun>();
            OfficeXmlNode control = GetNativeAncestorSdtRuns(run.RunNode()).FirstOrDefault(CanProjectNativeInlineContentControl);
            if (control.IsNone || nativeDefaults.RunningContentContext != null) return false;
            OfficeXmlNode sdtRun = control;
            if (!_staticBackgroundControls.TryGetValue(control, out bool staticBackground)) {
                // PDF widgets paint above page text. A run background belongs
                // beneath that text and can overlap neighboring glyphs. Keep
                // the whole source control in ordinary paragraph paint rather
                // than moving only part of its cache into an overlay widget.
                staticBackground = control.WChild("sdtContent").ModelDescendants("r").Any(source => {
                    WordParagraph candidate = run._document.RunView(run.Node(), source);
                    return !IsNativeHiddenTextRun(candidate) && ResolveNativeTextRunStyle(candidate,
                        tableRunStyleDefaults: tableStyleDefaults.RunStyle, nativeDefaults: nativeDefaults,
                        nativeFontMap: fontMap).BackgroundColor.HasValue;
                });
                _staticBackgroundControls.Add(control, staticBackground);
                if (staticBackground) fontMap?.ReportInlineControlStaticBackground(
                    GetNativeContentControlFieldName(sdtRun, _staticBackgroundControls.Count - 1, "ContentControl"));
            }
            if (staticBackground) return false;
            NativeResolvedTextStyle style = ResolveNativeTextRunStyle(run, tableRunStyleDefaults: tableStyleDefaults.RunStyle,
                nativeDefaults: nativeDefaults, nativeFontMap: fontMap);
            double size = style.EffectiveFontSize;

            if (IsNativeCheckBoxControl(sdtRun)) {
                if (!_paintedCheckBoxes.Add(control)) return true;
                PdfCore.PdfTextRun measured = CreateNativeCellTextRun(run.Text, run,
                    tableStyleDefaults, nativeDefaults, fontMap, style);
                PdfCore.PdfWriter.PositionedTextBoxMeasurement? box = fontMap?.MeasureTextBox(measured);
                bool isChecked = IsNativeCheckBoxChecked(sdtRun);
                PdfCore.PdfTextRun? cachedAppearance = box == null ? null : measured;
                PdfCore.PdfTextRun? otherAppearance = CreateCheckBoxStateAppearance(control, !isChecked, run,
                    tableStyleDefaults, nativeDefaults, fontMap);
                var checkBoxField = CreateField(control, size, style.Color,
                    isChecked ? cachedAppearance : otherAppearance,
                    isChecked ? otherAppearance : cachedAppearance);
                _fields.Add(control, checkBoxField);
                if (cachedAppearance == null) fontMap?.ReportCheckBoxAppearanceFallback(checkBoxField.Field.Name, isChecked);
                if (otherAppearance == null) fontMap?.ReportCheckBoxAppearanceFallback(checkBoxField.Field.Name, !isChecked);
                double width = box?.Width ?? size;
                if (width <= 0D || double.IsNaN(width) || double.IsInfinity(width)) width = size;
                double height = box?.Ascent + box?.Descent ?? size;
                double baselineOffset = -(box?.Descent ?? 0D);
                result = new[] { PdfCore.PdfTextRun.Inline(new PdfCore.PdfInlineFormWidget(checkBoxField, width, height,
                    baselineOffset, sourceRun: measured)) };
                return true;
            }

            if (!_fields.TryGetValue(control, out PdfCore.PdfInlineFormField? field)) {
                field = CreateField(control, size, style.Color);
                _fields.Add(control, field);
            }

            var textRuns = new List<PdfCore.PdfTextRun>();
            AddNativeCasedCellTextRuns(textRuns, run.Text, run, tableStyleDefaults, nativeDefaults,
                fontMap, tabStops, ref tabIndex, run.IsHyperLink ? run.Hyperlink : null);
            result = textRuns.Select(textRun => textRun.WithFormField(field)).ToArray();
            return true;
        }

        private static PdfCore.PdfTextRun? CreateCheckBoxStateAppearance(OfficeXmlNode control, bool isChecked,
            WordParagraph source, NativeTableStyleDefaults tableStyleDefaults,
            NativeDocumentDefaults nativeDefaults, NativeFontMap? fontMap) {
            OfficeXmlNode properties = control.WChild("sdtPr").Child(WordXml.W14, "checkbox");
            OfficeXmlNode checkedState = properties.Child(WordXml.W14, "checkedState");
            OfficeXmlNode uncheckedState = properties.Child(WordXml.W14, "uncheckedState");
            string? value = isChecked ? checkedState.Attr(WordXml.W14, "val") : uncheckedState.Attr(WordXml.W14, "val");
            if (!int.TryParse(value, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out int scalar) ||
                scalar <= 0 || scalar > 0x10FFFF || scalar is >= 0xD800 and <= 0xDFFF) return null;
            string? font = isChecked ? checkedState.Attr(WordXml.W14, "font") : uncheckedState.Attr(WordXml.W14, "font");
            OfficeXmlNode paragraph = source.Node();
            OfficeXmlNode runElement = paragraph.CreateElement(WordXml.W, "r");
            OfficeXmlNode sourceProperties = GetNativeRunProperties(source);
            OfficeXmlNode runProperties = runElement.Append(sourceProperties.Exists
                ? sourceProperties.Clone() : runElement.CreateElement(WordXml.W, "rPr"));
            if (!string.IsNullOrWhiteSpace(font)) {
                OfficeXmlNode fonts = WordXmlEdits.ReplaceChild(runProperties, "rFonts", WRPrView.Order);
                fonts.SetAttribute(WordXml.W, "ascii", font);
                fonts.SetAttribute(WordXml.W, "hAnsi", font);
                fonts.SetAttribute(WordXml.W, "eastAsia", font);
                fonts.SetAttribute(WordXml.W, "cs", font);
            }
            string glyph = char.ConvertFromUtf32(scalar);
            OfficeXmlNode text = runElement.CreateElement(WordXml.W, "t");
            text.SetInnerText(glyph);
            runElement.Append(text);
            WordParagraph run = source._document.RunView(paragraph, runElement);
            PdfCore.PdfTextRun appearance = CreateNativeCellTextRun(glyph, run,
                tableStyleDefaults, nativeDefaults, fontMap);
            return fontMap?.MeasureTextBox(appearance) == null ? null : appearance;
        }

        private PdfCore.PdfInlineFormField CreateField(OfficeXmlNode control, double fontSize, global::OfficeIMO.OfficeColor? color,
            PdfCore.PdfTextRun? checkedAppearanceRun = null, PdfCore.PdfTextRun? uncheckedAppearanceRun = null) {
            OfficeXmlNode sdtRun = control;
            var style = new PdfCore.PdfFormFieldStyle {
                BorderWidth = global::OfficeIMO.OfficeLength.Points(0D), BorderColor = null, BackgroundColor = null,
                TextColor = color ?? global::OfficeIMO.OfficeColor.Black
            };
            string prefix = _tableCell ? "WordTable" : "Word";
            if (IsNativeCheckBoxControl(sdtRun)) {
                style.BorderWidth = global::OfficeIMO.OfficeLength.Points(.5D);
                style.BorderColor = global::OfficeIMO.OfficeColor.Black;
                return new PdfCore.PdfInlineFormField(PdfCore.PdfCanvasFormFieldItem.CheckBox(
                    GetNativeCheckBoxFieldName(sdtRun, _fields.Count, prefix + "CheckBox"),
                    IsNativeCheckBoxChecked(sdtRun), "Yes", "Yes", 0D, 0D, fontSize, fontSize, style),
                    checkedAppearanceRun, uncheckedAppearanceRun);
            }
            if (IsNativeDatePickerControl(sdtRun)) {
                return new PdfCore.PdfInlineFormField(PdfCore.PdfCanvasFormFieldItem.Text(
                    GetNativeContentControlFieldName(sdtRun, _fields.Count, prefix + "DatePicker"),
                    GetNativeDatePickerValue(sdtRun), GetNativeContentControlDisplayValue(control),
                    0D, 0D, 1D, fontSize, fontSize, style, style));
            }
            IReadOnlyList<string> options = GetNativeChoiceFieldOptions(sdtRun);
            string value = GetNativeChoiceFieldValue(sdtRun, options) ?? string.Empty;
            string kind = control.WChild("sdtPr").WChild("comboBox").Exists ? "ComboBox" : "DropDownList";
            return new PdfCore.PdfInlineFormField(PdfCore.PdfCanvasFormFieldItem.Choice(
                GetNativeContentControlFieldName(sdtRun, _fields.Count, prefix + kind), options,
                new[] { value }, 0D, 0D, 1D, fontSize, fontSize, isComboBox: true,
                allowsMultipleSelection: false, style: style));
        }
    }
}
