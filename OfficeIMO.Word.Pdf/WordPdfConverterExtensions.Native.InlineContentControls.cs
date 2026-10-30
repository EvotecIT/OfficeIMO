using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using PdfCore = OfficeIMO.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static bool CanProjectNativeInlineContentControl(W.SdtRun control) =>
        (IsNativeCheckBoxControl(control) || IsNativeSupportedFormFieldContentControl(control)) &&
        control.SdtContentRun?.Descendants<W.Run>().Any() == true &&
        !control.Descendants<W.FieldChar>().Any() &&
        control.Ancestors<W.Paragraph>().FirstOrDefault()?.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>().Any() != true &&
        (IsNativeCheckBoxControl(control)
            ? !string.IsNullOrEmpty(GetNativeSdtText(control))
            : !string.IsNullOrEmpty(GetNativeContentControlDisplayValue(control)));

    private static string GetNativeContentControlDisplayValue(W.SdtRun control) {
        string? cached = GetNativeSdtText(control);
        if (!string.IsNullOrEmpty(cached)) return cached!;
        if (IsNativeDatePickerControl(control)) return GetNativeDatePickerValue(control);
        return GetNativeChoiceFieldValue(control, GetNativeChoiceFieldOptions(control)) ?? string.Empty;
    }

    private static IEnumerable<string> GetNativeCheckBoxAppearanceFontFamilies(WordParagraph paragraph) {
        foreach (W.SdtRun control in GetNativeCheckBoxControls(paragraph).Where(CanProjectNativeInlineContentControl)) {
            W14.SdtContentCheckBox properties = control.SdtProperties!.GetFirstChild<W14.SdtContentCheckBox>()!;
            string? checkedFont = properties.GetFirstChild<W14.CheckedState>()?.Font?.Value;
            string? uncheckedFont = properties.GetFirstChild<W14.UncheckedState>()?.Font?.Value;
            if (!string.IsNullOrWhiteSpace(checkedFont)) yield return checkedFont!;
            if (!string.IsNullOrWhiteSpace(uncheckedFont)) yield return uncheckedFont!;
        }
    }

    // The source run remains attached to the SDT so ordering, visibility and
    // formatting still use the original document. Only absent cached text needs
    // a view containing the control's current logical value.
    private static WordParagraph CreateNativeInlineControlFallbackRun(WordParagraph paragraph, W.SdtRun control,
        W.Hyperlink? hyperlink) {
        W.Run source = control.SdtContentRun!.Descendants<W.Run>().First();
        return new WordParagraph(paragraph._document, paragraph._paragraph!, source) {
            _visibleRun = CreateNativeResolvedSdtRun(control, GetNativeContentControlDisplayValue(control)),
            _hyperlink = hyperlink
        };
    }

    private sealed class NativeInlineContentControlProjection {
        private readonly Dictionary<W.SdtRun, PdfCore.PdfInlineFormField> _fields = new();
        private readonly HashSet<W.SdtRun> _paintedCheckBoxes = new();
        private readonly Dictionary<W.SdtRun, bool> _staticBackgroundControls = new();
        private readonly bool _tableCell;

        internal NativeInlineContentControlProjection(bool tableCell) => _tableCell = tableCell;

        internal bool TryCreateRuns(WordParagraph run, NativeTableStyleDefaults tableStyleDefaults,
            NativeDocumentDefaults nativeDefaults, NativeFontMap? fontMap,
            IReadOnlyList<WordTabStop> tabStops, ref int tabIndex, out IReadOnlyList<PdfCore.PdfTextRun> result) {
            result = Array.Empty<PdfCore.PdfTextRun>();
            W.SdtRun? control = run._run?.Ancestors<W.SdtRun>().FirstOrDefault(CanProjectNativeInlineContentControl);
            if (control == null || nativeDefaults.RunningContentContext != null) return false;
            if (!_staticBackgroundControls.TryGetValue(control, out bool staticBackground)) {
                // PDF widgets paint above page text. A run background belongs
                // beneath that text and can overlap neighboring glyphs. Keep
                // the whole source control in ordinary paragraph paint rather
                // than moving only part of its cache into an overlay widget.
                staticBackground = control.SdtContentRun!.Descendants<W.Run>().Any(source => {
                    var candidate = new WordParagraph(run._document, run._paragraph!, source);
                    return !IsNativeHiddenTextRun(candidate) && ResolveNativeTextRunStyle(candidate,
                        tableRunStyleDefaults: tableStyleDefaults.RunStyle, nativeDefaults: nativeDefaults,
                        nativeFontMap: fontMap).BackgroundColor.HasValue;
                });
                _staticBackgroundControls.Add(control, staticBackground);
                if (staticBackground) fontMap?.ReportInlineControlStaticBackground(
                    GetNativeContentControlFieldName(control, _staticBackgroundControls.Count - 1, "ContentControl"));
            }
            if (staticBackground) return false;
            NativeResolvedTextStyle style = ResolveNativeTextRunStyle(run, tableRunStyleDefaults: tableStyleDefaults.RunStyle,
                nativeDefaults: nativeDefaults, nativeFontMap: fontMap);
            double size = style.EffectiveFontSize;

            if (IsNativeCheckBoxControl(control)) {
                if (!_paintedCheckBoxes.Add(control)) return true;
                PdfCore.PdfTextRun measured = CreateNativeCellTextRun(run.Text, run,
                    tableStyleDefaults, nativeDefaults, fontMap, style);
                PdfCore.PdfWriter.PositionedTextBoxMeasurement? box = fontMap?.MeasureTextBox(measured);
                bool isChecked = IsNativeCheckBoxChecked(control);
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

        private static PdfCore.PdfTextRun? CreateCheckBoxStateAppearance(W.SdtRun control, bool isChecked,
            WordParagraph source, NativeTableStyleDefaults tableStyleDefaults,
            NativeDocumentDefaults nativeDefaults, NativeFontMap? fontMap) {
            W14.SdtContentCheckBox properties = control.SdtProperties!.GetFirstChild<W14.SdtContentCheckBox>()!;
            W14.CheckedState? checkedState = properties.GetFirstChild<W14.CheckedState>();
            W14.UncheckedState? uncheckedState = properties.GetFirstChild<W14.UncheckedState>();
            string? value = isChecked ? checkedState?.Val?.Value : uncheckedState?.Val?.Value;
            if (!int.TryParse(value, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out int scalar) ||
                scalar <= 0 || scalar > 0x10FFFF || scalar is >= 0xD800 and <= 0xDFFF) return null;
            string? font = isChecked ? checkedState?.Font?.Value : uncheckedState?.Font?.Value;
            W.RunProperties runProperties = (W.RunProperties?)GetNativeRunProperties(source)?.CloneNode(true)
                ?? new W.RunProperties();
            if (!string.IsNullOrWhiteSpace(font)) {
                runProperties.RunFonts = new W.RunFonts {
                    Ascii = font, HighAnsi = font, EastAsia = font, ComplexScript = font
                };
            }
            string glyph = char.ConvertFromUtf32(scalar);
            var run = new WordParagraph(source._document, source._paragraph,
                new W.Run(runProperties, new W.Text(glyph)));
            PdfCore.PdfTextRun appearance = CreateNativeCellTextRun(glyph, run,
                tableStyleDefaults, nativeDefaults, fontMap);
            return fontMap?.MeasureTextBox(appearance) == null ? null : appearance;
        }

        private PdfCore.PdfInlineFormField CreateField(W.SdtRun control, double fontSize, global::OfficeIMO.OfficeColor? color,
            PdfCore.PdfTextRun? checkedAppearanceRun = null, PdfCore.PdfTextRun? uncheckedAppearanceRun = null) {
            var style = new PdfCore.PdfFormFieldStyle {
                BorderWidth = global::OfficeIMO.OfficeLength.Points(0D), BorderColor = null, BackgroundColor = null,
                TextColor = color ?? global::OfficeIMO.OfficeColor.Black
            };
            string prefix = _tableCell ? "WordTable" : "Word";
            if (IsNativeCheckBoxControl(control)) {
                style.BorderWidth = global::OfficeIMO.OfficeLength.Points(.5D);
                style.BorderColor = global::OfficeIMO.OfficeColor.Black;
                return new PdfCore.PdfInlineFormField(PdfCore.PdfCanvasFormFieldItem.CheckBox(
                    GetNativeCheckBoxFieldName(control, _fields.Count, prefix + "CheckBox"),
                    IsNativeCheckBoxChecked(control), "Yes", "Yes", 0D, 0D, fontSize, fontSize, style),
                    checkedAppearanceRun, uncheckedAppearanceRun);
            }
            if (IsNativeDatePickerControl(control)) {
                return new PdfCore.PdfInlineFormField(PdfCore.PdfCanvasFormFieldItem.Text(
                    GetNativeContentControlFieldName(control, _fields.Count, prefix + "DatePicker"),
                    GetNativeDatePickerValue(control), GetNativeContentControlDisplayValue(control),
                    0D, 0D, 1D, fontSize, fontSize, style, style));
            }
            IReadOnlyList<string> options = GetNativeChoiceFieldOptions(control);
            string value = GetNativeChoiceFieldValue(control, options) ?? string.Empty;
            string kind = control.SdtProperties?.Elements<W.SdtContentComboBox>().Any() == true ? "ComboBox" : "DropDownList";
            return new PdfCore.PdfInlineFormField(PdfCore.PdfCanvasFormFieldItem.Choice(
                GetNativeContentControlFieldName(control, _fields.Count, prefix + kind), options,
                new[] { value }, 0D, 0D, 1D, fontSize, fontSize, isComboBox: true,
                allowsMultipleSelection: false, style: style));
        }
    }
}
