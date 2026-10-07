using System.Globalization;
using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void ReportUnsupportedFormFieldBorderStyleFallback(string source, string borderStyle) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldBorderStyleStaticFallback,
            "An HTML form control used faithful static rendering because its authored border style cannot be represented by a PDF widget appearance.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "border-style=" + borderStyle,
            OfficeConversionLossKind.Approximation);
    }

    private static bool CanPreserveInteractiveFieldTypography(OfficeFontInfo font) {
        if (font.Style != OfficeFontStyle.Regular) return false;
        string familyList = font.FamilyName.Trim();
        IReadOnlyList<string> families = HtmlRenderCssValues.SplitTopLevelCommas(familyList);
        string family = (families.Count == 0 ? familyList : families[0]).Trim().Trim('\'', '"');
        return string.Equals(family, "Arial", StringComparison.OrdinalIgnoreCase)
            || string.Equals(family, "Helvetica", StringComparison.OrdinalIgnoreCase)
            || string.Equals(family, "sans-serif", StringComparison.OrdinalIgnoreCase);
    }

    private void ReportFormFieldTypographyFallback(string source, OfficeFontInfo font) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldTypographyStaticFallback,
            "An HTML text or choice control used faithful static rendering because a PDF widget cannot preserve its authored typography.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "font=" + font,
            OfficeConversionLossKind.Approximation);
    }

    private void ReportNoWrapFormFieldFallback(string source) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldNoWrapStaticFallback,
            "An HTML textarea with wrap=off used faithful static rendering because a PDF multiline widget appearance cannot preserve no-wrap semantics.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "textarea[wrap=off]",
            OfficeConversionLossKind.Approximation);
    }

    private void ReportTransformedFormFieldFallback(string source, string detail) {
        if (!_reportedTransformedFormFieldFallbacks.Add(source)) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldTransformStaticFallback,
            "An HTML form control inside a transformed, translucent, or clipped paint group was rendered as static content because PDF widget annotations cannot preserve the authored appearance.",
            HtmlDiagnosticSeverity.Warning,
            source,
            detail,
            OfficeConversionLossKind.Approximation);
    }

    private void ReportDuplicateRadioValueFallback(string source, string groupKey) {
        if (!_reportedStaticRadioGroups.Add(groupKey)) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.RadioDuplicateValueStaticFallback,
            "An HTML radio group with duplicate submitted values was rendered as static content because PDF radio appearance-state values must be unique.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "group=" + groupKey.Substring(groupKey.LastIndexOf('\n') + 1),
            OfficeConversionLossKind.Approximation);
    }

    private void ReportMixedDisabledRadioGroupFallback(string source, string groupKey) {
        if (!_reportedStaticRadioGroups.Add(groupKey)) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.RadioMixedDisabledStateStaticFallback,
            "An HTML radio group mixing enabled and disabled options was rendered as static content because PDF radio widgets cannot preserve disabled state per option.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "group=" + groupKey.Substring(groupKey.LastIndexOf('\n') + 1),
            OfficeConversionLossKind.Approximation);
    }

    private void ReportZeroMaximumLengthFallback(string source) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldZeroMaximumLengthStaticFallback,
            "An HTML text control with maxlength=0 was rendered as static content because PDF /MaxLen must be positive.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "maxlength=0",
            OfficeConversionLossKind.Approximation);
    }

    private void ReportInitialValueExceedsMaximumLengthFallback(string source, int maximumLength, int valueLength) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldInitialValueExceedsMaximumLengthStaticFallback,
            "An HTML text control whose initial value exceeds maxlength was rendered as static content because PDF /MaxLen cannot preserve that authored state consistently.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "maxlength=" + maximumLength.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + ";value-length=" + valueLength.ToString(System.Globalization.CultureInfo.InvariantCulture),
            OfficeConversionLossKind.Approximation);
    }

    private void ReportMultipleFileSelectionFallback(string source) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FileMultipleSelectionStaticFallback,
            "An HTML multiple-file input was rendered as static content because PDF file-select fields cannot preserve multiple-file selection semantics.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "input[type=file][multiple]",
            OfficeConversionLossKind.Approximation);
    }

    private void ReportBlankFormFieldNameFallback(string source) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldBlankNameStaticFallback,
            "An HTML form control with a whitespace-only name was rendered as static content because PDF form field names must contain a non-whitespace character.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "whitespace-only name",
            OfficeConversionLossKind.Approximation);
    }

    private void ReportBlankButtonValueFallback(string source, string? groupKey) {
        if (groupKey != null && !_reportedStaticRadioGroups.Add(groupKey)) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldBlankButtonValueStaticFallback,
            "An HTML checkbox or radio control with an empty or whitespace-only submitted value was rendered as static content because PDF button export values must contain a non-whitespace character.",
            HtmlDiagnosticSeverity.Warning,
            source,
            groupKey == null ? "blank submitted value" : "group=" + groupKey.Substring(groupKey.LastIndexOf('\n') + 1),
            OfficeConversionLossKind.Approximation);
    }

    private void ReportDuplicateSelectedChoiceValueFallback(string source) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.ChoiceDuplicateSelectedValueStaticFallback,
            "An HTML multi-select with duplicate selected submitted values was rendered as static content because a value-only PDF choice selection cannot preserve both selected option identities.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "duplicate selected option values",
            OfficeConversionLossKind.Approximation);
    }

    private void ReportDisabledChoiceOptionFallback(string source) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.ChoiceDisabledOptionStaticFallback,
            "An HTML select containing disabled options was rendered as static content because PDF choice fields cannot preserve disabled state per option.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "disabled option",
            OfficeConversionLossKind.Approximation);
    }

    private void ReportRepeatedFormControlNameFallback(string source, string groupKey) {
        if (!_reportedStaticRepeatedControlGroups.Add(groupKey)) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldRepeatedNameStaticFallback,
            "HTML controls that repeat across PDF pages or share one submitted name were rendered as static content because separate PDF widgets cannot safely preserve that authored field identity.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "name=" + groupKey.Substring(groupKey.LastIndexOf('\n') + 1),
            OfficeConversionLossKind.Approximation);
    }

    private void ReportBlankChoiceLabelFallback(string source) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.ChoiceBlankLabelStaticFallback,
            "An HTML select containing a blank option label was rendered as static content because PDF choice fields require non-empty display labels.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "blank option label",
            OfficeConversionLossKind.Approximation);
    }

    private void ReportEmptyChoiceOptionsFallback(string source) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.ChoiceEmptyOptionsStaticFallback,
            "An empty HTML select was rendered as static content because interactive PDF choice fields require at least one option.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "empty option list",
            OfficeConversionLossKind.Approximation);
    }

    private void ReportTransparentFormFieldPaintFallback(string source, string? groupKey) {
        if (groupKey != null && !_reportedStaticRadioGroups.Add("transparent\n" + groupKey)) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldColorTransparencyStaticFallback,
            "An HTML form control with translucent paint was rendered as static content because generated PDF widget appearances cannot preserve color alpha.",
            HtmlDiagnosticSeverity.Warning,
            source,
            groupKey == null
                ? "translucent form-control paint"
                : "translucent radio-group paint; group=" + groupKey.Substring(groupKey.LastIndexOf('\n') + 1),
            OfficeConversionLossKind.Approximation);
    }

    private void ReportBackgroundImageFormFieldFallback(string source, string? groupKey) {
        if (groupKey != null && !_reportedStaticRadioGroups.Add("background-image\n" + groupKey)) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldBackgroundImageStaticFallback,
            "An HTML form control with background-image paint was rendered as static content because generated PDF widget appearances cannot preserve background layers.",
            HtmlDiagnosticSeverity.Warning,
            source,
            groupKey == null
                ? "form-control background-image"
                : "radio-group background-image; group=" + groupKey.Substring(groupKey.LastIndexOf('\n') + 1),
            OfficeConversionLossKind.Approximation);
    }

}
