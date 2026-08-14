using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static void ResolveChoiceFieldValues(
        IElement select,
        out IReadOnlyList<string> options,
        out IReadOnlyList<string> optionValues,
        out IReadOnlyList<string> values,
        out IReadOnlyList<int> selectedOptionIndices,
        out bool hasDuplicateSelectedValues,
        out bool hasAmbiguousSelectedValue) {
        IElement[] optionElements = select.QuerySelectorAll("option").ToArray();
        var labels = new List<string>(optionElements.Length);
        var exports = new List<string>(optionElements.Length);
        var selectedExports = new List<string>();
        var selectedIndices = new List<int>();
        hasDuplicateSelectedValues = false;
        hasAmbiguousSelectedValue = false;
        IReadOnlyList<IElement> selectedOptions = HtmlFormControlSemantics.GetEffectiveSelectedOptions(select);
        for (int index = 0; index < optionElements.Length; index++) {
            IElement option = optionElements[index];
            string label = NormalizeControlText(HtmlFormControlSemantics.GetOptionLabel(option));
            string export = HtmlFormControlSemantics.GetOptionValue(option);
            labels.Add(label);
            exports.Add(export);
            if (selectedOptions.Contains(option)) {
                if (selectedExports.Contains(export, StringComparer.Ordinal)) hasDuplicateSelectedValues = true;
                selectedExports.Add(export);
                selectedIndices.Add(index);
            }
        }
        for (int selectedIndex = 0; selectedIndex < selectedIndices.Count; selectedIndex++) {
            string selectedExport = exports[selectedIndices[selectedIndex]];
            if (exports.Count(export => string.Equals(export, selectedExport, StringComparison.Ordinal)) > 1) {
                hasAmbiguousSelectedValue = true;
                break;
            }
        }
        options = labels;
        optionValues = exports;
        values = selectedExports;
        selectedOptionIndices = selectedIndices;
    }

    private void ReportNonUniformFormFieldRadiusFallback(string source) {
        if (!_reportedNonUniformFormFieldRadiusFallbacks.Add(source)) {
            return;
        }

        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldNonUniformRadiusStaticFallback,
            "An HTML form control with non-uniform or elliptical corner radii was rendered as static content because PDF widgets currently preserve uniform circular radii.",
            HtmlDiagnosticSeverity.Warning,
            source,
            null,
            OfficeConversionLossKind.Approximation);
    }
}
