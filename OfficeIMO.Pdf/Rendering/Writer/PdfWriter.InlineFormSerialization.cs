using System.Runtime.CompilerServices;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>One logical source field may have several widgets after line and page layout.</summary>
    private sealed class InlineFormSerializationPlan {
        public PdfInlineFormField Descriptor { get; }
        public FormFieldAnnotation FirstWidget { get; }
        public int ParentFieldId { get; }
        public List<int> WidgetObjectIds { get; } = new();
        public IReadOnlyList<PdfFormFieldOption> ChoiceOptions { get; private set; } = Array.Empty<PdfFormFieldOption>();
        public IReadOnlyList<string> ChoiceValues { get; private set; } = Array.Empty<string>();
        public IReadOnlyList<int> ChoiceSelectedIndices { get; private set; } = Array.Empty<int>();
        public int? ChoiceTopIndex { get; private set; }

        public InlineFormSerializationPlan(PdfInlineFormField descriptor, FormFieldAnnotation firstWidget, int parentFieldId) {
            Descriptor = descriptor;
            FirstWidget = firstWidget;
            ParentFieldId = parentFieldId;
        }

        public void CaptureChoiceData(IReadOnlyList<PdfFormFieldOption> options, IReadOnlyList<string> values, IReadOnlyList<int> selectedIndices, int? topIndex) {
            if (ChoiceOptions.Count != 0) return;
            ChoiceOptions = options;
            ChoiceValues = values;
            ChoiceSelectedIndices = selectedIndices;
            ChoiceTopIndex = topIndex;
        }
    }

    private sealed class InlineFormIdentityComparer : IEqualityComparer<PdfInlineFormField> {
        public bool Equals(PdfInlineFormField? x, PdfInlineFormField? y) => ReferenceEquals(x, y);
        public int GetHashCode(PdfInlineFormField obj) => RuntimeHelpers.GetHashCode(obj);
    }

    private static bool IsSharedInlineFormField(FormFieldAnnotation field) =>
        field.InlineField != null &&
        (field.Kind == FormFieldAnnotationKind.Text || field.Kind == FormFieldAnnotationKind.Choice);

    private static bool IsInlineFormContinuation(FormFieldAnnotation first, FormFieldAnnotation next) =>
        IsSharedInlineFormField(first) && first.Kind == next.Kind &&
        ReferenceEquals(first.InlineField, next.InlineField);

    private static InlineFormSerializationPlan? GetInlineFormSerializationPlan(FormFieldAnnotation field,
        IDictionary<PdfInlineFormField, InlineFormSerializationPlan> plans, IList<byte[]> objects, List<int> formFieldIds) {
        if (!IsSharedInlineFormField(field)) return null;
        PdfInlineFormField descriptor = field.InlineField!;
        PdfCanvasFormFieldKind expectedKind = field.Kind == FormFieldAnnotationKind.Text
            ? PdfCanvasFormFieldKind.Text : PdfCanvasFormFieldKind.Choice;
        if (descriptor.Field.Kind != expectedKind || !string.Equals(descriptor.Field.Name, field.Name, StringComparison.Ordinal)) {
            throw new InvalidOperationException("Inline widget kind and name must match its source field descriptor.");
        }

        if (!plans.TryGetValue(descriptor, out InlineFormSerializationPlan? plan)) {
            plan = new InlineFormSerializationPlan(descriptor, field, ReserveObject(objects));
            plans.Add(descriptor, plan);
            formFieldIds.Add(plan.ParentFieldId);
        }
        return plan;
    }

    private static string BuildInlineFormParentDictionary(InlineFormSerializationPlan plan) {
        PdfCanvasFormFieldItem source = plan.Descriptor.Field;
        FormFieldAnnotation first = plan.FirstWidget;
        PdfFormFieldStyle style = ScaleFormWidgetStyle(source.Style, first.AppearanceScale);
        double fontSize = source.FontSize * first.AppearanceScale;
        if (source.Kind == PdfCanvasFormFieldKind.Text) {
            return PdfAnnotationDictionaryBuilder.BuildTextFieldDictionary(source.Name, source.Value, fontSize, plan.WidgetObjectIds, style);
        }

        return source.ChoiceOptions.Count > 0
            ? PdfAnnotationDictionaryBuilder.BuildChoiceFieldDictionary(source.Name, plan.ChoiceOptions, plan.ChoiceValues,
                fontSize, source.IsComboBox, source.AllowsMultipleSelection, plan.WidgetObjectIds, style,
                plan.ChoiceSelectedIndices, plan.ChoiceTopIndex)
            : PdfAnnotationDictionaryBuilder.BuildChoiceFieldDictionary(source.Name, source.Options, plan.ChoiceValues,
                fontSize, source.IsComboBox, source.AllowsMultipleSelection, plan.WidgetObjectIds, style,
                plan.ChoiceSelectedIndices, plan.ChoiceTopIndex);
    }
}
