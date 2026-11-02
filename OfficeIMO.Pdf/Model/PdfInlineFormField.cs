namespace OfficeIMO.Pdf;

/// <summary>Shared identity and immutable source snapshot for a form control that follows rich text.</summary>
internal sealed class PdfInlineFormField {
    internal PdfCanvasFormFieldItem Field { get; }
    internal PdfTextRun? CheckedAppearanceRun { get; }
    internal PdfTextRun? UncheckedAppearanceRun { get; }

    internal PdfInlineFormField(PdfCanvasFormFieldItem field, PdfTextRun? checkedAppearanceRun = null,
        PdfTextRun? uncheckedAppearanceRun = null) {
        Guard.NotNull(field, nameof(field));
        if (field.Kind != PdfCanvasFormFieldKind.CheckBox &&
            (checkedAppearanceRun != null || uncheckedAppearanceRun != null))
            throw new ArgumentException("State glyph appearances require a check box field.", nameof(field));
        if (checkedAppearanceRun?.BackgroundColor != null || uncheckedAppearanceRun?.BackgroundColor != null)
            throw new NotSupportedException("An inline state appearance cannot preserve a run background beneath surrounding page text. Keep the source run static.");
        CheckedAppearanceRun = checkedAppearanceRun?.SnapshotForInlineAppearance();
        UncheckedAppearanceRun = uncheckedAppearanceRun?.SnapshotForInlineAppearance();
        Field = field.Kind switch {
            PdfCanvasFormFieldKind.Text => PdfCanvasFormFieldItem.Text(field.Name, field.Value,
                field.AppearanceValue, field.X, field.Y, field.Width, field.Height, field.FontSize,
                field.Style, field.AppearanceStyle),
            PdfCanvasFormFieldKind.CheckBox => PdfCanvasFormFieldItem.CheckBox(field.Name, field.IsSelected,
                field.Option, field.ExportValue, field.X, field.Y, field.Width, field.Height, field.Style),
            PdfCanvasFormFieldKind.Choice when field.ChoiceOptions.Count > 0 => PdfCanvasFormFieldItem.Choice(
                field.Name, field.ChoiceOptions.ToArray(), field.Values.ToArray(), field.SelectedIndices.ToArray(),
                field.X, field.Y, field.Width, field.Height, field.FontSize, field.IsComboBox,
                field.AllowsMultipleSelection, field.Style),
            PdfCanvasFormFieldKind.Choice => PdfCanvasFormFieldItem.Choice(field.Name,
                field.Options.ToArray(), field.Values.ToArray(), field.X, field.Y, field.Width, field.Height,
                field.FontSize, field.IsComboBox, field.AllowsMultipleSelection, field.Style),
            _ => throw new ArgumentException("Inline form controls support text, choice and check box fields.", nameof(field))
        };
    }
}

/// <summary>A fixed-size control glyph; text controls retain ordinary text wrapping instead.</summary>
internal sealed class PdfInlineFormWidget : PdfInlineElement {
    internal PdfInlineFormField Field { get; }
    internal PdfTextRun? SourceRun { get; }

    internal PdfInlineFormWidget(PdfInlineFormField field, double width, double height, double baselineOffset = 0D,
        PdfTextRun? sourceRun = null)
        : base(global::OfficeIMO.OfficeLength.Points(width), global::OfficeIMO.OfficeLength.Points(height), global::OfficeIMO.OfficeLength.Points(baselineOffset), null) {
        Guard.NotNull(field, nameof(field));
        if (sourceRun?.InlineElement != null)
            throw new ArgumentException("A form widget's source line metrics require a text run.", nameof(sourceRun));
        if (sourceRun?.BackgroundColor != null)
            throw new NotSupportedException("An inline widget cannot preserve a run background beneath surrounding page text. Keep the source run static.");
        Field = field;
        SourceRun = sourceRun;
    }
}
