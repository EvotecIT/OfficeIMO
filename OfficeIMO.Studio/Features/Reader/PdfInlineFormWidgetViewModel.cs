using Avalonia.Controls;
using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Features.Reader;

/// <summary>One on-page editor; its value belongs to the shared field, including repeated widgets.</summary>
public sealed partial class PdfInlineFormWidgetViewModel : ObservableObject {
    private readonly PdfPageInteractionRegion _region;
    internal PdfPageViewModel Page { get; }
    public PdfFormFieldViewModel Field { get; }
    public int? ObjectNumber => _region.ObjectNumber;
    public PdfFormChoiceViewModel? RadioChoice { get; }
    // Each visual radio owns its native group; the field owns mutual exclusion, including repeated values.
    public string RadioGroupName { get; } = Guid.NewGuid().ToString("N");
    public string RadioLabel => $"{Field.DisplayName}: {RadioChoice?.DisplayText}";
    public bool IsDropDown => Field.IsFixedSingleChoiceEditor && !Field.IsRadioButtonEditor && !Field.IsListBoxEditor;
    public SelectionMode ChoiceSelectionMode => Field.IsMultipleChoiceEditor ? SelectionMode.Multiple | SelectionMode.Toggle : SelectionMode.Single;

    [ObservableProperty] private double _left;
    [ObservableProperty] private double _top;
    [ObservableProperty] private double _width;
    [ObservableProperty] private double _height;
    [ObservableProperty] private double _fontSize;

    private PdfInlineFormWidgetViewModel(PdfPageViewModel page, PdfFormFieldViewModel field, PdfPageInteractionRegion region, PdfFormChoiceViewModel? radioChoice) {
        Page = page;
        Field = field;
        _region = region;
        RadioChoice = radioChoice;
    }

    internal static PdfInlineFormWidgetViewModel? Create(PdfPageViewModel page, PdfFormFieldViewModel field, PdfPageInteractionRegion region) {
        var widget = field.Widgets.FirstOrDefault(widget => widget.ObjectNumber == region.ObjectNumber && widget.PageNumber == page.PageNumber);
        if (widget is null || !CanShow(field, widget)) return null;
        PdfFormChoiceViewModel? radioChoice = null;
        if (field.IsRadioButtonEditor) {
            string[] states = widget.NormalAppearanceStates.Where(state => state != "Off").ToArray();
            // Ambiguous or missing appearance states remain available in the inspector.
            if (states.Length != 1) return null;
            radioChoice = field.Choices.FirstOrDefault(choice => choice.ExportValue == states[0]);
            if (radioChoice is null) return null;
        }
        return new(page, field, region, radioChoice);
    }

    internal static bool CanShow(PdfFormFieldViewModel field, PdfFormWidget widget) => widget.PageNumber is > 0 &&
        !widget.IsHidden && !widget.IsInvisible && !widget.IsNoView && !widget.IsReadOnly &&
        (!field.IsRadioButtonEditor || widget.NormalAppearanceStates.Count(state => state != "Off") == 1);

    internal void UpdateBounds(double scaleX, double scaleY) {
        bool button = Field.IsCheckBoxEditor || Field.IsRadioButtonEditor;
        Left = _region.Quad.Left * scaleX;
        Top = _region.Quad.Top * scaleY;
        Width = Math.Max(button ? 24D : 60D, _region.Quad.Width * scaleX);
        Height = Math.Max(button ? 24D : 22D, _region.Quad.Height * scaleY);
        FontSize = Math.Clamp(Height * (Field.IsMultiline ? 0.3D : 0.55D), 9D, 16D);
    }

    internal void Activate() {
        if (ReferenceEquals(Page.InlineFormField, Field) && Page.FormAnchorObjectNumber == ObjectNumber) return;
        Page.SelectObject(new(PdfEditorSelectionKind.FormField, Page.PageNumber,
            new(_region.Quad.Left, _region.Quad.Top, _region.Quad.Right, _region.Quad.Bottom), ObjectNumber: ObjectNumber, FieldName: Field.Name));
    }
}
