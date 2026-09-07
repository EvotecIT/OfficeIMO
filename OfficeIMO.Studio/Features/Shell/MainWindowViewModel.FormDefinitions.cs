using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    [ObservableProperty] private string _formDefinitionName = string.Empty;
    [ObservableProperty] private bool _formDefinitionRequired;
    [ObservableProperty] private bool _formDefinitionReadOnly;
    [ObservableProperty] private int _formTabOrderPage = 1;
    [ObservableProperty] private PdfPageTabOrder _formTabOrder = PdfPageTabOrder.Row;
    [ObservableProperty] private bool _formHasDefaultValue;
    [ObservableProperty] private string _formDefaultValue = string.Empty;
    [ObservableProperty] private int _formWidgetPage = 1;
    [ObservableProperty] private double _formWidgetX;
    [ObservableProperty] private double _formWidgetY;
    [ObservableProperty] private double _formWidgetWidth = 180;
    [ObservableProperty] private double _formWidgetHeight = 28;

    public IReadOnlyList<PdfPageTabOrder> FormTabOrders { get; } =
        [PdfPageTabOrder.Row, PdfPageTabOrder.Column, PdfPageTabOrder.Structure, PdfPageTabOrder.Annotations];
    public bool CanEditFormDefinition => CanAuthorForms && !IsWorkspaceBusy && !HasFormDrafts && SelectedFormField is not null;
    public bool CanSetFormTabOrder => CanAuthorForms && !IsWorkspaceBusy;
    public bool CanMoveFormWidget => CanEditFormDefinition && SelectedFormField?.SingleWidget?.ObjectNumber is not null;
    public bool CanSetFormDefault => CanEditFormDefinition && SelectedFormField is { IsSignature: false, IsUnsupported: false };

    [RelayCommand]
    private void ResetFormDefinition() {
        FormDefinitionName = SelectedFormField?.Name ?? string.Empty;
        FormDefinitionRequired = SelectedFormField?.IsRequired == true;
        FormDefinitionReadOnly = SelectedFormField?.IsReadOnly == true;
        FormHasDefaultValue = SelectedFormField?.SavedDefaultValue is not null;
        FormDefaultValue = SelectedFormField?.SavedDefaultValue ?? string.Empty;
        var widget = SelectedFormField?.SingleWidget;
        FormWidgetPage = widget?.PageNumber ?? 1;
        FormWidgetX = widget?.X1 ?? 0;
        FormWidgetY = widget?.Y1 ?? 0;
        FormWidgetWidth = widget?.Width ?? 180;
        FormWidgetHeight = widget?.Height ?? 28;
    }

    [RelayCommand]
    private async Task ApplyFormDefinitionAsync(CancellationToken token) {
        if (!CanEditFormDefinition || _workspace is null || SelectedFormField is null) return;
        string name = SelectedFormField.Name;
        string newName = FormDefinitionName.Trim();
        bool required = FormDefinitionRequired, readOnly = FormDefinitionReadOnly;
        bool applied = await RunMutationAsync(cancellation => _workspace.UpdateFormDefinitionAsync(
            name, newName, required, readOnly, cancellation, CreateProgress()), token).ConfigureAwait(true);
        if (applied) SelectedFormField = FormFields.FirstOrDefault(field => field.Name == newName);
    }

    [RelayCommand]
    private async Task ApplyFormTabOrderAsync(CancellationToken token) {
        if (!CanSetFormTabOrder || _workspace is null) return;
        int page = FormTabOrderPage;
        PdfPageTabOrder order = FormTabOrder;
        await RunMutationAsync(cancellation => _workspace.SetFormTabOrderAsync(page, order, cancellation, CreateProgress()), token).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task MoveFormWidgetAsync(CancellationToken token) {
        if (!CanMoveFormWidget || _workspace is null || SelectedFormField is null) return;
        string name = SelectedFormField.Name;
        int page = FormWidgetPage;
        double x = FormWidgetX, y = FormWidgetY, width = FormWidgetWidth, height = FormWidgetHeight;
        if (await RunMutationAsync(cancellation => _workspace.MoveFormWidgetAsync(name, page, x, y, width, height,
            cancellation, CreateProgress()), token).ConfigureAwait(true)) ShowSelectedFormField();
    }

    [RelayCommand]
    private async Task RemoveFormDefinitionAsync(CancellationToken token) {
        if (!CanEditFormDefinition || _workspace is null || SelectedFormField is null) return;
        string name = SelectedFormField.Name;
        await RunMutationAsync(cancellation => _workspace.RemoveFormDefinitionAsync(name, cancellation, CreateProgress()), token).ConfigureAwait(true);
    }

    [RelayCommand]
    private async Task SetFormDefaultAsync(CancellationToken token) {
        if (!CanSetFormDefault || _workspace is null || SelectedFormField is null) return;
        string name = SelectedFormField.Name;
        string? value = FormHasDefaultValue ? FormDefaultValue : null;
        await RunMutationAsync(cancellation => _workspace.SetFormDefaultAsync(name, value, cancellation, CreateProgress()), token).ConfigureAwait(true);
    }

    private PdfFormFieldCreateOptions CaptureNewFormFieldOptions() {
        string[] options = ParseFormFieldOptions(NewFormFieldOptions);
        PdfFormFieldCreationKind kind = SelectedFormFieldCreationChoice.Kind;
        bool allowsMultipleSelection = kind == PdfFormFieldCreationKind.Choice && NewFormFieldAllowsMultipleSelection;
        bool isComboBox = kind == PdfFormFieldCreationKind.Choice && NewFormFieldIsComboBox && !allowsMultipleSelection;
        double height = kind == PdfFormFieldCreationKind.RadioButtonGroup
            ? Math.Max(NewFormFieldHeight, GetRequiredRadioGroupHeight(options.Length))
            : NewFormFieldHeight;
        return new PdfFormFieldCreateOptions {
            Name = NewFormFieldName?.Trim() ?? string.Empty,
            Kind = kind,
            PageNumber = NewFormFieldPageNumber,
            X = NewFormFieldX,
            Y = NewFormFieldY,
            Width = NewFormFieldWidth,
            Height = height,
            Value = kind == PdfFormFieldCreationKind.CheckBox
                ? NewFormFieldIsChecked ? "Yes" : "Off"
                : NewFormFieldValue ?? string.Empty,
            ChoiceOptions = options,
            IsComboBox = isComboBox,
            Caption = string.IsNullOrWhiteSpace(NewFormFieldCaption) ? "Button" : NewFormFieldCaption.Trim(),
            FieldFlags = allowsMultipleSelection ? 2097152 : 0,
            Style = new PdfFormFieldStyle {
                AlternateName = string.IsNullOrWhiteSpace(NewFormFieldDisplayName) ? null : NewFormFieldDisplayName.Trim(),
                IsRequired = NewFormFieldIsRequired,
                IsReadOnly = NewFormFieldIsReadOnly,
                MaxLength = kind == PdfFormFieldCreationKind.Text && NewFormFieldMaxLength > 0 ? NewFormFieldMaxLength : null,
                IsMultiline = kind == PdfFormFieldCreationKind.Text && NewFormFieldIsMultiline,
                IsPassword = kind == PdfFormFieldCreationKind.Text && NewFormFieldIsPassword,
                IsEditableChoice = isComboBox && NewFormFieldAllowsCustomValue
            }
        };
    }
}
