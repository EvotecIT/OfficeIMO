using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Editor;

public sealed partial class PdfFormFieldViewModel : ObservableObject {
    private readonly string _checkedValue;
    private readonly IStudioLocalizer _localizer;
    private string[] _savedValues;
    private readonly PdfFormField _field;

    internal PdfFormFieldViewModel(PdfFormField field, IStudioLocalizer? localizer = null) {
        ArgumentNullException.ThrowIfNull(field);
        _field = field;
        _localizer = localizer ?? new StudioLocalizer(System.Globalization.CultureInfo.GetCultureInfo("en"));
        Name = field.Name ?? throw new ArgumentException("A named form field is required.", nameof(field));
        Kind = GetKindLabel(field, _localizer);
        IsReadOnly = field.IsReadOnly;
        IsRequired = field.IsRequired;
        PageNumbers = field.PageNumbers;
        IsTextEditor = field.IsTextField;
        IsMultiline = field.IsMultiline;
        IsPassword = field.IsPassword;
        IsCheckBoxEditor = field.IsCheckBox;
        IsChoiceEditor = field.IsChoiceField || field.IsRadioButton;
        IsMultipleChoiceEditor = field.IsChoiceField && field.AllowsMultipleSelection;
        IsSingleChoiceEditor = IsChoiceEditor && !IsMultipleChoiceEditor;
        IsEditableChoice = field.IsChoiceField && field.IsEditableChoice;
        IsSignature = field.IsSignatureField;
        IsUnsupported = field.Kind == PdfFormFieldKind.Unknown || field.IsPushButton;
        TextValue = field.Value ?? string.Empty;

        string[] selectedValues = field.Values.Count > 0
            ? field.Values.ToArray()
            : string.IsNullOrEmpty(field.Value) ? [] : [field.Value];
        IReadOnlyList<PdfFormChoiceOption> sourceOptions = GetOptions(field);
        foreach (PdfFormChoiceOption option in sourceOptions) {
            var choice = new PdfFormChoiceViewModel(
                option.ExportValue,
                option.DisplayText,
                selectedValues.Contains(option.ExportValue, StringComparer.Ordinal));
            Choices.Add(choice);
        }

        _selectedChoice = Choices.FirstOrDefault(static choice => choice.IsSelected);
        if (IsSingleChoiceEditor && !IsEditableChoice && _selectedChoice is null &&
            !string.IsNullOrEmpty(field.Value) && !(field.IsRadioButton && field.Value == "Off")) {
            _selectedChoice = new PdfFormChoiceViewModel(field.Value, field.Value, true);
            Choices.Add(_selectedChoice);
        }
        _checkedValue = GetCheckedValue(field);
        IsChecked = field.IsCheckBox &&
                    !string.IsNullOrWhiteSpace(field.Value) &&
                    !string.Equals(field.Value, "Off", StringComparison.OrdinalIgnoreCase);
        _savedValues = CreateValue().Values.ToArray();
        foreach (var choice in Choices) choice.PropertyChanged += (_, _) => NotifyValueChanged();
    }

    public string Name { get; }

    internal PdfFormWidget? SingleWidget => _field.Widgets.Count == 1 ? _field.Widgets[0] : null;
    internal string? SavedDefaultValue => _field.DefaultValue;

    internal bool HasSameDefinition(PdfFormFieldViewModel other) => Name == other.Name &&
        _field.Flags == other._field.Flags && _field.DefaultValues.SequenceEqual(other._field.DefaultValues, StringComparer.Ordinal) &&
        _field.Widgets.Select(widget => (widget.ObjectNumber, widget.PageNumber, widget.X1, widget.Y1, widget.X2, widget.Y2))
            .SequenceEqual(other._field.Widgets.Select(widget => (widget.ObjectNumber, widget.PageNumber, widget.X1, widget.Y1, widget.X2, widget.Y2)));

    public string Kind { get; }

    public bool IsReadOnly { get; }

    public bool IsRequired { get; }

    public IReadOnlyList<int> PageNumbers { get; }

    public bool IsTextEditor { get; }

    public bool IsPlainTextEditor => IsTextEditor && !IsPassword;

    public bool IsMultiline { get; }

    public bool IsPassword { get; }

    public bool IsCheckBoxEditor { get; }

    public bool IsChoiceEditor { get; }

    public bool IsSingleChoiceEditor { get; }

    public bool IsMultipleChoiceEditor { get; }

    public bool IsEditableChoice { get; }

    public bool IsSignature { get; }

    public bool IsUnsupported { get; }

    public bool CanFill => !IsReadOnly && !IsSignature && !IsUnsupported;

    public string DisplayName => string.IsNullOrWhiteSpace(_field.AlternateName) ? Name : _field.AlternateName;
    public string FieldState => string.Join(" · ", new[] {
        IsReadOnly ? _localizer.GetOrDefault("FormField.ReadOnly", "Read-only") : null,
        IsRequired ? _localizer.GetOrDefault("FormField.Required", "Required") : null,
        _field.MaxLength is int maximum ? _localizer.FormatOrDefault("FormField.MaximumLength", "Up to {0} characters", maximum) : null
    }.OfType<string>());
    public PdfFormFieldValueAssessment ValueAssessment => PdfFormFieldValueAssessment.Assess(_field, CreateValue());
    public string ValidationSummary => string.Join(" ", ValueAssessment.Issues.Select(issue => issue.Message));
    public bool HasValidationMessage => ValidationSummary.Length != 0;
    public bool CanApplyValue => CanFill && !HasDraftConflict && !ValueAssessment.HasErrors;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanApplyValue))]
    private bool _hasDraftConflict;

    public bool HasDraft => !_savedValues.SequenceEqual(CreateValue().Values, StringComparer.Ordinal);
    public string DraftText => string.Join(Environment.NewLine, CreateValue().Values);
    public char DraftPasswordChar => IsPassword ? '●' : '\0';

    internal void RecordAppliedValue(PdfFormFieldValue value) {
        _savedValues = value.Values.ToArray();
        NotifyValueChanged();
    }

    internal void PreserveDraft(PdfFormFieldViewModel previous) {
        if (!previous.HasDraft || Kind != previous.Kind ||
            _savedValues.SequenceEqual(previous.CreateValue().Values, StringComparer.Ordinal)) return;
        HasDraftConflict = previous.HasDraftConflict ||
            !_savedValues.SequenceEqual(previous._savedValues, StringComparer.Ordinal) ||
            _field.Flags != previous._field.Flags || _field.MaxLength != previous._field.MaxLength ||
            !Choices.Select(choice => choice.ExportValue).SequenceEqual(previous.Choices.Select(choice => choice.ExportValue), StringComparer.Ordinal);
        foreach (var old in previous.Choices.Where(choice => choice.IsSelected)) {
            if (!Choices.Any(choice => choice.ExportValue == old.ExportValue)) {
                var retained = new PdfFormChoiceViewModel(old.ExportValue, old.DisplayText, false);
                retained.PropertyChanged += (_, _) => NotifyValueChanged();
                Choices.Add(retained);
            }
        }
        IsChecked = previous.IsChecked;
        SelectedChoice = Choices.FirstOrDefault(choice => choice.ExportValue == previous.SelectedChoice?.ExportValue);
        TextValue = previous.TextValue;
        foreach (var choice in Choices) choice.IsSelected = previous.Choices.Any(old => old.IsSelected && old.ExportValue == choice.ExportValue);
    }

    [RelayCommand]
    private void KeepDraft() => HasDraftConflict = false;

    [RelayCommand]
    private void ResetDraft() {
        SelectedChoice = Choices.FirstOrDefault(choice => _savedValues.Contains(choice.ExportValue, StringComparer.Ordinal));
        foreach (var choice in Choices) choice.IsSelected = _savedValues.Contains(choice.ExportValue, StringComparer.Ordinal);
        TextValue = _savedValues.FirstOrDefault() ?? string.Empty;
        IsChecked = _field.IsCheckBox && !string.IsNullOrEmpty(TextValue) && TextValue != "Off";
        HasDraftConflict = false;
        NotifyValueChanged();
    }

    public string Label => IsReadOnly
        ? _localizer.Format("FormField.ReadOnlyLabel", DisplayName, Kind)
        : _localizer.Format("FormField.Label", DisplayName, Kind);

    public string PageLabel => PageNumbers.Count switch {
        0 => _localizer.Get("FormField.NoPageLocation"),
        1 => _localizer.Format("PdfPage.Label", PageNumbers[0]),
        _ => _localizer.Format("FormField.PagesLabel", string.Join(", ", PageNumbers))
    };

    public ObservableCollection<PdfFormChoiceViewModel> Choices { get; } = new();

    [ObservableProperty]
    private string _textValue = string.Empty;

    [ObservableProperty]
    private bool _isChecked;

    [ObservableProperty]
    private PdfFormChoiceViewModel? _selectedChoice;

    partial void OnSelectedChoiceChanged(PdfFormChoiceViewModel? value) {
        foreach (PdfFormChoiceViewModel choice in Choices) choice.IsSelected = ReferenceEquals(choice, value);
        if (IsEditableChoice && value is not null) TextValue = value.ExportValue;
        NotifyValueChanged();
    }

    partial void OnTextValueChanged(string value) => NotifyValueChanged();
    partial void OnIsCheckedChanged(bool value) => NotifyValueChanged();
    private void NotifyValueChanged() {
        OnPropertyChanged(nameof(HasDraft));
        OnPropertyChanged(nameof(DraftText));
        OnPropertyChanged(nameof(ValueAssessment));
        OnPropertyChanged(nameof(ValidationSummary));
        OnPropertyChanged(nameof(HasValidationMessage));
        OnPropertyChanged(nameof(CanApplyValue));
    }

    internal PdfFormFieldValue CreateValue() {
        if (IsCheckBoxEditor) return PdfFormFieldValue.From(IsChecked ? _checkedValue : "Off");
        if (IsMultipleChoiceEditor) {
            string[] values = Choices.Where(static choice => choice.IsSelected)
                .Select(static choice => choice.ExportValue)
                .ToArray();
            return values.Length == 0 ? PdfFormFieldValue.From(string.Empty) : PdfFormFieldValue.FromValues(values);
        }
        if (IsChoiceEditor) {
            string value = IsEditableChoice
                ? TextValue
                : SelectedChoice?.ExportValue ?? string.Empty;
            return PdfFormFieldValue.From(value);
        }
        return PdfFormFieldValue.From(TextValue ?? string.Empty);
    }

    private static string GetKindLabel(PdfFormField field, IStudioLocalizer localizer) {
        string key = field switch {
            { IsPassword: true } => "Password",
            { IsMultiline: true } => "MultilineText",
            { IsTextField: true } => "Text",
            { IsCheckBox: true } => "CheckBox",
            { IsRadioButton: true } => "RadioGroup",
            { IsPushButton: true } => "Button",
            { IsCombo: true, IsEditableChoice: true } => "EditableChoice",
            { IsCombo: true } => "DropDown",
            { IsChoiceField: true, AllowsMultipleSelection: true } => "MultipleChoice",
            { IsChoiceField: true } => "Choice",
            { IsSignatureField: true } => "Signature",
            _ => "Unsupported"
        };
        return localizer.Get($"FormField.Kind.{key}");
    }

    private static IReadOnlyList<PdfFormChoiceOption> GetOptions(PdfFormField field) {
        if (field.IsChoiceField) {
            return field.Options.Select(static option => new PdfFormChoiceOption(option.ExportValue, option.DisplayText)).ToArray();
        }
        if (!field.IsRadioButton) return [];

        return field.Widgets
            .SelectMany(static widget => widget.NormalAppearanceStates)
            .Where(static value => !string.Equals(value, "Off", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.Ordinal)
            .Select(static value => new PdfFormChoiceOption(value, value))
            .ToArray();
    }

    private static string GetCheckedValue(PdfFormField field) => field.Widgets
        .SelectMany(static widget => widget.NormalAppearanceStates)
        .FirstOrDefault(static value => !string.Equals(value, "Off", StringComparison.OrdinalIgnoreCase))
        ?? "Yes";

    private sealed record PdfFormChoiceOption(string ExportValue, string DisplayText);
}

public sealed partial class PdfFormChoiceViewModel : ObservableObject {
    internal PdfFormChoiceViewModel(string exportValue, string displayText, bool isSelected) {
        ExportValue = exportValue;
        DisplayText = displayText;
        _isSelected = isSelected;
    }

    public string ExportValue { get; }

    public string DisplayText { get; }

    [ObservableProperty]
    private bool _isSelected;
}

public sealed record PdfFormFieldCreationChoice(PdfFormFieldCreationKind Kind, string Label);
