using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Editor;

public sealed partial class PdfFormFieldViewModel : ObservableObject {
    private readonly string _checkedValue;

    internal PdfFormFieldViewModel(PdfFormField field) {
        ArgumentNullException.ThrowIfNull(field);
        Name = field.Name ?? throw new ArgumentException("A named form field is required.", nameof(field));
        Kind = GetKindLabel(field);
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

        _selectedChoice = Choices.FirstOrDefault(static choice => choice.IsSelected)
            ?? Choices.FirstOrDefault();
        _checkedValue = GetCheckedValue(field);
        IsChecked = field.IsCheckBox &&
                    !string.IsNullOrWhiteSpace(field.Value) &&
                    !string.Equals(field.Value, "Off", StringComparison.OrdinalIgnoreCase);
    }

    public string Name { get; }

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

    public string Label => Name + " · " + Kind + (IsReadOnly ? " · read only" : string.Empty);

    public string PageLabel => PageNumbers.Count switch {
        0 => "No page location",
        1 => "Page " + PageNumbers[0].ToString(System.Globalization.CultureInfo.InvariantCulture),
        _ => "Pages " + string.Join(", ", PageNumbers)
    };

    public ObservableCollection<PdfFormChoiceViewModel> Choices { get; } = new();

    [ObservableProperty]
    private string _textValue = string.Empty;

    [ObservableProperty]
    private bool _isChecked;

    [ObservableProperty]
    private PdfFormChoiceViewModel? _selectedChoice;

    partial void OnSelectedChoiceChanged(PdfFormChoiceViewModel? value) {
        if (value is null) return;
        foreach (PdfFormChoiceViewModel choice in Choices) choice.IsSelected = ReferenceEquals(choice, value);
        if (IsEditableChoice) TextValue = value.ExportValue;
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

    private static string GetKindLabel(PdfFormField field) {
        if (field.IsPassword) return "Password";
        if (field.IsMultiline) return "Multiline text";
        if (field.IsTextField) return "Text";
        if (field.IsCheckBox) return "Check box";
        if (field.IsRadioButton) return "Radio group";
        if (field.IsPushButton) return "Button";
        if (field.IsCombo) return field.IsEditableChoice ? "Editable choice" : "Drop-down";
        if (field.IsChoiceField) return field.AllowsMultipleSelection ? "Multiple choice" : "Choice";
        if (field.IsSignatureField) return "Signature";
        return "Unsupported";
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
