namespace OfficeIMO.Pdf;

public sealed partial class PdfPageCanvas {
    internal PdfPageCanvas SearchableText(string text, double x, double y) {
        Guard.NotNullOrWhiteSpace(text, nameof(text));
        ValidateCanvasCoordinate(x, nameof(x));
        ValidateCanvasCoordinate(y, nameof(y));
        _items.Add(new PdfCanvasSearchableTextItem(text, x, y));
        return this;
    }

    /// <summary>Adds an interactive text field at fixed top-left page coordinates.</summary>
    public PdfPageCanvas TextField(string name, string? value, double x, double y, double width, double height, double fontSize = 10D, PdfFormFieldStyle? style = null) {
        ValidateFormFieldBox(name, x, y, width, height);
        Guard.Positive(fontSize, nameof(fontSize));
        string resolvedValue = value ?? string.Empty;
        _items.Add(PdfCanvasFormFieldItem.Text(name, resolvedValue, resolvedValue, x, y, width, height, fontSize, style, style));
        return this;
    }

    internal PdfPageCanvas TextFieldWithInitialAppearance(string name, string? value, string appearanceValue, double x, double y, double width, double height, double fontSize, PdfFormFieldStyle? style, PdfFormFieldStyle? appearanceStyle) {
        ValidateFormFieldBox(name, x, y, width, height);
        Guard.NotNull(appearanceValue, nameof(appearanceValue));
        Guard.Positive(fontSize, nameof(fontSize));
        _items.Add(PdfCanvasFormFieldItem.Text(name, value ?? string.Empty, appearanceValue, x, y, width, height, fontSize, style, appearanceStyle));
        return this;
    }

    internal PdfPageCanvas ChoiceFieldWithSelectedIndices(string name, IEnumerable<PdfFormFieldOption> options, IEnumerable<string>? values, IEnumerable<int>? selectedIndices, double x, double y, double width, double height, double fontSize, bool isComboBox, bool allowsMultipleSelection, PdfFormFieldStyle? style) {
        ValidateFormFieldBox(name, x, y, width, height);
        ValidateChoiceFieldMode(isComboBox, allowsMultipleSelection);
        Guard.NotNull(options, nameof(options));
        Guard.Positive(fontSize, nameof(fontSize));
        var optionSnapshot = options.ToList();
        if (optionSnapshot.Count == 0 || optionSnapshot.Any(option => option == null || string.IsNullOrWhiteSpace(option.DisplayText))) {
            throw new ArgumentException("Canvas choice fields require at least one option with non-empty display text.", nameof(options));
        }
        var valueSnapshot = values?.ToList() ?? new List<string>();
        var indexSnapshot = selectedIndices?.ToList() ?? new List<int>();
        if (!allowsMultipleSelection && valueSnapshot.Count > 1) {
            throw new ArgumentException("A single-select canvas choice field accepts at most one value.", nameof(values));
        }
        if (!allowsMultipleSelection && indexSnapshot.Count > 1) {
            throw new ArgumentException("A single-select canvas choice field accepts at most one selected index.", nameof(selectedIndices));
        }
        if (indexSnapshot.Any(index => index < 0 || index >= optionSnapshot.Count) || indexSnapshot.Distinct().Count() != indexSnapshot.Count) {
            throw new ArgumentException("Canvas choice field selected indices must be unique and refer to provided options.", nameof(selectedIndices));
        }
        if (valueSnapshot.Any(value => !optionSnapshot.Any(option => string.Equals(option.ExportValue, value, StringComparison.Ordinal)))) {
            throw new ArgumentException("Canvas choice field values must match provided export values.", nameof(values));
        }
        if (indexSnapshot.Count > 0 && (indexSnapshot.Count != valueSnapshot.Count || indexSnapshot.Where((index, valueIndex) => !string.Equals(optionSnapshot[index].ExportValue, valueSnapshot[valueIndex], StringComparison.Ordinal)).Any())) {
            throw new ArgumentException("Canvas choice field selected indices must identify the selected export values in the same order.", nameof(selectedIndices));
        }
        _items.Add(PdfCanvasFormFieldItem.Choice(name, optionSnapshot, valueSnapshot, indexSnapshot, x, y, width, height, fontSize, isComboBox, allowsMultipleSelection, style));
        return this;
    }

    /// <summary>Adds an interactive check box at fixed top-left page coordinates.</summary>
    public PdfPageCanvas CheckBox(string name, bool isChecked, double x, double y, double width, double height, string checkedValueName = "Yes", PdfFormFieldStyle? style = null) {
        ValidateFormFieldBox(name, x, y, width, height);
        ValidateCheckBoxAppearanceStateName(checkedValueName, nameof(checkedValueName));
        _items.Add(PdfCanvasFormFieldItem.CheckBox(name, isChecked, checkedValueName, checkedValueName, x, y, width, height, style));
        return this;
    }

    /// <summary>Adds an interactive check box with separate PDF appearance-state and exported values.</summary>
    public PdfPageCanvas CheckBoxWithExportValue(string name, bool isChecked, double x, double y, double width, double height, string checkedValueName, string exportValue, PdfFormFieldStyle? style = null) {
        ValidateFormFieldBox(name, x, y, width, height);
        ValidateCheckBoxAppearanceStateName(checkedValueName, nameof(checkedValueName));
        Guard.NotNullOrWhiteSpace(exportValue, nameof(exportValue));
        _items.Add(PdfCanvasFormFieldItem.CheckBox(name, isChecked, checkedValueName, exportValue, x, y, width, height, style));
        return this;
    }

    /// <summary>Adds an interactive choice field at fixed top-left page coordinates.</summary>
    public PdfPageCanvas ChoiceField(string name, IEnumerable<string> options, IEnumerable<string>? values, double x, double y, double width, double height, double fontSize = 10D, bool isComboBox = true, bool allowsMultipleSelection = false, PdfFormFieldStyle? style = null) {
        ValidateFormFieldBox(name, x, y, width, height);
        ValidateChoiceFieldMode(isComboBox, allowsMultipleSelection);
        Guard.NotNull(options, nameof(options));
        Guard.Positive(fontSize, nameof(fontSize));
        var optionSnapshot = options.ToList();
        if (optionSnapshot.Count == 0 || optionSnapshot.Any(string.IsNullOrWhiteSpace)) {
            throw new ArgumentException("Canvas choice fields require at least one non-empty option.", nameof(options));
        }
        if (optionSnapshot.Distinct(StringComparer.Ordinal).Count() != optionSnapshot.Count) {
            throw new ArgumentException("Canvas choice field options must be unique.", nameof(options));
        }
        var valueSnapshot = values?.Distinct(StringComparer.Ordinal).ToList() ?? new List<string>();
        if (!allowsMultipleSelection && valueSnapshot.Count > 1) {
            throw new ArgumentException("A single-select canvas choice field accepts at most one value.", nameof(values));
        }
        if (valueSnapshot.Any(value => !optionSnapshot.Contains(value, StringComparer.Ordinal))) {
            throw new ArgumentException("Canvas choice field values must match provided options.", nameof(values));
        }
        _items.Add(PdfCanvasFormFieldItem.Choice(name, optionSnapshot, valueSnapshot, x, y, width, height, fontSize, isComboBox, allowsMultipleSelection, style));
        return this;
    }

    /// <summary>Adds an interactive choice field whose export values differ from its displayed labels.</summary>
    public PdfPageCanvas ChoiceField(string name, IEnumerable<PdfFormFieldOption> options, IEnumerable<string>? values, double x, double y, double width, double height, double fontSize = 10D, bool isComboBox = true, bool allowsMultipleSelection = false, PdfFormFieldStyle? style = null) {
        ValidateFormFieldBox(name, x, y, width, height);
        ValidateChoiceFieldMode(isComboBox, allowsMultipleSelection);
        Guard.NotNull(options, nameof(options));
        Guard.Positive(fontSize, nameof(fontSize));
        var optionSnapshot = options.ToList();
        if (optionSnapshot.Count == 0 || optionSnapshot.Any(option => option == null || string.IsNullOrWhiteSpace(option.DisplayText))) {
            throw new ArgumentException("Canvas choice fields require at least one option with non-empty display text.", nameof(options));
        }
        if (optionSnapshot.Select(option => option.ExportValue).Distinct(StringComparer.Ordinal).Count() != optionSnapshot.Count) {
            throw new ArgumentException("Canvas choice field export values must be unique when selected values are provided without indices.", nameof(options));
        }
        var valueSnapshot = values?.Distinct(StringComparer.Ordinal).ToList() ?? new List<string>();
        if (!allowsMultipleSelection && valueSnapshot.Count > 1) {
            throw new ArgumentException("A single-select canvas choice field accepts at most one value.", nameof(values));
        }
        if (valueSnapshot.Any(value => !optionSnapshot.Any(option => string.Equals(option.ExportValue, value, StringComparison.Ordinal)))) {
            throw new ArgumentException("Canvas choice field values must match provided export values.", nameof(values));
        }
        _items.Add(PdfCanvasFormFieldItem.Choice(name, optionSnapshot, valueSnapshot, x, y, width, height, fontSize, isComboBox, allowsMultipleSelection, style));
        return this;
    }

    /// <summary>Adds one interactive radio-button widget. Widgets with the same field name are emitted as one radio group on the page.</summary>
    public PdfPageCanvas RadioButton(string name, string option, bool isSelected, double x, double y, double width, double height, PdfFormFieldStyle? style = null) {
        return RadioButtonWithExportValue(name, option, option, isSelected, x, y, width, height, style);
    }

    /// <summary>Adds one interactive radio-button widget with separate PDF appearance-state and exported values.</summary>
    internal PdfPageCanvas RadioButtonWithExportValue(string name, string option, string exportValue, bool isSelected, double x, double y, double width, double height, PdfFormFieldStyle? style = null) {
        ValidateFormFieldBox(name, x, y, width, height);
        Guard.NotNullOrWhiteSpace(option, nameof(option));
        Guard.NotNullOrWhiteSpace(exportValue, nameof(exportValue));
        if (string.Equals(option, "Off", StringComparison.Ordinal)) {
            throw new ArgumentException("PDF radio button option value cannot be Off.", nameof(option));
        }
        for (int index = 0; index < option.Length; index++) {
            if (option[index] > 0x7E) {
                throw new ArgumentException("PDF radio button option values must contain only ASCII PDF name characters.", nameof(option));
            }
        }
        _items.Add(PdfCanvasFormFieldItem.RadioButton(name, option, exportValue, isSelected, x, y, width, height, style));
        return this;
    }

    private void ValidateFormFieldBox(string name, double x, double y, double width, double height) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        ValidateCanvasCoordinate(x, nameof(x));
        ValidateCanvasCoordinate(y, nameof(y));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
    }

    private static void ValidateChoiceFieldMode(bool isComboBox, bool allowsMultipleSelection) {
        if (isComboBox && allowsMultipleSelection) {
            throw new ArgumentException("PDF multi-select choice fields must be list boxes, not combo boxes.", nameof(isComboBox));
        }
    }

    private static void ValidateCheckBoxAppearanceStateName(string value, string paramName) {
        Guard.NotNullOrWhiteSpace(value, paramName);
        if (string.Equals(value, "Off", StringComparison.Ordinal)) {
            throw new ArgumentException("Canvas check box selected value name cannot be Off.", paramName);
        }

        for (int index = 0; index < value.Length; index++) {
            if (value[index] > 0x7E) {
                throw new ArgumentException("Canvas check box selected value name must contain only ASCII PDF name characters.", paramName);
            }
        }
    }
}

internal sealed class PdfCanvasSearchableTextItem : PdfCanvasItem {
    internal PdfCanvasSearchableTextItem(string text, double x, double y)
        : base(x, y) {
        Text = text;
    }

    internal string Text { get; }
}

internal enum PdfCanvasFormFieldKind {
    Text,
    CheckBox,
    Choice,
    RadioButton
}

internal sealed class PdfCanvasFormFieldItem : PdfCanvasItem {
    private PdfCanvasFormFieldItem(PdfCanvasFormFieldKind kind, string name, double x, double y, double width, double height, PdfFormFieldStyle? style)
        : base(x, y) {
        Kind = kind;
        Name = name;
        Width = width;
        Height = height;
        Style = style?.Clone() ?? new PdfFormFieldStyle();
    }

    internal static PdfCanvasFormFieldItem Text(string name, string value, string appearanceValue, double x, double y, double width, double height, double fontSize, PdfFormFieldStyle? style, PdfFormFieldStyle? appearanceStyle) =>
        new(PdfCanvasFormFieldKind.Text, name, x, y, width, height, style) { Value = value, AppearanceValue = appearanceValue, AppearanceStyle = appearanceStyle?.Clone() ?? style?.Clone() ?? new PdfFormFieldStyle(), FontSize = fontSize };

    internal static PdfCanvasFormFieldItem CheckBox(string name, bool isChecked, string checkedValueName, string exportValue, double x, double y, double width, double height, PdfFormFieldStyle? style) =>
        new(PdfCanvasFormFieldKind.CheckBox, name, x, y, width, height, style) { IsSelected = isChecked, Option = checkedValueName, ExportValue = exportValue, Value = isChecked ? checkedValueName : "Off" };

    internal static PdfCanvasFormFieldItem Choice(string name, IReadOnlyList<string> options, IReadOnlyList<string> values, double x, double y, double width, double height, double fontSize, bool isComboBox, bool allowsMultipleSelection, PdfFormFieldStyle? style) =>
        new(PdfCanvasFormFieldKind.Choice, name, x, y, width, height, style) { Options = options, Values = values, Value = values.Count == 0 ? string.Empty : values[0], FontSize = fontSize, IsComboBox = isComboBox, AllowsMultipleSelection = allowsMultipleSelection };

    internal static PdfCanvasFormFieldItem Choice(string name, IReadOnlyList<PdfFormFieldOption> options, IReadOnlyList<string> values, double x, double y, double width, double height, double fontSize, bool isComboBox, bool allowsMultipleSelection, PdfFormFieldStyle? style) =>
        new(PdfCanvasFormFieldKind.Choice, name, x, y, width, height, style) { ChoiceOptions = options, Values = values, Value = values.Count == 0 ? string.Empty : values[0], FontSize = fontSize, IsComboBox = isComboBox, AllowsMultipleSelection = allowsMultipleSelection };

    internal static PdfCanvasFormFieldItem Choice(string name, IReadOnlyList<PdfFormFieldOption> options, IReadOnlyList<string> values, IReadOnlyList<int> selectedIndices, double x, double y, double width, double height, double fontSize, bool isComboBox, bool allowsMultipleSelection, PdfFormFieldStyle? style) =>
        new(PdfCanvasFormFieldKind.Choice, name, x, y, width, height, style) { ChoiceOptions = options, Values = values, SelectedIndices = selectedIndices, Value = values.Count == 0 ? string.Empty : values[0], FontSize = fontSize, IsComboBox = isComboBox, AllowsMultipleSelection = allowsMultipleSelection };

    internal static PdfCanvasFormFieldItem RadioButton(string name, string option, string exportValue, bool isSelected, double x, double y, double width, double height, PdfFormFieldStyle? style) =>
        new(PdfCanvasFormFieldKind.RadioButton, name, x, y, width, height, style) { Option = option, ExportValue = exportValue, Value = option, IsSelected = isSelected };

    internal PdfCanvasFormFieldKind Kind { get; }
    internal string Name { get; }
    internal string Value { get; private set; } = string.Empty;
    internal string AppearanceValue { get; private set; } = string.Empty;
    internal PdfFormFieldStyle AppearanceStyle { get; private set; } = new PdfFormFieldStyle();
    internal IReadOnlyList<string> Values { get; private set; } = Array.Empty<string>();
    internal IReadOnlyList<string> Options { get; private set; } = Array.Empty<string>();
    internal IReadOnlyList<PdfFormFieldOption> ChoiceOptions { get; private set; } = Array.Empty<PdfFormFieldOption>();
    internal IReadOnlyList<int> SelectedIndices { get; private set; } = Array.Empty<int>();
    internal string Option { get; private set; } = string.Empty;
    internal string ExportValue { get; private set; } = string.Empty;
    internal bool IsSelected { get; private set; }
    internal double Width { get; }
    internal double Height { get; }
    internal double FontSize { get; private set; }
    internal bool IsComboBox { get; private set; }
    internal bool AllowsMultipleSelection { get; private set; }
    internal PdfFormFieldStyle Style { get; }
}
