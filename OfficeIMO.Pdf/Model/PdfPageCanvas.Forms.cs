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
        _items.Add(PdfCanvasFormFieldItem.Text(name, value ?? string.Empty, x, y, width, height, fontSize, style));
        return this;
    }

    /// <summary>Adds an interactive check box at fixed top-left page coordinates.</summary>
    public PdfPageCanvas CheckBox(string name, bool isChecked, double x, double y, double width, double height, string checkedValueName = "Yes", PdfFormFieldStyle? style = null) {
        ValidateFormFieldBox(name, x, y, width, height);
        Guard.NotNullOrWhiteSpace(checkedValueName, nameof(checkedValueName));
        _items.Add(PdfCanvasFormFieldItem.CheckBox(name, isChecked, checkedValueName, checkedValueName, x, y, width, height, style));
        return this;
    }

    /// <summary>Adds an interactive check box with separate PDF appearance-state and exported values.</summary>
    public PdfPageCanvas CheckBoxWithExportValue(string name, bool isChecked, double x, double y, double width, double height, string checkedValueName, string exportValue, PdfFormFieldStyle? style = null) {
        ValidateFormFieldBox(name, x, y, width, height);
        Guard.NotNullOrWhiteSpace(checkedValueName, nameof(checkedValueName));
        Guard.NotNullOrWhiteSpace(exportValue, nameof(exportValue));
        _items.Add(PdfCanvasFormFieldItem.CheckBox(name, isChecked, checkedValueName, exportValue, x, y, width, height, style));
        return this;
    }

    /// <summary>Adds an interactive choice field at fixed top-left page coordinates.</summary>
    public PdfPageCanvas ChoiceField(string name, IEnumerable<string> options, IEnumerable<string>? values, double x, double y, double width, double height, double fontSize = 10D, bool isComboBox = true, bool allowsMultipleSelection = false, PdfFormFieldStyle? style = null) {
        ValidateFormFieldBox(name, x, y, width, height);
        Guard.NotNull(options, nameof(options));
        Guard.Positive(fontSize, nameof(fontSize));
        var optionSnapshot = options.ToList();
        if (optionSnapshot.Count == 0 || optionSnapshot.Any(string.IsNullOrWhiteSpace)) {
            throw new ArgumentException("Canvas choice fields require at least one non-empty option.", nameof(options));
        }
        if (optionSnapshot.Distinct(StringComparer.Ordinal).Count() != optionSnapshot.Count) {
            throw new ArgumentException("Canvas choice field options must be unique.", nameof(options));
        }
        var valueSnapshot = values?.ToList() ?? new List<string>();
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
        Guard.NotNull(options, nameof(options));
        Guard.Positive(fontSize, nameof(fontSize));
        var optionSnapshot = options.ToList();
        if (optionSnapshot.Count == 0 || optionSnapshot.Any(option => option == null || string.IsNullOrWhiteSpace(option.DisplayText))) {
            throw new ArgumentException("Canvas choice fields require at least one option with non-empty display text.", nameof(options));
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

    internal static PdfCanvasFormFieldItem Text(string name, string value, double x, double y, double width, double height, double fontSize, PdfFormFieldStyle? style) =>
        new(PdfCanvasFormFieldKind.Text, name, x, y, width, height, style) { Value = value, FontSize = fontSize };

    internal static PdfCanvasFormFieldItem CheckBox(string name, bool isChecked, string checkedValueName, string exportValue, double x, double y, double width, double height, PdfFormFieldStyle? style) =>
        new(PdfCanvasFormFieldKind.CheckBox, name, x, y, width, height, style) { IsSelected = isChecked, Option = checkedValueName, ExportValue = exportValue, Value = isChecked ? checkedValueName : "Off" };

    internal static PdfCanvasFormFieldItem Choice(string name, IReadOnlyList<string> options, IReadOnlyList<string> values, double x, double y, double width, double height, double fontSize, bool isComboBox, bool allowsMultipleSelection, PdfFormFieldStyle? style) =>
        new(PdfCanvasFormFieldKind.Choice, name, x, y, width, height, style) { Options = options, Values = values, Value = values.Count == 0 ? string.Empty : values[0], FontSize = fontSize, IsComboBox = isComboBox, AllowsMultipleSelection = allowsMultipleSelection };

    internal static PdfCanvasFormFieldItem Choice(string name, IReadOnlyList<PdfFormFieldOption> options, IReadOnlyList<string> values, double x, double y, double width, double height, double fontSize, bool isComboBox, bool allowsMultipleSelection, PdfFormFieldStyle? style) =>
        new(PdfCanvasFormFieldKind.Choice, name, x, y, width, height, style) { ChoiceOptions = options, Values = values, Value = values.Count == 0 ? string.Empty : values[0], FontSize = fontSize, IsComboBox = isComboBox, AllowsMultipleSelection = allowsMultipleSelection };

    internal static PdfCanvasFormFieldItem RadioButton(string name, string option, string exportValue, bool isSelected, double x, double y, double width, double height, PdfFormFieldStyle? style) =>
        new(PdfCanvasFormFieldKind.RadioButton, name, x, y, width, height, style) { Option = option, ExportValue = exportValue, Value = option, IsSelected = isSelected };

    internal PdfCanvasFormFieldKind Kind { get; }
    internal string Name { get; }
    internal string Value { get; private set; } = string.Empty;
    internal IReadOnlyList<string> Values { get; private set; } = Array.Empty<string>();
    internal IReadOnlyList<string> Options { get; private set; } = Array.Empty<string>();
    internal IReadOnlyList<PdfFormFieldOption> ChoiceOptions { get; private set; } = Array.Empty<PdfFormFieldOption>();
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
