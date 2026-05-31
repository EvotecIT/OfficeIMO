namespace OfficeIMO.Pdf;

internal sealed class ChoiceFieldBlock : IPdfBlock {
    public string Name { get; }
    public IReadOnlyList<string> Options { get; }
    public string Value { get; }
    public IReadOnlyList<string> Values { get; }
    public double Width { get; }
    public double Height { get; }
    public PdfAlign Align { get; }
    public double FontSize { get; }
    public double SpacingBefore { get; }
    public double SpacingAfter { get; }
    public bool IsComboBox { get; }
    public bool AllowsMultipleSelection { get; }
    public PdfFormFieldStyle Style { get; }

    public ChoiceFieldBlock(string name, IEnumerable<string> options, string? value, double width, double height, PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, bool isComboBox, PdfFormFieldStyle? style = null)
        : this(name, options, value is null ? null : new[] { value }, width, height, align, fontSize, spacingBefore, spacingAfter, isComboBox, allowsMultipleSelection: false, style) {
    }

    public ChoiceFieldBlock(string name, IEnumerable<string> options, IEnumerable<string>? values, double width, double height, PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, bool isComboBox, bool allowsMultipleSelection, PdfFormFieldStyle? style = null) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        Guard.NotNull(options, nameof(options));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.LeftCenterRightAlign(align, nameof(align), "Choice field");
        Guard.Positive(fontSize, nameof(fontSize));
        Guard.NonNegative(spacingBefore, nameof(spacingBefore));
        Guard.NonNegative(spacingAfter, nameof(spacingAfter));
        if (allowsMultipleSelection && isComboBox) {
            throw new ArgumentException("PDF multi-select choice fields must be list boxes, not combo boxes.", nameof(isComboBox));
        }

        var normalizedOptions = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (string? option in options) {
            Guard.NotNullOrWhiteSpace(option, nameof(options));
            if (!seen.Add(option!)) {
                throw new ArgumentException("PDF choice field options must be unique.", nameof(options));
            }

            normalizedOptions.Add(option!);
        }

        if (normalizedOptions.Count == 0) {
            throw new ArgumentException("PDF choice field requires at least one option.", nameof(options));
        }

        List<string> selectedValues = NormalizeSelectedValues(values, normalizedOptions, seen, allowsMultipleSelection);
        if (!allowsMultipleSelection && selectedValues.Count != 1) {
            throw new ArgumentException("PDF scalar choice field must have exactly one selected value.", nameof(values));
        }

        Name = name;
        Options = normalizedOptions.AsReadOnly();
        Values = selectedValues.AsReadOnly();
        Value = Values[0];
        Width = width;
        Height = height;
        Align = align;
        FontSize = fontSize;
        SpacingBefore = spacingBefore;
        SpacingAfter = spacingAfter;
        IsComboBox = isComboBox;
        AllowsMultipleSelection = allowsMultipleSelection;
        Style = style?.Clone() ?? new PdfFormFieldStyle();
    }

    private static List<string> NormalizeSelectedValues(IEnumerable<string>? values, List<string> options, HashSet<string> optionSet, bool allowsMultipleSelection) {
        if (values is null) {
            return new List<string> { options[0] };
        }

        var selected = new List<string>();
        var selectedSet = new HashSet<string>(StringComparer.Ordinal);
        foreach (string? value in values) {
            Guard.NotNullOrWhiteSpace(value, nameof(values));
            if (!optionSet.Contains(value!)) {
                throw new ArgumentException("PDF choice field values must match the provided options.", nameof(values));
            }

            if (!selectedSet.Add(value!)) {
                throw new ArgumentException("PDF choice field selected values must be unique.", nameof(values));
            }

            selected.Add(value!);
        }

        if (selected.Count == 0) {
            throw new ArgumentException("PDF choice field requires at least one selected value.", nameof(values));
        }

        if (!allowsMultipleSelection && selected.Count > 1) {
            throw new ArgumentException("PDF scalar choice field cannot contain multiple selected values.", nameof(values));
        }

        return selected;
    }
}
