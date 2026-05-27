namespace OfficeIMO.Pdf;

internal sealed class ChoiceFieldBlock : IPdfBlock {
    public string Name { get; }
    public IReadOnlyList<string> Options { get; }
    public string Value { get; }
    public double Width { get; }
    public double Height { get; }
    public PdfAlign Align { get; }
    public double FontSize { get; }
    public double SpacingBefore { get; }
    public double SpacingAfter { get; }
    public bool IsComboBox { get; }

    public ChoiceFieldBlock(string name, IEnumerable<string> options, string? value, double width, double height, PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, bool isComboBox) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        Guard.NotNull(options, nameof(options));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.LeftCenterRightAlign(align, nameof(align), "Choice field");
        Guard.Positive(fontSize, nameof(fontSize));
        Guard.NonNegative(spacingBefore, nameof(spacingBefore));
        Guard.NonNegative(spacingAfter, nameof(spacingAfter));

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

        string selectedValue = value ?? normalizedOptions[0];
        if (!seen.Contains(selectedValue)) {
            throw new ArgumentException("PDF choice field value must match one of the provided options.", nameof(value));
        }

        Name = name;
        Options = normalizedOptions.AsReadOnly();
        Value = selectedValue;
        Width = width;
        Height = height;
        Align = align;
        FontSize = fontSize;
        SpacingBefore = spacingBefore;
        SpacingAfter = spacingAfter;
        IsComboBox = isComboBox;
    }
}
