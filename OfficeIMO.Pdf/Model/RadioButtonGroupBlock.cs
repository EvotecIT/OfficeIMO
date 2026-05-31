namespace OfficeIMO.Pdf;

internal sealed class RadioButtonGroupBlock : IPdfBlock {
    public string Name { get; }
    public IReadOnlyList<string> Options { get; }
    public string Value { get; }
    public double Size { get; }
    public double Gap { get; }
    public PdfAlign Align { get; }
    public double SpacingBefore { get; }
    public double SpacingAfter { get; }
    public PdfFormFieldStyle Style { get; }
    public double Height => Options.Count * Size + Math.Max(0, Options.Count - 1) * Gap;

    public RadioButtonGroupBlock(string name, IEnumerable<string> options, string? value, double size, double gap, PdfAlign align, double spacingBefore, double spacingAfter, PdfFormFieldStyle? style = null) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        Guard.NotNull(options, nameof(options));
        Guard.Positive(size, nameof(size));
        Guard.NonNegative(gap, nameof(gap));
        Guard.LeftCenterRightAlign(align, nameof(align), "Radio button group");
        Guard.NonNegative(spacingBefore, nameof(spacingBefore));
        Guard.NonNegative(spacingAfter, nameof(spacingAfter));

        var normalizedOptions = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (string? option in options) {
            Guard.NotNullOrWhiteSpace(option, nameof(options));
            if (string.Equals(option, "Off", StringComparison.Ordinal)) {
                throw new ArgumentException("PDF radio button option value cannot be Off.", nameof(options));
            }

            ValidateAsciiPdfNameValue(option!, nameof(options));
            if (!seen.Add(option!)) {
                throw new ArgumentException("PDF radio button options must be unique.", nameof(options));
            }

            normalizedOptions.Add(option!);
        }

        if (normalizedOptions.Count == 0) {
            throw new ArgumentException("PDF radio button group requires at least one option.", nameof(options));
        }

        string selectedValue = value ?? normalizedOptions[0];
        Guard.NotNullOrWhiteSpace(selectedValue, nameof(value));
        if (!seen.Contains(selectedValue)) {
            throw new ArgumentException("PDF radio button value must match the provided options.", nameof(value));
        }

        Name = name;
        Options = normalizedOptions.AsReadOnly();
        Value = selectedValue;
        Size = size;
        Gap = gap;
        Align = align;
        SpacingBefore = spacingBefore;
        SpacingAfter = spacingAfter;
        Style = style?.Clone() ?? new PdfFormFieldStyle();
    }

    private static void ValidateAsciiPdfNameValue(string value, string paramName) {
        for (int i = 0; i < value.Length; i++) {
            if (value[i] > 0x7E) {
                throw new ArgumentException("PDF radio button option values must contain only ASCII PDF name characters.", paramName);
            }
        }
    }
}
