namespace OfficeIMO.Pdf;

internal sealed class CheckBoxBlock : IPdfBlock {
    public string Name { get; }
    public bool IsChecked { get; }
    public double Size { get; }
    public OfficeAlignment Align { get; }
    public double SpacingBefore { get; }
    public double SpacingAfter { get; }
    public string CheckedValueName { get; }
    public PdfFormFieldStyle Style { get; }

    public CheckBoxBlock(string name, bool isChecked, double size, OfficeAlignment align, double spacingBefore, double spacingAfter, string checkedValueName, PdfFormFieldStyle? style = null) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        Guard.Positive(size, nameof(size));
        align = Guard.LeftCenterRightAlign(align, nameof(align), "Check box");
        Guard.NonNegative(spacingBefore, nameof(spacingBefore));
        Guard.NonNegative(spacingAfter, nameof(spacingAfter));
        Guard.NotNullOrWhiteSpace(checkedValueName, nameof(checkedValueName));
        if (string.Equals(checkedValueName, "Off", StringComparison.Ordinal)) {
            throw new ArgumentException("Check box selected value name cannot be Off.", nameof(checkedValueName));
        }

        Name = name;
        IsChecked = isChecked;
        Size = size;
        Align = align;
        SpacingBefore = spacingBefore;
        SpacingAfter = spacingAfter;
        CheckedValueName = checkedValueName;
        Style = style?.Clone() ?? new PdfFormFieldStyle();
    }
}
