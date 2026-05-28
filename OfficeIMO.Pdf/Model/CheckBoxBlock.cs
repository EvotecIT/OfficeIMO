namespace OfficeIMO.Pdf;

internal sealed class CheckBoxBlock : IPdfBlock {
    public string Name { get; }
    public bool IsChecked { get; }
    public double Size { get; }
    public PdfAlign Align { get; }
    public double SpacingBefore { get; }
    public double SpacingAfter { get; }
    public string CheckedValueName { get; }

    public CheckBoxBlock(string name, bool isChecked, double size, PdfAlign align, double spacingBefore, double spacingAfter, string checkedValueName) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        Guard.Positive(size, nameof(size));
        Guard.LeftCenterRightAlign(align, nameof(align), "Check box");
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
    }
}
