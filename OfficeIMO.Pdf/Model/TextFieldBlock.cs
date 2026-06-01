namespace OfficeIMO.Pdf;

internal sealed class TextFieldBlock : IPdfBlock {
    public string Name { get; }
    public double Width { get; }
    public double Height { get; }
    public string Value { get; }
    public PdfAlign Align { get; }
    public double FontSize { get; }
    public double SpacingBefore { get; }
    public double SpacingAfter { get; }
    public PdfFormFieldStyle Style { get; }

    public TextFieldBlock(string name, double width, double height, string? value, PdfAlign align, double fontSize, double spacingBefore, double spacingAfter, PdfFormFieldStyle? style = null) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.LeftCenterRightAlign(align, nameof(align), "Text field");
        Guard.Positive(fontSize, nameof(fontSize));
        Guard.NonNegative(spacingBefore, nameof(spacingBefore));
        Guard.NonNegative(spacingAfter, nameof(spacingAfter));

        Name = name;
        Width = width;
        Height = height;
        Value = value ?? string.Empty;
        Align = align;
        FontSize = fontSize;
        SpacingBefore = spacingBefore;
        SpacingAfter = spacingAfter;
        Style = style?.Clone() ?? new PdfFormFieldStyle();
    }
}
