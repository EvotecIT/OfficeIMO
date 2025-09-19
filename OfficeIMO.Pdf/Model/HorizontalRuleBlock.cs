namespace OfficeIMO.Pdf;

internal sealed class HorizontalRuleBlock : IPdfBlock {
    public double Thickness { get; }
    public PdfColor Color { get; }
    public double SpacingBefore { get; }
    public double SpacingAfter { get; }
    public HorizontalRuleBlock(double thickness, PdfColor color, double spacingBefore, double spacingAfter) {
        Thickness = thickness; Color = color; SpacingBefore = spacingBefore; SpacingAfter = spacingAfter;
    }
}

