namespace OfficeIMO.Pdf;

internal sealed class HorizontalRuleBlock : IPdfBlock {
    public PdfHorizontalRuleStyle? Style { get; }

    public HorizontalRuleBlock(PdfHorizontalRuleStyle? style = null) {
        Style = style?.Clone();
    }

    public HorizontalRuleBlock(double thickness, PdfColor color, double spacingBefore, double spacingAfter)
        : this(new PdfHorizontalRuleStyle {
            Thickness = thickness,
            Color = color,
            SpacingBefore = spacingBefore,
            SpacingAfter = spacingAfter
        }) {
    }
}
